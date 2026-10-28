using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace ClosedXML.Excel
{
    /// <summary>
    /// Imports XML data into the cells and tables of an XML map.
    /// </summary>
    internal static class XmlMapImporter
    {
        private const string XsiNamespace = "http://www.w3.org/2001/XMLSchema-instance";

        private static readonly Regex NamespaceDeclaration = new(@"xmlns(?::(?<prefix>[^\s=]+))?\s*=\s*(?<quote>['""])(?<uri>.*?)\k<quote>", RegexOptions.Compiled);

        private static readonly XmlReaderSettings ReaderSettings = new()
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
        };

        internal static XLXmlImportResult Import(XLXmlMap map, Stream stream, bool overwrite)
        {
            using var reader = XmlReader.Create(stream, ReaderSettings);
            return Import(map, reader, overwrite);
        }

        internal static XLXmlImportResult Import(XLXmlMap map, TextReader textReader, bool overwrite)
        {
            using var reader = XmlReader.Create(textReader, ReaderSettings);
            return Import(map, reader, overwrite);
        }

        private static XLXmlImportResult Import(XLXmlMap map, XmlReader reader, bool overwrite)
        {
            var document = new XmlDocument { XmlResolver = null, PreserveWhitespace = true };
            document.Load(reader);

            var schema = XmlMapSchema.Create(map);
            var valid = schema.Validate(document);
            var namespaces = CreateNamespaceManager(document.NameTable, map.SelectionNamespaces);
            var context = new ImportContext(map, document, namespaces, schema);

            // Single cells aren't changed when the data are appended, appending requires an XML table.
            var append = !overwrite || map.AppendOnImport;
            if (append && !map.Store.GetTableBindings(map).Any())
                throw new InvalidOperationException("Import error. No data was imported: data can only be appended to XML tables.");

            // Nothing is changed when the data can't be imported into any of the tables.
            var cellBindings = map.Store.GetCellBindings(map).ToList();
            var tables = map.Store.GetTableBindings(map).ToList()
                .Select(t => PlanTable(context, t.Table, t.Binding, append))
                .OfType<TablePlan>()
                .ToList();
            CheckCanImport(cellBindings, tables);

            if (!append)
            {
                foreach (var binding in cellBindings)
                    ImportCell(context, binding);
            }

            foreach (var plan in tables)
                ImportTable(context, plan);

            if (map.AdjustColumnWidth)
                AdjustColumnWidths(context);

            if (!valid)
                return XLXmlImportResult.ValidationFailed;

            return context.Truncated ? XLXmlImportResult.ElementsTruncated : XLXmlImportResult.Success;
        }

        private static void ImportCell(ImportContext context, XLXmlCellBinding binding)
        {
            var cell = binding.Worksheet.Cell(binding.Cell.RangeAddress.FirstAddress.RowNumber, binding.Cell.RangeAddress.FirstAddress.ColumnNumber);
            var node = SelectNode(context, context.Document, binding.XPath);
            SetValue(cell, context.Convert(node, binding.XmlDataType), binding.XmlDataType);
            context.AdjustedColumns.Add((cell.Worksheet, cell.Address.ColumnNumber));
        }

        /// <summary>
        /// Read the rows of a table from the XML data without changing the worksheet.
        /// </summary>
        /// <returns>The plan or <c>null</c>, if the map has no columns in the table.</returns>
        private static TablePlan? PlanTable(ImportContext context, XLTable table, XLXmlTableBinding binding, bool append)
        {
            var fieldNames = table.Fields.OrderBy(f => f.Index).Select(f => f.Name).ToList();
            var columns = binding.Columns
                .Where(c => c.Map == context.Map)
                .Select(c => (Index: c.ResolveIndex(fieldNames), Binding: c, Steps: XPathStep.Parse(c.XPath)))
                .Where(c => c.Index >= 0 && c.Steps.Count > 0)
                .ToList();
            if (columns.Count == 0)
                return null;

            var levels = GetRowLevels(context.Schema, context.Namespaces, columns.Select(c => c.Steps).ToList());
            var rows = new List<List<XLCellValue>>();
            if (levels.Count > 0)
            {
                var flattener = new RowFlattener(context, levels, columns.Select(c => (c.Steps, c.Binding.XmlDataType)).ToList());
                flattener.Collect(context.Document, 0, new XLCellValue[columns.Count], rows);
            }

            // Repeating elements without any mapped data don't become rows.
            rows = rows.Where(values => values.Any(v => !v.IsBlank)).ToList();
            return new TablePlan(table, binding, columns.Select(c => (c.Index, c.Binding)).ToList(), rows, append);
        }

        /// <summary>
        /// An import is refused without any change, when a sheet with mapped cells or tables of the map is protected or
        /// a table can't grow. A table grows into empty rows below it, otherwise the cells below it are shifted down in the
        /// columns of the table, which fails when it would split other tables, merged cells or array formulas or shift cells off the sheet.
        /// </summary>
        private static void CheckCanImport(List<XLXmlCellBinding> cellBindings, List<TablePlan> tables)
        {
            var sheets = cellBindings.Select(b => b.Worksheet).Concat(tables.Select(t => t.Table.Worksheet));
            if (sheets.Any(ws => ws.IsProtected))
                throw new InvalidOperationException("Import error. The XML data can't be imported, because a mapped worksheet is protected.");

            foreach (var plan in tables)
            {
                if (plan.AddedRows <= 0)
                    continue;

                var table = plan.Table;
                var ws = table.Worksheet;
                var lastTableRow = table.RangeAddress.LastAddress.RowNumber;
                if (lastTableRow + plan.AddedRows > XLHelper.MaxRowNumber)
                    throw CannotResize();

                var shiftRow = GetFirstRowInTheWay(table, plan.AddedRows);
                if (shiftRow is null)
                    continue;

                var firstColumn = plan.FirstColumn;
                var lastColumn = plan.LastColumn;
                var shift = lastTableRow + plan.AddedRows - shiftRow.Value + 1;
                var lastUsedRow = ws.Range(shiftRow.Value, firstColumn, XLHelper.MaxRowNumber, lastColumn).LastRowUsed(XLCellsUsedOptions.Contents)?.RowNumber() ?? 0;
                if (lastUsedRow + shift > XLHelper.MaxRowNumber)
                    throw CannotResize();

                // Areas in the shifted cells must be completely inside of the columns of the table.
                bool Splits(IXLRangeAddress area) =>
                    area.LastAddress.RowNumber >= shiftRow.Value &&
                    area.LastAddress.ColumnNumber >= firstColumn && area.FirstAddress.ColumnNumber <= lastColumn &&
                    (area.FirstAddress.ColumnNumber < firstColumn || area.LastAddress.ColumnNumber > lastColumn);

                var arrayFormulas = ws.Range(shiftRow.Value, firstColumn, XLHelper.MaxRowNumber, lastColumn)
                    .CellsUsed(XLCellsUsedOptions.Contents, c => c.HasArrayFormula)
                    .Select(c => c.FormulaReference);
                if (ws.Tables.OfType<IXLTable>().Where(t => t != table).Select(t => t.RangeAddress)
                    .Concat(ws.MergedRanges.Select(r => r.RangeAddress))
                    .Concat(arrayFormulas)
                    .Any(Splits))
                {
                    throw CannotResize();
                }
            }
        }

        private static InvalidOperationException CannotResize() => new("Import error. An XML table can't be resized to accommodate the data.");

        /// <summary>
        /// Find the first row below the table that isn't empty in the columns of the table, within the rows the table grows into.
        /// Merged cells are in the way even when they are empty.
        /// </summary>
        private static int? GetFirstRowInTheWay(XLTable table, int count)
        {
            var ws = table.Worksheet;
            var firstColumn = table.RangeAddress.FirstAddress.ColumnNumber;
            var lastColumn = table.RangeAddress.LastAddress.ColumnNumber;
            var lastTableRow = table.RangeAddress.LastAddress.RowNumber;
            for (var row = lastTableRow + 1; row <= Math.Min(lastTableRow + count, XLHelper.MaxRowNumber); row++)
            {
                if (!IsRowEmpty(ws, row, firstColumn, lastColumn) ||
                    ws.MergedRanges.Any(r => r.RangeAddress.FirstAddress.RowNumber <= row && r.RangeAddress.LastAddress.RowNumber >= row &&
                                             r.RangeAddress.FirstAddress.ColumnNumber <= lastColumn && r.RangeAddress.LastAddress.ColumnNumber >= firstColumn))
                {
                    return row;
                }
            }

            return null;
        }

        private static void ImportTable(ImportContext context, TablePlan plan)
        {
            var (table, binding, columns, rows, append) = (plan.Table, plan.Binding, plan.Columns, plan.Rows, plan.Append);
            var ws = table.Worksheet;
            var firstColumn = plan.FirstColumn;
            var lastColumn = plan.LastColumn;
            var firstDataRow = plan.FirstDataRow;
            var lastDataRow = plan.LastDataRow;
            var dataRowCount = plan.DataRowCount;

            foreach (var column in columns)
            {
                if (column.Binding.DataStyle is null && dataRowCount > 0)
                    column.Binding.DataStyle = ((XLCell)ws.Cell(firstDataRow, firstColumn + column.Index)).StyleValue;
            }

            if (plan.TargetRowCount is not { } targetRowCount)
                return;

            // The insert row is shown only for a table that had data and was emptied by the import.
            var hasData = binding.InsertRow != true && columns.Any(c => !IsColumnEmpty(ws, firstDataRow, lastDataRow, firstColumn + c.Index));
            var existingRows = plan.ExistingRows;

            // Rows hidden by the filter of the table are shown, the filter is applied again after the import.
            var filtered = table.ShowAutoFilter && table.AutoFilter.Columns.Count > 0;
            if (filtered)
                ws.Rows(firstDataRow, lastDataRow).Unhide();

            // Added rows get the formulas of calculated columns: columns that aren't mapped and have the same formula in all data rows.
            var mappedIndexes = new HashSet<int>(columns.Select(c => c.Index));
            var calculatedColumns = Enumerable.Range(0, lastColumn - firstColumn + 1)
                .Where(i => !mappedIndexes.Contains(i))
                .Select(i => (Column: firstColumn + i, Formula: GetCalculatedFormula(ws, firstDataRow, lastDataRow, firstColumn + i)))
                .Where(c => c.Formula is not null)
                .ToList();
            if (targetRowCount > dataRowCount)
            {
                AddRows(table, targetRowCount - dataRowCount, lastDataRow, firstColumn, lastColumn);
            }
            else if (targetRowCount < dataRowCount)
            {
                RemoveRows(table, dataRowCount - targetRowCount, firstDataRow + targetRowCount, lastDataRow, firstColumn, lastColumn);
            }

            // Fields cache their column range, which doesn't follow a resize of the table (used e.g. to find calculated columns when saving).
            foreach (var field in table.Fields.Cast<XLTableField>())
                field.Column = null!;

            if (rows.Count == 0 && !append)
            {
                // The single remaining row of an empty table is cleared including the formats.
                ws.Range(firstDataRow, firstColumn, firstDataRow, lastColumn).Clear(XLClearOptions.Contents | XLClearOptions.NormalFormats);
            }

            for (var i = 0; i < rows.Count; i++)
            {
                var rowNumber = firstDataRow + existingRows + i;
                for (var c = 0; c < columns.Count; c++)
                {
                    var cell = (XLCell)ws.Cell(rowNumber, firstColumn + columns[c].Index);
                    var dataStyle = columns[c].Binding.DataStyle;
                    if (dataStyle is not null && (rowNumber > lastDataRow || cell.StyleValue.Equals(ws.StyleValue)))
                        ((IXLStylized)cell).ModifyStyle(_ => dataStyle.Key);

                    SetValue(cell, rows[i][c], columns[c].Binding.XmlDataType);
                }
            }

            for (var row = lastDataRow + 1; row < firstDataRow + targetRowCount; row++)
            {
                foreach (var (column, formula) in calculatedColumns)
                    ws.Cell(row, column).FormulaR1C1 = formula;
            }

            // The filter of the table is applied to the imported rows, but they aren't sorted.
            if (filtered)
                table.AutoFilter.Reapply();

            if (existingRows + rows.Count > 0)
                binding.InsertRow = null;
            else if (hasData)
                binding.InsertRow = true;
            foreach (var column in columns)
                context.AdjustedColumns.Add((ws, firstColumn + column.Index));
        }

        /// <summary>
        /// Add rows to the end of a table. Cells below the table are only moved down
        /// when they are in the way, and only in the columns of the table.
        /// </summary>
        private static void AddRows(XLTable table, int count, int lastDataRow, int firstColumn, int lastColumn)
        {
            var ws = table.Worksheet;
            var lastTableRow = table.RangeAddress.LastAddress.RowNumber;
            var shiftRow = GetFirstRowInTheWay(table, count);
            var shift = shiftRow is null ? 0 : lastTableRow + count - shiftRow.Value + 1;
            KeepArrayFormulas(ws, lastTableRow + 1, firstColumn, lastColumn, row => row >= shiftRow ? row + shift : row, () =>
            {
                if (shiftRow is not null)
                    ws.Range(shiftRow.Value, firstColumn, shiftRow.Value, lastColumn).InsertRowsAbove(shift);

                if (table.ShowTotalsRow)
                {
                    // Move the totals row down into the now empty rows: insert rows above it (the table grows)
                    // and remove the same number of empty rows below it, so cells below the table stay in place.
                    var totalsRow = lastDataRow + 1;
                    ws.Range(totalsRow, firstColumn, totalsRow, lastColumn).InsertRowsAbove(count);
                    ws.Range(totalsRow + count + 1, firstColumn, totalsRow + 2 * count, lastColumn).Delete(XLShiftDeletedCells.ShiftCellsUp);
                }
            });

            if (table.ShowTotalsRow)
                return;

            var first = table.RangeAddress.FirstAddress;
            table.Resize(first.RowNumber, firstColumn, lastTableRow + count, lastColumn);
        }

        /// <summary>
        /// ClosedXML turns array formulas into normal formulas when it shifts their cells, but they must be kept.
        /// Restore the array formulas between <paramref name="firstColumn"/> and <paramref name="lastColumn"/> at or below
        /// <paramref name="firstRow"/> after <paramref name="shiftCells"/>, <paramref name="newRow"/> gives their new row.
        /// </summary>
        private static void KeepArrayFormulas(XLWorksheet ws, int firstRow, int firstColumn, int lastColumn, Func<int, int> newRow, Action shiftCells)
        {
            var arrays = firstRow > XLHelper.MaxRowNumber
                ? new List<IXLRangeAddress>()
                : ws.Range(firstRow, firstColumn, XLHelper.MaxRowNumber, lastColumn)
                    .CellsUsed(XLCellsUsedOptions.Contents, c => c.HasArrayFormula)
                    .Select(c => c.FormulaReference)
                    .GroupBy(a => a.ToString())
                    .Select(g => g.First())
                    .ToList();
            shiftCells();

            foreach (var address in arrays)
            {
                var offset = newRow(address.FirstAddress.RowNumber) - address.FirstAddress.RowNumber;
                var range = ws.Range(address.FirstAddress.RowNumber + offset, address.FirstAddress.ColumnNumber, address.LastAddress.RowNumber + offset, address.LastAddress.ColumnNumber);
                var first = range.FirstCell();
                if (!first.HasArrayFormula && first.HasFormula)
                    range.FormulaArrayA1 = first.FormulaA1;
            }
        }

        /// <summary>
        /// Remove rows from the end of a table. The rows are cleared and cells below the table stay where they are.
        /// </summary>
        private static void RemoveRows(XLTable table, int count, int firstRemovedRow, int lastDataRow, int firstColumn, int lastColumn)
        {
            var ws = table.Worksheet;
            var removed = ws.Range(firstRemovedRow, firstColumn, lastDataRow, lastColumn);
            if (table.ShowTotalsRow)
            {
                // Move the totals row up (the table shrinks) and insert empty rows below it, so cells below the table stay in place.
                KeepArrayFormulas(ws, lastDataRow + 2, firstColumn, lastColumn, row => row, () =>
                {
                    removed.Delete(XLShiftDeletedCells.ShiftCellsUp);
                    var totalsRow = firstRemovedRow;
                    var inserted = ws.Range(totalsRow + 1, firstColumn, totalsRow + 1, lastColumn).InsertRowsAbove(count);
                    inserted.Clear(XLClearOptions.Contents | XLClearOptions.NormalFormats);
                });
                return;
            }

            removed.Clear(XLClearOptions.Contents | XLClearOptions.NormalFormats);
            var first = table.RangeAddress.FirstAddress;
            table.Resize(first.RowNumber, firstColumn, table.RangeAddress.LastAddress.RowNumber - count, lastColumn);
        }

        /// <summary>
        /// Find the repeating elements that form the rows of an XML table, from the outermost to the innermost.
        /// The innermost is the deepest repeating element of any column. Columns mapped to an element above
        /// it (denormalized columns) repeat their value in each row of the element.
        /// </summary>
        internal static List<List<XPathStep>> GetRowLevels(XmlMapSchema schema, XmlNamespaceManager namespaces, List<List<XPathStep>> columnSteps)
        {
            var deepest = new List<XPathStep>();
            foreach (var path in columnSteps.Select(s => s.Where(step => !step.IsAttribute).ToList()))
            {
                for (var length = path.Count; length > deepest.Count; length--)
                {
                    if (IsRepeating(schema, namespaces, path, length))
                    {
                        deepest = path.Take(length).ToList();
                        break;
                    }
                }
            }

            if (deepest.Count == 0)
            {
                var rowSteps = GetRowSteps(schema, namespaces, columnSteps);
                return rowSteps.Count > 0 ? new List<List<XPathStep>> { rowSteps } : new List<List<XPathStep>>();
            }

            return Enumerable.Range(1, deepest.Count)
                .Where(length => IsRepeating(schema, namespaces, deepest, length))
                .Select(length => deepest.Take(length).ToList())
                .ToList();
        }

        private static bool IsRepeating(XmlMapSchema schema, XmlNamespaceManager namespaces, List<XPathStep> path, int length)
        {
            return schema.IsRepeating(path.Take(length).Select(s => s.ToQualifiedName(namespaces)).ToList()) == true;
        }

        /// <summary>
        /// Find the path of the repeating element that forms a row of an XML table, when the schema doesn't know the elements.
        /// </summary>
        private static List<XPathStep> GetRowSteps(XmlMapSchema schema, XmlNamespaceManager namespaces, List<List<XPathStep>> columnSteps)
        {
            var elementPaths = columnSteps.Select(s => s.Where(step => !step.IsAttribute).ToList()).ToList();
            var commonLength = elementPaths.Min(p => p.Count);
            for (var i = 0; i < commonLength; i++)
            {
                if (elementPaths.Any(p => !p[i].Equals(elementPaths[0][i])))
                {
                    commonLength = i;
                    break;
                }
            }

            var common = elementPaths[0].Take(commonLength).ToList();
            for (var length = common.Count; length > 0; length--)
            {
                var names = common.Take(length).Select(s => s.ToQualifiedName(namespaces)).ToList();
                if (schema.IsRepeating(names) == true)
                    return common.Take(length).ToList();
            }

            // Without schema information, a single column of elements is the child of the row element.
            var allColumnsAreTheCommonElement = elementPaths.All(p => p.Count == common.Count) && columnSteps.Any(s => !s.Last().IsAttribute);
            return allColumnsAreTheCommonElement && common.Count > 1 ? common.Take(common.Count - 1).ToList() : common;
        }

        private static void AdjustColumnWidths(ImportContext context)
        {
            foreach (var (worksheet, column) in context.AdjustedColumns)
                worksheet.Column(column).AdjustToContents();
        }

        private static bool IsRowEmpty(IXLWorksheet ws, int row, int firstColumn, int lastColumn)
        {
            for (var column = firstColumn; column <= lastColumn; column++)
            {
                if (!ws.Cell(row, column).IsEmpty(XLCellsUsedOptions.Contents))
                    return false;
            }

            return true;
        }

        /// <returns>The R1C1 formula all cells of the column have, or <c>null</c>.</returns>
        private static string? GetCalculatedFormula(IXLWorksheet ws, int firstRow, int lastRow, int column)
        {
            string? formula = null;
            for (var row = firstRow; row <= lastRow; row++)
            {
                var cell = ws.Cell(row, column);
                if (!cell.HasFormula || (formula is not null && formula != cell.FormulaR1C1))
                    return null;

                formula = cell.FormulaR1C1;
            }

            return formula;
        }

        private static bool IsColumnEmpty(IXLWorksheet ws, int firstRow, int lastRow, int column)
        {
            for (var row = firstRow; row <= lastRow; row++)
            {
                if (!ws.Cell(row, column).IsEmpty(XLCellsUsedOptions.Contents))
                    return false;
            }

            return true;
        }

        private static XmlNode? SelectNode(ImportContext context, XmlNode node, string xpath)
        {
            return string.IsNullOrEmpty(xpath) ? null : node.SelectSingleNode(xpath, context.Namespaces);
        }

        /// <summary>
        /// Set the value without changing the style. A cell with General format gets the default
        /// format of the XML data type, when a value is imported.
        /// </summary>
        private static void SetValue(IXLCell cell, XLCellValue value, string? xmlDataType)
        {
            var xlCell = (XLCell)cell;
            if (xlCell.HasFormula)
                xlCell.FormulaA1 = string.Empty;

            xlCell.SetOnlyValue(value);
            var numberFormat = xlCell.Style.NumberFormat;
            var currentFormatId = string.IsNullOrEmpty(numberFormat.Format) ? numberFormat.NumberFormatId : (int?)null;
            if (XmlDataTypeConverter.GetNumberFormatId(xmlDataType, value, currentFormatId) is { } numberFormatId)
                numberFormat.NumberFormatId = numberFormatId;
        }

        internal static XmlNamespaceManager CreateNamespaceManager(XmlNameTable nameTable, string selectionNamespaces)
        {
            var namespaces = new XmlNamespaceManager(nameTable);
            foreach (Match match in NamespaceDeclaration.Matches(selectionNamespaces))
            {
                var prefix = match.Groups["prefix"].Success ? match.Groups["prefix"].Value : string.Empty;
                namespaces.AddNamespace(prefix, match.Groups["uri"].Value);
            }

            return namespaces;
        }

        private sealed class ImportContext
        {
            internal ImportContext(XLXmlMap map, XmlDocument document, XmlNamespaceManager namespaces, XmlMapSchema schema)
            {
                Map = map;
                Document = document;
                Namespaces = namespaces;
                Schema = schema;
                Use1904DateSystem = map.Store.Workbook.Use1904DateSystem;
            }

            internal XLXmlMap Map { get; }

            internal XmlDocument Document { get; }

            internal XmlNamespaceManager Namespaces { get; }

            internal XmlMapSchema Schema { get; }

            internal bool Use1904DateSystem { get; }

            internal bool Truncated { get; private set; }

            internal HashSet<(XLWorksheet Worksheet, int Column)> AdjustedColumns { get; } = new();

            internal XLCellValue Convert(XmlNode? node, string? xmlDataType)
            {
                // The schema validation adds attributes with default values to the document, the defaults are ignored.
                if (node is null || node is XmlAttribute { Specified: false } || IsNil(node))
                    return Blank.Value;

                var value = XmlDataTypeConverter.Convert(GetText(node), xmlDataType, Use1904DateSystem, out var truncated);
                Truncated |= truncated;
                return value;
            }

            /// <summary>
            /// Get the text of an element (including the text of child elements) or attribute. Leading and trailing
            /// whitespace is removed, but whitespace inside of CDATA sections is kept.
            /// </summary>
            private static string GetText(XmlNode node)
            {
                if (node is XmlAttribute attribute)
                    return attribute.Value.Trim(XmlDataTypeConverter.XmlWhitespace);

                var text = new StringBuilder();
                var cdataStart = -1;
                var cdataEnd = -1;
                AppendText(node, text, ref cdataStart, ref cdataEnd);

                var start = 0;
                var end = text.Length;
                while (start < end && (cdataStart < 0 || start < cdataStart) && Array.IndexOf(XmlDataTypeConverter.XmlWhitespace, text[start]) >= 0)
                    start++;
                while (end > start && (cdataEnd < 0 || end > cdataEnd) && Array.IndexOf(XmlDataTypeConverter.XmlWhitespace, text[end - 1]) >= 0)
                    end--;

                return text.ToString(start, end - start);
            }

            private static void AppendText(XmlNode node, StringBuilder text, ref int cdataStart, ref int cdataEnd)
            {
                foreach (XmlNode child in node.ChildNodes)
                {
                    switch (child)
                    {
                        case XmlCDataSection cdata:
                            if (cdataStart < 0)
                                cdataStart = text.Length;
                            text.Append(cdata.Value);
                            cdataEnd = text.Length;
                            break;
                        case XmlCharacterData data when data is XmlText or XmlWhitespace or XmlSignificantWhitespace:
                            text.Append(data.Value);
                            break;
                        case XmlElement element:
                            AppendText(element, text, ref cdataStart, ref cdataEnd);
                            break;
                    }
                }
            }

            private static bool IsNil(XmlNode node)
            {
                return node is XmlElement element && element.GetAttribute("nil", XsiNamespace) is "true" or "1";
            }
        }

        /// <summary>
        /// The rows of an XML table read from the XML data and how the table must change for them.
        /// </summary>
        private sealed class TablePlan
        {
            internal TablePlan(XLTable table, XLXmlTableBinding binding, List<(int Index, XLXmlColumnBinding Binding)> columns, List<List<XLCellValue>> rows, bool append)
            {
                Table = table;
                Binding = binding;
                Columns = columns;
                Rows = rows;
                Append = append;
                FirstColumn = table.RangeAddress.FirstAddress.ColumnNumber;
                LastColumn = table.RangeAddress.LastAddress.ColumnNumber;
                FirstDataRow = table.RangeAddress.FirstAddress.RowNumber + (table.ShowHeaderRow ? 1 : 0);
                LastDataRow = table.RangeAddress.LastAddress.RowNumber - (table.ShowTotalsRow ? 1 : 0);

                // The insert row of a table is shown when the table has no data, it doesn't count as a data row.
                ExistingRows = append && binding.InsertRow != true ? DataRowCount : 0;
                TargetRowCount = rows.Count == 0 && append ? null : Math.Max(1, ExistingRows + rows.Count);
            }

            internal XLTable Table { get; }

            internal XLXmlTableBinding Binding { get; }

            internal List<(int Index, XLXmlColumnBinding Binding)> Columns { get; }

            internal List<List<XLCellValue>> Rows { get; }

            internal bool Append { get; }

            internal int FirstColumn { get; }

            internal int LastColumn { get; }

            internal int FirstDataRow { get; }

            internal int LastDataRow { get; }

            internal int DataRowCount => LastDataRow - FirstDataRow + 1;

            /// <summary>
            /// Data rows that are kept when appending.
            /// </summary>
            internal int ExistingRows { get; }

            /// <summary>
            /// Number of data rows after the import, <c>null</c> if the table isn't changed.
            /// </summary>
            internal int? TargetRowCount { get; }

            internal int AddedRows => (TargetRowCount ?? DataRowCount) - DataRowCount;
        }

        /// <summary>
        /// Flattens nested repeating elements into table rows: each element of the innermost level is a row,
        /// values of outer levels are repeated in all rows of their element and an element without any element of the
        /// next level still gives one row (like an outer join).
        /// </summary>
        private sealed class RowFlattener
        {
            private readonly ImportContext _context;
            private readonly List<string> _levelPaths;
            private readonly List<(int Level, string Path, string? XmlDataType)> _columns;

            internal RowFlattener(ImportContext context, List<List<XPathStep>> levels, List<(List<XPathStep> Steps, string? XmlDataType)> columns)
            {
                _context = context;
                _levelPaths = levels.Select((level, i) => XPathStep.Format(level.Skip(i == 0 ? 0 : levels[i - 1].Count).ToList(), absolute: false)).ToList();
                _columns = columns.Select(column =>
                {
                    // The column belongs to the deepest level its path starts with, -1 is the document.
                    var level = levels.FindLastIndex(l => l.Count <= column.Steps.Count && l.SequenceEqual(column.Steps.Take(l.Count)));
                    var skip = level < 0 ? 0 : levels[level].Count;
                    return (level, XPathStep.Format(column.Steps.Skip(skip).ToList(), absolute: false), column.XmlDataType);
                }).ToList();
            }

            /// <summary>
            /// Add the rows of a node of the level before <paramref name="level"/> (the document for level 0).
            /// <paramref name="values"/> contains the values of the outer levels.
            /// </summary>
            internal void Collect(XmlNode node, int level, XLCellValue[] values, List<List<XLCellValue>> rows)
            {
                for (var c = 0; c < _columns.Count; c++)
                {
                    if (_columns[c].Level != level - 1)
                        continue;

                    var columnNode = _columns[c].Path.Length == 0 ? node : node.SelectSingleNode(_columns[c].Path, _context.Namespaces);
                    values[c] = _context.Convert(columnNode, _columns[c].XmlDataType);
                }

                var children = level < _levelPaths.Count ? node.SelectNodes(_levelPaths[level], _context.Namespaces)! : null;
                if (children is null || children.Count == 0)
                {
                    rows.Add(values.ToList());
                    return;
                }

                foreach (XmlNode child in children)
                {
                    var childValues = (XLCellValue[])values.Clone();
                    for (var c = 0; c < _columns.Count; c++)
                    {
                        if (_columns[c].Level >= level)
                            childValues[c] = Blank.Value;
                    }

                    Collect(child, level + 1, childValues, rows);
                }
            }
        }

        /// <summary>
        /// A step of the simple XPaths used by mappings, e.g. <c>ns1:person</c> or <c>@id</c>.
        /// </summary>
        internal readonly struct XPathStep : IEquatable<XPathStep>
        {
            private XPathStep(string prefix, string localName, bool isAttribute)
            {
                Prefix = prefix;
                LocalName = localName;
                IsAttribute = isAttribute;
            }

            internal string Prefix { get; }

            internal string LocalName { get; }

            internal bool IsAttribute { get; }

            internal static List<XPathStep> Parse(string xpath)
            {
                var steps = new List<XPathStep>();
                foreach (var part in xpath.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    var isAttribute = part.StartsWith("@", StringComparison.Ordinal);
                    var name = isAttribute ? part.Substring(1) : part;
                    var colon = name.IndexOf(':');
                    steps.Add(colon < 0
                        ? new XPathStep(string.Empty, name, isAttribute)
                        : new XPathStep(name.Substring(0, colon), name.Substring(colon + 1), isAttribute));
                }

                return steps;
            }

            internal static string Format(IReadOnlyList<XPathStep> steps, bool absolute)
            {
                var path = string.Join("/", steps.Select(s => s.ToString()));
                return absolute ? "/" + path : path;
            }

            internal XmlQualifiedName ToQualifiedName(XmlNamespaceManager namespaces)
            {
                return new XmlQualifiedName(LocalName, namespaces.LookupNamespace(Prefix) ?? string.Empty);
            }

            public bool Equals(XPathStep other)
            {
                return Prefix == other.Prefix && LocalName == other.LocalName && IsAttribute == other.IsAttribute;
            }

            public override bool Equals(object? obj) => obj is XPathStep other && Equals(other);

            public override int GetHashCode() => (Prefix, LocalName, IsAttribute).GetHashCode();

            public override string ToString()
            {
                var name = Prefix.Length == 0 ? LocalName : Prefix + ":" + LocalName;
                return IsAttribute ? "@" + name : name;
            }
        }
    }
}
