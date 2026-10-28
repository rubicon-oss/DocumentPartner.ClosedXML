using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ClosedXML.XmlMaps.Tests.Infrastructure
{
    /// <summary>
    /// The content of a saved xlsx file relevant for XML maps, read directly with the OpenXML SDK,
    /// so files written by Excel and by ClosedXML are compared independently of ClosedXML's loader.
    /// </summary>
    internal sealed class WorkbookSnapshot
    {
        private static readonly Dictionary<uint, string> BuiltInFormats = new()
        {
            [0] = "General", [1] = "0", [2] = "0.00", [3] = "#,##0", [4] = "#,##0.00", [9] = "0%", [10] = "0.00%",
            [11] = "0.00E+00", [12] = "# ?/?", [13] = "# ??/??", [14] = "mm-dd-yy", [15] = "d-mmm-yy", [16] = "d-mmm",
            [17] = "mmm-yy", [18] = "h:mm AM/PM", [19] = "h:mm:ss AM/PM", [20] = "h:mm", [21] = "h:mm:ss",
            [22] = "m/d/yy h:mm", [37] = "#,##0 ;(#,##0)", [38] = "#,##0 ;[Red](#,##0)", [39] = "#,##0.00;(#,##0.00)",
            [40] = "#,##0.00;[Red](#,##0.00)", [45] = "mm:ss", [46] = "[h]:mm:ss", [47] = "mmss.0", [48] = "##0.0E+0", [49] = "@",
        };

        internal Dictionary<string, SheetSnapshot> Sheets { get; } = new();

        internal List<MapSnapshot> Maps { get; } = new();

        internal bool HasXmlMapsPart { get; private set; }

        internal static WorkbookSnapshot Load(string path)
        {
            using var document = SpreadsheetDocument.Open(path, false);
            var workbookPart = document.WorkbookPart!;
            var snapshot = new WorkbookSnapshot();

            var mapNames = new Dictionary<uint, string>();
            var mapInfo = workbookPart.CustomXmlMappingsPart?.MapInfo;
            snapshot.HasXmlMapsPart = workbookPart.CustomXmlMappingsPart is not null;
            if (mapInfo is not null)
            {
                foreach (var map in mapInfo.Elements<Map>())
                {
                    var schema = mapInfo.Elements<Schema>().FirstOrDefault(s => s.Id?.Value == map.SchemaId?.Value);
                    mapNames[map.ID!.Value] = map.Name!.Value!;
                    snapshot.Maps.Add(new MapSnapshot(
                        map.Name.Value!,
                        map.RootElement?.Value,
                        schema?.Namespace?.Value,
                        NormalizeWhitespace(schema?.InnerXml ?? string.Empty),
                        map.ShowImportExportErrors?.Value ?? false,
                        map.AutoFit?.Value ?? false,
                        map.AppendData?.Value ?? false,
                        map.PreserveAutoFilterState?.Value ?? false,
                        map.PreserveFormat?.Value ?? false));
                }
            }

            var sharedStrings = workbookPart.SharedStringTablePart?.SharedStringTable?.Elements<SharedStringItem>()
                .Select(si => si.Text?.Text ?? string.Concat(si.Elements<Run>().Select(r => r.Text?.Text)))
                .ToList() ?? new List<string>();
            var stylesheet = workbookPart.WorkbookStylesPart?.Stylesheet;
            var customFormats = stylesheet?.NumberingFormats?.Elements<NumberingFormat>()
                .ToDictionary(f => f.NumberFormatId!.Value, f => f.FormatCode!.Value!) ?? new Dictionary<uint, string>();
            var cellFormats = stylesheet?.CellFormats?.Elements<CellFormat>().ToList() ?? new List<CellFormat>();

            string NumberFormat(uint styleIndex)
            {
                var id = styleIndex < cellFormats.Count ? cellFormats[(int)styleIndex].NumberFormatId?.Value ?? 0 : 0;
                // An empty custom format code is the same as General.
                if (customFormats.TryGetValue(id, out var code) && code.Length > 0)
                    return code;

                return BuiltInFormats.TryGetValue(id, out code) ? code : "id:" + id;
            }

            foreach (var sheet in workbookPart.Workbook.Sheets!.Elements<Sheet>())
            {
                var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!);
                var sheetSnapshot = new SheetSnapshot(sheet.Name!.Value!);
                snapshot.Sheets[sheetSnapshot.Name] = sheetSnapshot;

                var worksheet = worksheetPart.Worksheet;
                foreach (var column in worksheet.Elements<Columns>().SelectMany(c => c.Elements<Column>()))
                {
                    for (var c = column.Min!.Value; c <= column.Max!.Value; c++)
                    {
                        sheetSnapshot.ColumnStyles[(int)c] = column.Style?.Value ?? 0;
                        if (column.Width is not null)
                            sheetSnapshot.ColumnWidths[(int)c] = column.Width.Value;
                    }
                }

                // Excel writes the formula of a shared formula only into the first cell. Formulas are compared
                // as text and the scenarios use only structured references, so the other cells get the same text.
                var sharedFormulas = new Dictionary<uint, string>();

                // Excel writes the formula of an array formula only into the first cell, ClosedXML into all cells, and only Excel calculates
                // the values of the other cells. These cells are compared as members of the array formula.
                var arrayRanges = worksheet.Descendants<CellFormula>()
                    .Where(f => f.FormulaType?.Value == CellFormulaValues.Array && f.Reference?.Value is not null)
                    .Select(f => ParseRange(f.Reference!.Value!))
                    .ToList();
                foreach (var row in worksheet.Descendants<Row>())
                {
                    if (row.CustomFormat?.Value == true && row.StyleIndex is not null)
                        sheetSnapshot.RowStyles[(int)row.RowIndex!.Value] = row.StyleIndex.Value;
                    if (row.Hidden?.Value == true)
                        sheetSnapshot.HiddenRows.Add((int)row.RowIndex!.Value);

                    foreach (var cell in row.Elements<Cell>())
                    {
                        var (rowNumber, columnNumber) = ParseReference(cell.CellReference!.Value!);
                        var (type, value) = ReadValue(cell, sharedStrings);
                        var formula = cell.CellFormula;
                        var formulaText = formula?.Text;
                        if (formula?.FormulaType?.Value == CellFormulaValues.Shared && formula.SharedIndex is { } sharedIndex)
                        {
                            if (!string.IsNullOrEmpty(formulaText))
                                sharedFormulas[sharedIndex.Value] = formulaText!;
                            else
                                formulaText = sharedFormulas.GetValueOrDefault(sharedIndex.Value);
                        }

                        if (arrayRanges.Any(r => r.FirstRow <= rowNumber && rowNumber <= r.LastRow && r.FirstColumn <= columnNumber && columnNumber <= r.LastColumn &&
                                                 (r.FirstRow != rowNumber || r.FirstColumn != columnNumber)))
                        {
                            formulaText = "{array member}";
                        }

                        sheetSnapshot.Cells[(rowNumber, columnNumber)] = new CellSnapshot(
                            type,
                            value,
                            formulaText,
                            cell.StyleIndex?.Value ?? 0);
                    }
                }

                sheetSnapshot.NumberFormatResolver = NumberFormat;

                foreach (var tablePart in worksheetPart.TableDefinitionParts)
                {
                    var table = tablePart.Table;
                    var columns = table.TableColumns!.Elements<TableColumn>().Select(tc => new TableColumnSnapshot(
                        tc.Name!.Value!,
                        tc.UniqueName?.Value,
                        tc.XmlColumnProperties?.XPath?.Value,
                        tc.XmlColumnProperties?.MapId is { } mapId ? mapNames.GetValueOrDefault(mapId.Value, "<missing map " + mapId.Value + ">") : null,
                        tc.XmlColumnProperties?.XmlDataType?.InnerText,
                        tc.CalculatedColumnFormula?.Text,
                        tc.TotalsRowFunction?.InnerText,
                        tc.TotalsRowLabel?.Value)).ToList();
                    sheetSnapshot.Tables.Add(new TableSnapshot(
                        table.Name!.Value!,
                        table.Reference!.Value!,
                        table.TableType?.InnerText ?? "worksheet",
                        table.Id!.Value,
                        table.TotalsRowCount?.Value ?? 0,
                        columns)
                    {
                        Filters = string.Join("; ", table.AutoFilter?.Elements<FilterColumn>().Select(DescribeFilter) ?? Enumerable.Empty<string>()),
                        Sort = DescribeSort(table.SortState ?? table.AutoFilter?.GetFirstChild<SortState>()),
                    });
                }

                var singleXmlCells = worksheetPart.SingleCellTablePart?.SingleXmlCells;
                foreach (var singleXmlCell in singleXmlCells?.Elements<SingleXmlCell>() ?? Enumerable.Empty<SingleXmlCell>())
                {
                    var xmlProperties = singleXmlCell.XmlCellProperties?.XmlProperties;
                    sheetSnapshot.MappedCells.Add(new MappedCellSnapshot(
                        singleXmlCell.CellReference!.Value!,
                        singleXmlCell.Id!.Value,
                        xmlProperties?.XPath?.Value,
                        xmlProperties?.MapId is { } mapId ? mapNames.GetValueOrDefault(mapId.Value, "<missing map " + mapId.Value + ">") : null,
                        xmlProperties?.XmlDataType?.InnerText,
                        singleXmlCell.XmlCellProperties?.UniqueName?.Value));
                }
            }

            return snapshot;
        }

        internal static (int Row, int Column) ParseReference(string reference)
        {
            var match = Regex.Match(reference, @"^\$?([A-Z]+)\$?(\d+)$");
            var column = match.Groups[1].Value.Aggregate(0, (acc, c) => acc * 26 + c - 'A' + 1);
            return (int.Parse(match.Groups[2].Value, CultureInfo.InvariantCulture), column);
        }

        internal static (int FirstRow, int FirstColumn, int LastRow, int LastColumn) ParseRange(string reference)
        {
            var parts = reference.Split(':');
            var first = ParseReference(parts[0]);
            var last = parts.Length > 1 ? ParseReference(parts[1]) : first;
            return (first.Row, first.Column, last.Row, last.Column);
        }

        internal static string FormatReference(int row, int column)
        {
            var letters = string.Empty;
            for (var c = column; c > 0; c = (c - 1) / 26)
                letters = (char)('A' + (c - 1) % 26) + letters;

            return letters + row.ToString(CultureInfo.InvariantCulture);
        }

        private static (string Type, string? Value) ReadValue(Cell cell, IReadOnlyList<string> sharedStrings)
        {
            var dataType = cell.DataType?.InnerText ?? "n";
            var text = cell.CellValue?.Text;
            switch (dataType)
            {
                case "s":
                    return ("text", text is null ? null : sharedStrings[int.Parse(text, CultureInfo.InvariantCulture)]);
                case "inlineStr":
                    return ("text", cell.InlineString?.Text?.Text ?? string.Concat(cell.InlineString?.Elements<Run>().Select(r => r.Text?.Text) ?? Array.Empty<string>()));
                case "str":
                    return ("text", text);
                case "b":
                    return ("bool", text == "1" ? "TRUE" : "FALSE");
                case "e":
                    return ("error", text);
                case "d":
                    return ("date", text);
                default:
                    return string.IsNullOrEmpty(text) ? ("blank", null) : ("number", text);
            }
        }

        private static string DescribeFilter(FilterColumn column)
        {
            var conditions = column.Descendants<CustomFilter>().Select(f => $"{f.Operator?.InnerText ?? "equal"} {f.Val?.Value}")
                .Concat(column.Descendants<Filter>().Select(f => $"= {f.Val?.Value}"));
            return $"{column.ColumnId?.Value}: {string.Join(", ", conditions)}";
        }

        private static string DescribeSort(SortState? sortState)
        {
            if (sortState is null)
                return string.Empty;

            var conditions = sortState.Elements<SortCondition>().Select(c => $"{c.Reference?.Value}{(c.Descending?.Value == true ? " desc" : string.Empty)}");
            return $"{sortState.Reference?.Value}: {string.Join(", ", conditions)}";
        }

        private static string NormalizeWhitespace(string text) => Regex.Replace(text, @"\s+", " ").Replace("> <", "><").Trim();
    }

    internal sealed class SheetSnapshot
    {
        internal SheetSnapshot(string name)
        {
            Name = name;
        }

        internal string Name { get; }

        internal Dictionary<(int Row, int Column), CellSnapshot> Cells { get; } = new();

        internal Dictionary<int, uint> RowStyles { get; } = new();

        internal SortedSet<int> HiddenRows { get; } = new();

        internal Dictionary<int, uint> ColumnStyles { get; } = new();

        internal Dictionary<int, double> ColumnWidths { get; } = new();

        internal List<TableSnapshot> Tables { get; } = new();

        internal List<MappedCellSnapshot> MappedCells { get; } = new();

        internal Func<uint, string> NumberFormatResolver { get; set; } = _ => "General";

        internal CellSnapshot? GetCell(int row, int column) => Cells.TryGetValue((row, column), out var cell) ? cell : null;

        /// <summary>
        /// Number format of a cell, considering the styles of cells, rows and columns like Excel does.
        /// </summary>
        internal string GetNumberFormat(int row, int column)
        {
            if (Cells.TryGetValue((row, column), out var cell))
                return NumberFormatResolver(cell.StyleIndex);
            if (RowStyles.TryGetValue(row, out var rowStyle))
                return NumberFormatResolver(rowStyle);
            if (ColumnStyles.TryGetValue(column, out var columnStyle))
                return NumberFormatResolver(columnStyle);

            return NumberFormatResolver(0);
        }
    }

    internal sealed record CellSnapshot(string Type, string? Value, string? Formula, uint StyleIndex);

    internal sealed record TableSnapshot(string Name, string Reference, string TableType, uint Id, uint TotalsRowCount, List<TableColumnSnapshot> Columns)
    {
        /// <summary>
        /// Filter conditions of the autofilter by column.
        /// </summary>
        internal string Filters { get; init; } = string.Empty;

        /// <summary>
        /// Range and conditions of the sort state.
        /// </summary>
        internal string Sort { get; init; } = string.Empty;
    }

    internal sealed record TableColumnSnapshot(
        string Name,
        string? UniqueName,
        string? XPath,
        string? MapName,
        string? XmlDataType,
        string? CalculatedColumnFormula,
        string? TotalsRowFunction,
        string? TotalsRowLabel);

    internal sealed record MappedCellSnapshot(string Reference, uint Id, string? XPath, string? MapName, string? XmlDataType, string? UniqueName);

    internal sealed record MapSnapshot(
        string Name,
        string? RootElement,
        string? SchemaNamespace,
        string SchemaXml,
        bool ShowImportExportErrors,
        bool AutoFit,
        bool Append,
        bool PreserveAutoFilterState,
        bool PreserveFormat);
}
