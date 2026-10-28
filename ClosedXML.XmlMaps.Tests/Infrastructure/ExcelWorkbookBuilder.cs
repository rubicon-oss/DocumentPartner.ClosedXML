using System.IO;
using System.Reflection;

namespace ClosedXML.XmlMaps.Tests.Infrastructure
{
    /// <summary>
    /// Builds a workbook with XML maps in Excel, so test templates are created by Excel itself and need no hand-made files.
    /// </summary>
    internal sealed class ExcelWorkbookBuilder
    {
        private const int XlSrcRange = 1;
        private const int XlYes = 1;

        private readonly dynamic _workbook;
        private readonly string _directory;
        private int _schemaCount;

        private ExcelWorkbookBuilder(dynamic workbook, string directory)
        {
            _workbook = workbook;
            _directory = directory;
        }

        /// <summary>
        /// Create a new workbook in Excel, let <paramref name="build"/> fill it and save it to <paramref name="path"/>.
        /// </summary>
        internal static void Build(ExcelSession excel, string path, System.Action<ExcelWorkbookBuilder> build)
        {
            var workbook = excel.App.Workbooks.Add();
            try
            {
                // A new workbook may have more than one sheet depending on Excel settings.
                while (workbook.Worksheets.Count > 1)
                    workbook.Worksheets.Item(workbook.Worksheets.Count).Delete();

                build(new ExcelWorkbookBuilder(workbook, Path.GetDirectoryName(path)!));
                ExcelSession.SaveAs(workbook, path);
            }
            finally
            {
                ExcelSession.Close(workbook);
            }
        }

        internal dynamic Workbook => _workbook;

        /// <summary>
        /// Get a sheet by name. The first sheet of the workbook is renamed, further sheets are added.
        /// </summary>
        internal dynamic Sheet(string name)
        {
            foreach (var sheet in _workbook.Worksheets)
            {
                if (sheet.Name == name)
                    return sheet;
            }

            var first = _workbook.Worksheets.Item(1);
            if (first.Name.StartsWith("Sheet") || first.Name.StartsWith("Tabelle"))
            {
                first.Name = name;
                return first;
            }

            var added = _workbook.Worksheets.Add(After: _workbook.Worksheets.Item(_workbook.Worksheets.Count));
            added.Name = name;
            return added;
        }

        /// <summary>
        /// Add an XML map from a schema.
        /// </summary>
        internal dynamic AddMap(string xsd, string rootElement)
        {
            var path = Path.Combine(_directory, $"schema{++_schemaCount}.xsd");
            File.WriteAllText(path, xsd);
            return _workbook.XmlMaps.Add(path, rootElement);
        }

        /// <summary>
        /// Import XML data without schema into a new map. Excel infers the schema from the data, creates the map and a table for the data.
        /// </summary>
        internal dynamic ImportInferred(string sheet, string address, string xml)
        {
            var path = Path.Combine(_directory, $"inferred{++_schemaCount}.xml");
            File.WriteAllText(path, xml);
            // The map is an out parameter, which late-bound calls through dynamic don't support.
            object workbook = _workbook;
            var args = new object?[] { path, new System.Runtime.InteropServices.DispatchWrapper(null), true, Sheet(sheet).Range(address) };
            var modifier = new ParameterModifier(args.Length);
            modifier[1] = true;
            workbook.GetType().InvokeMember("XmlImport", BindingFlags.InvokeMethod, null, workbook, args, new[] { modifier }, null, null);
            return args[1]!;
        }

        internal void MapCell(string sheet, string address, dynamic map, string xpath, string selectionNamespace)
        {
            Sheet(sheet).Range(address).XPath.SetValue(map, xpath, selectionNamespace, false);
        }

        /// <summary>
        /// Create a table at <paramref name="range"/> (header row + data rows) and map its columns to repeating elements.
        /// </summary>
        internal dynamic MapTable(string sheet, string range, dynamic map, string selectionNamespace, params string[] xpaths)
        {
            var ws = Sheet(sheet);
            var table = ws.ListObjects.Add(XlSrcRange, ws.Range(range), Missing.Value, XlYes);
            for (var i = 0; i < xpaths.Length; i++)
                table.ListColumns.Item(i + 1).Range.XPath.SetValue(map, xpaths[i], selectionNamespace, true);

            return table;
        }

        internal const int TotalsSum = 1;
        internal const int TotalsAverage = 2;
        internal const int TotalsCount = 3;
        internal const int TotalsMax = 6;

        internal void Set(string sheet, string address, object value)
        {
            Sheet(sheet).Range(address).Value2 = value;
        }

        /// <summary>
        /// Set a formula. The <c>Formula</c> property always uses English function names and separators.
        /// </summary>
        internal void Formula(string sheet, string address, string formula)
        {
            Sheet(sheet).Range(address).Formula = formula;
        }

        /// <summary>
        /// Add a column that isn't mapped and has a formula for all rows.
        /// </summary>
        /// <param name="position">1-based position of the new column, <c>null</c> for the end of the table.</param>
        internal void AddCalculatedColumn(dynamic table, int? position, string name, string formula)
        {
            var column = position is null ? table.ListColumns.Add() : table.ListColumns.Add(position.Value);
            column.Name = name;

            // A new column takes the format of its neighbor, a formula in a text formatted cell would stay text.
            object range = column.DataBodyRange;
            range.GetType().InvokeMember("NumberFormat", BindingFlags.SetProperty, null, range, new object[] { "General" }, System.Globalization.CultureInfo.GetCultureInfo("en-US"));
            column.DataBodyRange.Formula = formula;
        }

        /// <summary>
        /// Set the function of a column in the totals row, e.g. <see cref="TotalsSum"/>.
        /// </summary>
        internal void SetTotals(dynamic table, string column, int calculation)
        {
            table.ListColumns.Item(column).TotalsCalculation = calculation;
        }

        /// <summary>
        /// Set a number format with English format codes. Late-bound COM calls pass the user locale to Excel,
        /// which would interpret e.g. <c>0.00</c> with the decimal separator of the user, so the property is
        /// set through <see cref="System.Type.InvokeMember(string, BindingFlags, Binder, object, object[], System.Globalization.CultureInfo)"/> with en-US.
        /// </summary>
        internal void Format(string sheet, string address, string numberFormat)
        {
            object range = Sheet(sheet).Range(address);
            range.GetType().InvokeMember("NumberFormat", BindingFlags.SetProperty, null, range, new object[] { numberFormat }, System.Globalization.CultureInfo.GetCultureInfo("en-US"));
        }
    }
}
