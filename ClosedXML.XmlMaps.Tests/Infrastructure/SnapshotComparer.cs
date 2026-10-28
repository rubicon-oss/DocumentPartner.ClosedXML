using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using NUnit.Framework;

namespace ClosedXML.XmlMaps.Tests.Infrastructure
{
    /// <summary>
    /// Compares a workbook written by Excel (expected) with a workbook written by ClosedXML (actual).
    /// </summary>
    internal sealed class SnapshotComparer
    {
        private readonly WorkbookSnapshot _expected;
        private readonly WorkbookSnapshot _actual;
        private readonly List<string> _differences = new();

        internal SnapshotComparer(WorkbookSnapshot expected, WorkbookSnapshot actual)
        {
            _expected = expected;
            _actual = actual;
        }

        internal IReadOnlyList<string> Differences => _differences;

        /// <summary>
        /// Fail the test with all differences found so far.
        /// </summary>
        internal void AssertNoDifferences()
        {
            if (_differences.Count > 0)
                Assert.Fail($"{_differences.Count} difference(s) between Excel (expected) and ClosedXML (actual):{Environment.NewLine}" +
                            string.Join(Environment.NewLine, _differences.Take(100)));
        }

        internal SnapshotComparer CompareMaps()
        {
            Check("map count", _expected.Maps.Count, _actual.Maps.Count);
            foreach (var expected in _expected.Maps)
            {
                var actual = _actual.Maps.FirstOrDefault(m => m.Name == expected.Name);
                if (actual is null)
                {
                    _differences.Add($"map '{expected.Name}' is missing");
                    continue;
                }

                if (actual != expected)
                    _differences.Add($"map '{expected.Name}': expected {expected with { SchemaXml = "..." }}, actual {actual with { SchemaXml = "..." }}");
                if (actual.SchemaXml != expected.SchemaXml)
                    _differences.Add($"map '{expected.Name}': schema differs");
            }

            return this;
        }

        internal SnapshotComparer CompareMappings()
        {
            foreach (var expectedSheet in _expected.Sheets.Values)
            {
                if (!_actual.Sheets.TryGetValue(expectedSheet.Name, out var actualSheet))
                {
                    _differences.Add($"sheet '{expectedSheet.Name}' is missing");
                    continue;
                }

                var expectedCells = expectedSheet.MappedCells.ToDictionary(c => c.Reference);
                var actualCells = actualSheet.MappedCells.ToDictionary(c => c.Reference);
                foreach (var reference in expectedCells.Keys.Union(actualCells.Keys))
                {
                    expectedCells.TryGetValue(reference, out var expected);
                    actualCells.TryGetValue(reference, out var actual);
                    if (expected is null || actual is null)
                    {
                        _differences.Add($"{expectedSheet.Name}!{reference}: mapped cell expected {Describe(expected)}, actual {Describe(actual)}");
                        continue;
                    }

                    if (expected with { Id = 0 } != actual with { Id = 0 })
                        _differences.Add($"{expectedSheet.Name}!{reference}: mapped cell expected {expected}, actual {actual}");
                }

                foreach (var expectedTable in expectedSheet.Tables)
                {
                    var actualTable = actualSheet.Tables.FirstOrDefault(t => t.Name == expectedTable.Name);
                    if (actualTable is null)
                    {
                        _differences.Add($"{expectedSheet.Name}: table '{expectedTable.Name}' is missing");
                        continue;
                    }

                    Check($"{expectedSheet.Name}: table '{expectedTable.Name}' reference", expectedTable.Reference, actualTable.Reference);
                    Check($"{expectedSheet.Name}: table '{expectedTable.Name}' type", expectedTable.TableType, actualTable.TableType);
                    Check($"{expectedSheet.Name}: table '{expectedTable.Name}' filters", expectedTable.Filters, actualTable.Filters);
                    Check($"{expectedSheet.Name}: table '{expectedTable.Name}' sort", expectedTable.Sort, actualTable.Sort);
                    Check($"{expectedSheet.Name}: table '{expectedTable.Name}' column count", expectedTable.Columns.Count, actualTable.Columns.Count);
                    foreach (var (expectedColumn, actualColumn) in expectedTable.Columns.Zip(actualTable.Columns))
                        Check($"{expectedSheet.Name}: table '{expectedTable.Name}' column", expectedColumn, actualColumn);
                }

                Check($"{expectedSheet.Name}: table count", expectedSheet.Tables.Count, actualSheet.Tables.Count);
                Check($"{expectedSheet.Name}: hidden rows", string.Join(",", expectedSheet.HiddenRows), string.Join(",", actualSheet.HiddenRows));
            }

            return this;
        }

        /// <summary>
        /// Compare values, types and number formats of cells.
        /// </summary>
        internal SnapshotComparer CompareCells(IEnumerable<(string Sheet, int Row, int Column)> cells)
        {
            foreach (var (sheetName, row, column) in cells.Distinct())
            {
                if (!_expected.Sheets.TryGetValue(sheetName, out var expectedSheet) || !_actual.Sheets.TryGetValue(sheetName, out var actualSheet))
                    continue;

                var reference = $"{sheetName}!{WorkbookSnapshot.FormatReference(row, column)}";
                var expected = expectedSheet.GetCell(row, column);
                var actual = actualSheet.GetCell(row, column);
                var expectedType = expected?.Type ?? "blank";
                var actualType = actual?.Type ?? "blank";
                var hasFormula = expected?.Formula is not null || actual?.Formula is not null;
                if (hasFormula)
                {
                    // ClosedXML doesn't recalculate formulas like Excel does, only compare the formula itself.
                    // Formulas of totals rows are written by ClosedXML without the table name (SUBTOTAL(109,[Qty])),
                    // which is equivalent inside the table.
                    Check($"{reference} formula", NormalizeFormula(expected?.Formula), NormalizeFormula(actual?.Formula));
                }
                else if (expectedType != actualType || !ValuesEqual(expectedType, expected?.Value, actual?.Value))
                {
                    _differences.Add($"{reference}: expected {expectedType} '{expected?.Value}', actual {actualType} '{actual?.Value}'");
                }

                Check($"{reference} number format", expectedSheet.GetNumberFormat(row, column), actualSheet.GetNumberFormat(row, column));
            }

            return this;
        }

        /// <summary>
        /// Compare values and types of cells, also of cells with formulas.
        /// </summary>
        internal SnapshotComparer CompareValues(IEnumerable<(string Sheet, int Row, int Column)> cells)
        {
            foreach (var (sheetName, row, column) in cells.Distinct())
            {
                if (!_expected.Sheets.TryGetValue(sheetName, out var expectedSheet) || !_actual.Sheets.TryGetValue(sheetName, out var actualSheet))
                    continue;

                var expected = expectedSheet.GetCell(row, column);
                var actual = actualSheet.GetCell(row, column);
                var expectedType = expected?.Type ?? "blank";
                var actualType = actual?.Type ?? "blank";
                if (expectedType != actualType || !ValuesEqual(expectedType, expected?.Value, actual?.Value))
                    _differences.Add($"{sheetName}!{WorkbookSnapshot.FormatReference(row, column)} value: expected {expectedType} '{expected?.Value}', actual {actualType} '{actual?.Value}'");
            }

            return this;
        }

        /// <summary>
        /// Cells with a formula in the expected workbook.
        /// </summary>
        internal IEnumerable<(string Sheet, int Row, int Column)> ExpectedFormulaCells()
        {
            return _expected.Sheets.Values.SelectMany(sheet => sheet.Cells.Where(c => c.Value.Formula is not null).Select(c => (sheet.Name, c.Key.Row, c.Key.Column)));
        }

        /// <summary>
        /// Cells where XML maps put data: mapped cells and the areas of tables of both workbooks,
        /// including some rows below and columns right of the tables to detect moved cells.
        /// </summary>
        internal IEnumerable<(string Sheet, int Row, int Column)> MappedArea(int extraRows = 30, int extraColumns = 2)
        {
            foreach (var snapshot in new[] { _expected, _actual })
            {
                foreach (var sheet in snapshot.Sheets.Values)
                {
                    foreach (var mappedCell in sheet.MappedCells)
                    {
                        var (row, column) = WorkbookSnapshot.ParseReference(mappedCell.Reference);
                        yield return (sheet.Name, row, column);
                    }

                    foreach (var table in sheet.Tables)
                    {
                        var (firstRow, firstColumn, lastRow, lastColumn) = WorkbookSnapshot.ParseRange(table.Reference);
                        for (var row = firstRow; row <= lastRow + extraRows; row++)
                        {
                            for (var column = firstColumn; column <= lastColumn + extraColumns; column++)
                                yield return (sheet.Name, row, column);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// All cells that have a value in either workbook.
        /// </summary>
        internal IEnumerable<(string Sheet, int Row, int Column)> AllCellsWithValue()
        {
            return new[] { _expected, _actual }
                .SelectMany(s => s.Sheets.Values)
                .SelectMany(sheet => sheet.Cells.Where(c => c.Value.Type != "blank").Select(c => (sheet.Name, c.Key.Row, c.Key.Column)));
        }

        internal static void AssertUniqueIds(WorkbookSnapshot snapshot)
        {
            var ids = snapshot.Sheets.Values.SelectMany(s => s.Tables.Select(t => ("table " + t.Name, t.Id))
                    .Concat(s.MappedCells.Select(c => ($"{s.Name}!{c.Reference}", c.Id))))
                .ToList();
            var duplicates = ids.GroupBy(i => i.Id).Where(g => g.Count() > 1).Select(g => $"id {g.Key}: {string.Join(", ", g.Select(i => i.Item1))}").ToList();
            Assert.That(duplicates, Is.Empty, "Table ids and ids of mapped cells must be unique in the workbook.");
        }

        private static bool ValuesEqual(string type, string? expected, string? actual)
        {
            if (type == "number" &&
                double.TryParse(expected, NumberStyles.Float, CultureInfo.InvariantCulture, out var e) &&
                double.TryParse(actual, NumberStyles.Float, CultureInfo.InvariantCulture, out var a))
            {
                return Math.Abs(e - a) <= 1e-9 * Math.Max(1, Math.Abs(e));
            }

            return string.Equals(expected ?? string.Empty, actual ?? string.Empty, StringComparison.Ordinal);
        }

        private static string? NormalizeFormula(string? formula)
        {
            return formula is null ? null : Regex.Replace(formula, @"\b[A-Za-z_][\w.]*(?=\[)", string.Empty);
        }

        private void Check<T>(string what, T expected, T actual)
        {
            if (!EqualityComparer<T>.Default.Equals(expected, actual))
                _differences.Add($"{what}: expected '{expected}', actual '{actual}'");
        }

        private static string Describe(object? value) => value?.ToString() ?? "<none>";
    }
}
