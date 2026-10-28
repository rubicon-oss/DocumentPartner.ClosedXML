using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using ClosedXML.XmlMaps.Tests.Infrastructure;
using ClosedXML.XmlMaps.Tests.Scenarios;
using NUnit.Framework;

namespace ClosedXML.XmlMaps.Tests
{
    /// <summary>
    /// The same operation is done by Excel (through COM) and by ClosedXML. The workbook saved by Excel is the
    /// expected result. Finally Excel opens the workbook saved by ClosedXML to check it understands the XML maps.
    /// </summary>
    [TestFixture]
    [Explicit("Automates Excel through COM.")]
    [Category("Excel")]
    [NonParallelizable]
    [Apartment(System.Threading.ApartmentState.STA)]
    public class ExcelComparisonTests
    {
        private const string ImportError = "Error";

        private readonly Dictionary<string, string> _templates = new();
        private ExcelSession? _excel;
        private string _templateDirectory = string.Empty;

        private ExcelSession Excel => _excel!;

        [OneTimeSetUp]
        public void StartExcel()
        {
            if (!ExcelSession.IsAvailable)
                Assert.Ignore("Excel is not installed.");

            _excel = ExcelSession.Start();
            _templateDirectory = Path.Combine(TestContext.CurrentContext.WorkDirectory, "Generated", "Templates");
            if (Directory.Exists(_templateDirectory))
                Directory.Delete(_templateDirectory, true);
            Directory.CreateDirectory(_templateDirectory);
        }

        [OneTimeTearDown]
        public void StopExcel()
        {
            _excel?.Dispose();
        }

        [TestCaseSource(typeof(ImportCases), nameof(ImportCases.All))]
        public void Import_MatchesExcel(ImportCase testCase)
        {
            var run = new ComparisonRun(this, testCase);
            var excelResults = run.RunExcel();
            var closedXmlResults = run.RunClosedXml();

            var differences = run.Compare();
            if (!closedXmlResults.SequenceEqual(excelResults))
                differences.Insert(0, $"import results: expected {string.Join(", ", excelResults)}, actual {string.Join(", ", closedXmlResults)}");
            foreach (var (mapName, expected) in run.ExcelExports)
            {
                var actual = run.ClosedXmlExports.GetValueOrDefault(mapName);
                if (actual != expected)
                    differences.Add($"export of map '{mapName}':{Environment.NewLine}expected {expected}{Environment.NewLine}actual {actual}");
            }

            AssertDifferences(differences, testCase.KnownDeviation);
            AssertExcelReadsSameMaps(run.ExcelPath, run.ClosedXmlPath);
            if (testCase.KnownDeviation is null && testCase.Scenario.CompareAllCells)
                AssertExcelCalculatesSameValues(run);
        }

        [TestCaseSource(nameof(DeleteCases))]
        public void Delete_MatchesExcel(ImportCase testCase, string mapName)
        {
            var run = new ComparisonRun(this, testCase);
            run.RunExcel(wb => wb.XmlMaps.Item(mapName).Delete());
            run.RunClosedXml(wb => wb.XmlMaps().Get(mapName).Delete());

            var differences = run.Compare();
            if (run.Expected.HasXmlMapsPart != run.Actual.HasXmlMapsPart)
                differences.Insert(0, $"XML maps part exists: expected {run.Expected.HasXmlMapsPart}, actual {run.Actual.HasXmlMapsPart}");

            AssertDifferences(differences, testCase.KnownDeviation);
            AssertExcelReadsSameMaps(run.ExcelPath, run.ClosedXmlPath);
        }

        [Test]
        public void DeleteWorksheet_MatchesExcel()
        {
            var run = new ComparisonRun(this, new ImportCase("TwoMaps delete sheet", Scenarios.Scenarios.TwoMaps));
            run.RunExcel(wb => wb.Worksheets.Item("Second").Delete());
            run.RunClosedXml(wb => wb.Worksheet("Second").Delete());

            AssertDifferences(run.Compare(), null);
            AssertExcelReadsSameMaps(run.ExcelPath, run.ClosedXmlPath);
        }

        [TestCaseSource(typeof(Scenarios.Scenarios), nameof(Scenarios.Scenarios.All))]
        public void RoundTrip_ExcelReadsMaps(XmlMapScenario scenario)
        {
            var directory = TestFiles.OutputDirectory();
            var template = GetTemplate(scenario);
            var closedXmlPath = Path.Combine(directory, "closedxml.xlsx");
            using (var workbook = new XLWorkbook(template))
                workbook.SaveAs(closedXmlPath, TestFiles.SaveOptions);

            var comparer = new SnapshotComparer(WorkbookSnapshot.Load(template), WorkbookSnapshot.Load(closedXmlPath));
            comparer.CompareMaps().CompareMappings().AssertNoDifferences();
            SnapshotComparer.AssertUniqueIds(WorkbookSnapshot.Load(closedXmlPath));
            AssertExcelReadsSameMaps(template, closedXmlPath);
        }

        /// <summary>
        /// A map imported by Excel from a file is bound to the file through a connection. ClosedXML keeps the binding when it saves
        /// into the loaded package, so Excel refreshes the data from the file like in the workbook it saved itself.
        /// </summary>
        [Test]
        public void DataBinding_IsKeptAndRefreshesLikeExcel()
        {
            var run = new ComparisonRun(this, new ImportCase("Inferred refresh", Scenarios.Scenarios.Inferred,
                new XmlImport(Scenarios.Scenarios.InferredContactsMap, Scenarios.Scenarios.InferredContactsXml(5))));
            run.RunClosedXml();

            var expectedPath = Path.Combine(Path.GetDirectoryName(run.ExcelPath)!, "excel_refreshed.xlsx");
            var actualPath = Path.Combine(Path.GetDirectoryName(run.ExcelPath)!, "closedxml_refreshed.xlsx");
            var sources = new List<string>();
            foreach (var (source, target) in new[] { (GetTemplate(Scenarios.Scenarios.Inferred), expectedPath), (run.ClosedXmlPath, actualPath) })
            {
                var workbook = Excel.Open(source, readOnly: true);
                try
                {
                    foreach (var map in workbook.XmlMaps)
                    {
                        sources.Add($"{map.Name}: {map.DataBinding.SourceUrl}");
                        Assert.That((int)map.DataBinding.Refresh(), Is.EqualTo(0), $"Excel: refresh of map '{map.Name}'");
                    }

                    ExcelSession.SaveAs(workbook, target);
                }
                finally
                {
                    ExcelSession.Close(workbook);
                }
            }

            Assert.That(sources.Skip(sources.Count / 2), Is.EqualTo(sources.Take(sources.Count / 2)), "Excel: source of data bindings");
            var comparer = new SnapshotComparer(WorkbookSnapshot.Load(expectedPath), WorkbookSnapshot.Load(actualPath));
            comparer.CompareMaps().CompareMappings().CompareCells(comparer.MappedArea()).AssertNoDifferences();
        }

        /// <summary>
        /// A workbook created from a template is a new package without the connections, the data bindings are left out,
        /// so Excel opens the workbook without errors and the maps work.
        /// </summary>
        [Test]
        public void DataBinding_FromTemplate_IsRemoved()
        {
            var path = Path.Combine(TestFiles.OutputDirectory(), "closedxml.xlsx");
            using (var workbook = XLWorkbook.OpenFromTemplate(GetTemplate(Scenarios.Scenarios.Inferred)))
            {
                workbook.XmlMaps().Get(Scenarios.Scenarios.InferredContactsMap).ImportXml(Scenarios.Scenarios.InferredContactsXml(3));
                workbook.SaveAs(path, TestFiles.SaveOptions);
            }

            using (var document = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(path, false))
            {
                var hasConnections = document.WorkbookPart!.ConnectionsPart is not null;
                var xml = new StreamReader(document.WorkbookPart.CustomXmlMappingsPart!.GetStream()).ReadToEnd();
                Assert.That(xml.Contains("DataBinding"), Is.EqualTo(hasConnections), "data bindings are written only with connections");
            }

            var excelWorkbook = Excel.Open(path, readOnly: true);
            try
            {
                Assert.That((int)excelWorkbook.XmlMaps.Count, Is.EqualTo(2), "Excel: number of maps");
                string xml;
                Assert.That((int)excelWorkbook.XmlMaps.Item(Scenarios.Scenarios.InferredContactsMap).ExportXml(out xml), Is.EqualTo(0), "Excel: export");
                Assert.That(xml, Does.Contain("<ns1:name>n3</ns1:name>").Or.Contain("<name>n3</name>"));
            }
            finally
            {
                ExcelSession.Close(excelWorkbook);
            }
        }

        /// <summary>
        /// Values entered into mapped cells (not imported) are exported like Excel does: numbers rounded to 15 digits without exponent,
        /// dates of the date system of the workbook, errors as empty elements and so on.
        /// </summary>
        [TestCase(false)]
        [TestCase(true)]
        public void Export_EnteredValues_MatchesExcel(bool date1904)
        {
            AssertExportMatchesExcel(date1904, new (string, object, string?)[]
            {
                ("string", "a<b>&\"c'd\nline2\r\nx\ty", null), ("normalizedString", 12345678901234567890d, null), ("token", true, null),
                ("int", 1e-20, null), ("integer", 1.5e300, null), ("long", "=1/0", null), ("decimal", 0.1, null),
                ("float", -0.000123456789012345, null), ("double", 123456.7890123456789, null), ("boolean", 1d, null), ("date", 59d, null),
                ("dateTime", 45000.75, null), ("time", 1.75, null), ("gYear", 2020.5, null), ("gYearMonth", 45000d, null),
                ("gMonthDay", 45000d, null), ("gDay", 31d, null), ("gMonth", 5d, null), ("duration", 3d, null), ("anyURI", " sp ", null),
                ("unsignedLong", -0d, null), ("short", "text in a number", null), ("nonNegativeInteger", false, null),
            });
        }

        /// <summary>
        /// Excel ignores the number format on export and has its own rules for numbers in date and g types: it rounds to milliseconds
        /// before splitting date and time, exports dates only in a range, and writes other numbers as plain numbers.
        /// </summary>
        [TestCaseSource(nameof(EnteredValueCases))]
        public void Export_EnteredValueEdgeCases_MatchesExcel(bool date1904, (string Type, object Value, string? Format)[] values)
        {
            AssertExportMatchesExcel(date1904, values);
        }

        private static IEnumerable<TestCaseData> EnteredValueCases()
        {
            var rounds = new Dictionary<string, (string, object, string?)[]>
            {
                ["number formats"] = new (string, object, string?)[]
                {
                    ("string", 1234.5, "#,##0.00"), ("normalizedString", 0.515, "0.00%"), ("token", 45000d, "yyyy-mm-dd"), ("decimal", 0.515, "0.00%"),
                    ("int", 42d, "0000"), ("integer", 1234.5678, "0.00"), ("long", "00042", "@"), ("double", 1234.5, "@"), ("float", -5d, "0;(0)"),
                    ("boolean", true, "@"), ("date", 45000d, "dd.mm.yyyy"), ("dateTime", 45000.5, "yyyy-mm-dd"), ("time", 0.5, "0.00"),
                    ("gYear", 45000d, "yyyy"), ("gMonth", 45000d, "mm"), ("gDay", 45000d, "dd"), ("anyURI", 1234.5, "$#,##0.00"),
                    ("short", 1e-7, "0.00E+00"), ("unsignedLong", 45000d, "dd.mm.yyyy"),
                },
                ["exponents"] = new (string, object, string?)[]
                {
                    ("double", 1e21, null), ("float", -1e-30, null), ("string", 1e-19, null), ("token", 1.2345e-16, null),
                    ("anyURI", 9.99999999999999e20, null), ("decimal", 1e25, null), ("long", -1e25, null), ("boolean", 1e25, null),
                    ("normalizedString", 2.2250738585072014e-308, null), ("date", 1e25, null), ("Name", 1.23456789012345e-7, null),
                    ("integer", 1e-30, null), ("short", -2.2251e-308, null), ("NCName", -2.2251e-308, null), ("duration", 1.5e300, null),
                },
                ["rounding to milliseconds"] = new (string, object, string?)[]
                {
                    ("date", 45000.9999999999, null), ("dateTime", 45000.9999999999, null), ("time", 0.9999999999, null),
                    ("gYearMonth", 45016.9999999999, null), ("gMonthDay", 59.9999999999, null), ("gYear", -5d, null), ("gDay", 0d, null),
                    ("gMonth", 12d, null),
                },
                ["29 February 1900"] = new (string, object, string?)[]
                {
                    ("date", 60d, null), ("dateTime", 59.9999999999, null), ("time", 2958466.5, null), ("gYearMonth", 3000000d, null),
                    ("gMonthDay", -2d, null), ("gYear", 0d, null), ("gDay", 1e-20, null), ("gMonth", 12.6, null),
                },
                ["negative serials"] = new (string, object, string?)[]
                {
                    ("date", -1e-10, null), ("dateTime", -1.25, null), ("time", -0.25, null), ("gYearMonth", -31d, null),
                    ("gMonthDay", -1461d, null), ("gYear", 999d, null), ("gDay", 30.999999999999996, null), ("gMonth", -0d, null),
                },
                ["lower limits of the 1900 date system"] = new (string, object, string?)[]
                {
                    ("date", -146087d, null), ("dateTime", -146087.5, null), ("time", -146088d, null), ("gYearMonth", -146037d, null),
                    ("gMonthDay", -146036d, null), ("gYear", 1e20, null), ("gDay", 31.0000000000001, null), ("gMonth", 1e15, null),
                },
                ["lower limits of the 1904 date system"] = new (string, object, string?)[]
                {
                    ("date", -146098d, null), ("dateTime", -146097.5, null), ("time", -146098.5, null), ("gYearMonth", -146099d, null),
                    ("gMonthDay", -1d, null), ("gYear", -12345d, null), ("gDay", 5d, null), ("gMonth", 13d, null),
                },
                ["upper limits"] = new (string, object, string?)[]
                {
                    ("date", 2958465.99999999, null), ("dateTime", 2957003.9999999, null), ("time", 2957004d, null), ("gYearMonth", 2958466d, null),
                    ("gMonthDay", 3000000.1, null), ("gYear", 1e-20, null), ("gDay", -1d, null), ("gMonth", 1d, null),
                },
            };
            foreach (var round in rounds)
                foreach (var date1904 in new[] { false, true })
                    yield return new TestCaseData(date1904, round.Value).SetName($"Export entered values {round.Key} (workbook {(date1904 ? "1904" : "1900")})");
        }

        private void AssertExportMatchesExcel(bool date1904, (string Type, object Value, string? Format)[] values)
        {
            var template = GetTemplate(date1904 ? Scenarios.Scenarios.Types1904 : Scenarios.Scenarios.Types);
            var directory = TestFiles.OutputDirectory();
            int Row(string type) => Array.IndexOf(Scenarios.Scenarios.TypeNames, type) + 1;
            var english = System.Globalization.CultureInfo.GetCultureInfo("en-US");

            var excelPath = Path.Combine(directory, "excel_export.xml");
            var workbook = Excel.Open(template, readOnly: true);
            string expectedResult;
            try
            {
                foreach (var (type, value, format) in values)
                {
                    object cell = workbook.Worksheets.Item("Types").Range($"B{Row(type)}");
                    var property = value is string text && text.StartsWith("=") ? "Formula" : "Value2";
                    try
                    {
                        // The number format first, so text stays text in a cell with the Text format.
                        if (format != null)
                            cell.GetType().InvokeMember("NumberFormat", System.Reflection.BindingFlags.SetProperty, null, cell, new object[] { format }, english);

                        cell.GetType().InvokeMember(property, System.Reflection.BindingFlags.SetProperty, null, cell, new[] { value }, english);
                    }
                    catch (System.Reflection.TargetInvocationException e)
                    {
                        throw new InvalidOperationException($"Excel: can't set {type} to '{value}'.", e);
                    }
                }

                expectedResult = ((XLXmlExportResult)(int)workbook.XmlMaps.Item(Scenarios.Scenarios.DataMap).Export(excelPath, true)).ToString();
            }
            finally
            {
                ExcelSession.Close(workbook);
            }

            string actualResult;
            string actualXml;
            using (var xlWorkbook = new XLWorkbook(template))
            {
                var ws = xlWorkbook.Worksheet("Types");
                foreach (var (type, value, format) in values)
                {
                    var cell = ws.Cell(Row(type), 2);
                    if (format != null)
                        cell.Style.NumberFormat.Format = format;

                    switch (value)
                    {
                        case string text when text.StartsWith("="):
                            cell.FormulaA1 = text;
                            break;
                        case string text:
                            cell.Value = text;
                            break;
                        case bool boolean:
                            cell.Value = boolean;
                            break;
                        default:
                            cell.Value = (double)value;
                            break;
                    }
                }

                actualResult = xlWorkbook.XmlMaps().Get(Scenarios.Scenarios.DataMap).ExportXml(out actualXml).ToString();
            }

            Assert.That(actualResult, Is.EqualTo(expectedResult), "export result");
            Assert.That(actualXml, Is.EqualTo(File.ReadAllText(excelPath)), "exported XML");
        }

        private static IEnumerable<TestCaseData> DeleteCases()
        {
            var rendering = ImportCases.RenderingTestCases.First();
            yield return new TestCaseData(new ImportCase("RenderingTest no import", rendering.Scenario), "root_Map").SetName("Delete RenderingTest no import");
            yield return new TestCaseData(rendering, "root_Map").SetName("Delete RenderingTest after import");
            var tables = ImportCases.TableCases.Single(c => c.Name == "Tables both");
            yield return new TestCaseData(tables, Scenarios.Scenarios.DataMap).SetName("Delete Tables after import");
            var twoMaps = ImportCases.TwoMapCases.Single(c => c.Name == "TwoMaps both");
            yield return new TestCaseData(twoMaps, Scenarios.Scenarios.DataMap).SetName("Delete TwoMaps first map");
            yield return new TestCaseData(twoMaps, Scenarios.Scenarios.OtherMap).SetName("Delete TwoMaps second map");
            var advanced = ImportCases.AdvancedCases.Single(c => c.Name == "Advanced valid");
            yield return new TestCaseData(advanced, Scenarios.Scenarios.AdvancedMap).SetName("Delete Advanced after import");
        }

        private static void AssertDifferences(IReadOnlyList<string> differences, string? knownDeviation)
        {
            var message = $"{differences.Count} difference(s) between Excel (expected) and ClosedXML (actual):{Environment.NewLine}" +
                          string.Join(Environment.NewLine, differences.Take(100));
            if (knownDeviation is null)
            {
                if (differences.Count > 0)
                    Assert.Fail(message);

                return;
            }

            if (differences.Count == 0)
                Assert.Fail($"The known deviation is gone, remove the marker: {knownDeviation}");

            Assert.Warn($"Known deviation: {knownDeviation}{Environment.NewLine}{message}");
        }

        /// <summary>
        /// Excel opens both workbooks. The workbook saved by ClosedXML must have the same maps, the same mapped cells
        /// and Excel must export the same XML from it.
        /// </summary>
        private void AssertExcelReadsSameMaps(string excelPath, string closedXmlPath)
        {
            var expectedWorkbook = Excel.Open(excelPath, readOnly: true);
            var actualWorkbook = Excel.Open(closedXmlPath, readOnly: true);
            try
            {
                Assert.That((int)actualWorkbook.XmlMaps.Count, Is.EqualTo((int)expectedWorkbook.XmlMaps.Count), "Excel: number of maps");
                foreach (var expectedMap in expectedWorkbook.XmlMaps)
                {
                    string name = expectedMap.Name;
                    var actualMap = actualWorkbook.XmlMaps.Item(name);
                    if (!(bool)expectedMap.IsExportable)
                        continue;

                    Assert.That((bool)actualMap.IsExportable, Is.True, $"Excel: map '{name}' is exportable");
                    var (expectedResult, expectedXml) = Export((object)expectedMap);
                    var (actualResult, actualXml) = Export((object)actualMap);
                    Assert.That(actualResult, Is.EqualTo(expectedResult), $"Excel: export result of map '{name}'");
                    Assert.That(NormalizeXml(actualXml), Is.EqualTo(NormalizeXml(expectedXml)), $"Excel: exported XML of map '{name}'");
                }

                foreach (var sheet in WorkbookSnapshot.Load(closedXmlPath).Sheets.Values)
                {
                    foreach (var mappedCell in sheet.MappedCells)
                    {
                        string xpath = actualWorkbook.Worksheets.Item(sheet.Name).Range(mappedCell.Reference).XPath.Value;
                        Assert.That(xpath, Is.EqualTo(mappedCell.XPath), $"Excel: XPath of {sheet.Name}!{mappedCell.Reference}");
                    }
                }
            }
            finally
            {
                ExcelSession.Close(actualWorkbook);
                ExcelSession.Close(expectedWorkbook);
            }
        }

        /// <summary>
        /// ClosedXML saves formulas without values, Excel calculates them when it opens the workbook. The values Excel calculates
        /// in the workbook saved by ClosedXML must be the values of the workbook Excel saved itself.
        /// </summary>
        private void AssertExcelCalculatesSameValues(ComparisonRun run)
        {
            var formulaCells = new SnapshotComparer(run.Expected, run.Actual).ExpectedFormulaCells().ToList();
            if (formulaCells.Count == 0)
                return;

            var calculatedPath = Path.Combine(Path.GetDirectoryName(run.ClosedXmlPath)!, "closedxml_calculated.xlsx");
            var workbook = Excel.Open(run.ClosedXmlPath, readOnly: true);
            try
            {
                Excel.App.CalculateFull();
                ExcelSession.SaveAs(workbook, calculatedPath);
            }
            finally
            {
                ExcelSession.Close(workbook);
            }

            var comparer = new SnapshotComparer(run.Expected, WorkbookSnapshot.Load(calculatedPath));
            comparer.CompareValues(formulaCells);
            if (comparer.Differences.Count > 0)
                Assert.Fail($"Values calculated by Excel differ:{Environment.NewLine}{string.Join(Environment.NewLine, comparer.Differences.Take(100))}");
        }

        private string GetTemplate(XmlMapScenario scenario)
        {
            if (_templates.TryGetValue(scenario.Name, out var path))
                return path;

            if (scenario.Resource is not null)
            {
                path = TestFiles.Extract(scenario.Resource, _templateDirectory);
            }
            else
            {
                path = Path.Combine(_templateDirectory, scenario.Name + ".xlsx");
                ExcelWorkbookBuilder.Build(Excel, path, scenario.Build!);
            }

            _templates[scenario.Name] = path;
            return path;
        }

        private static string NormalizeXml(string xml) => Regex.Replace(xml, @">\s+<", "><").Trim();

        /// <summary>
        /// Export the data of a map in Excel. Excel refuses to export a map without any data.
        /// </summary>
        private static (string Result, string Xml) Export(dynamic map)
        {
            try
            {
                int result = map.ExportXml(out string xml);
                return (result.ToString(), xml);
            }
            catch (COMException e)
            {
                return ("Error: " + e.Message, string.Empty);
            }
        }

        /// <summary>
        /// Result of an export of a map, the XML is empty, if the map couldn't be exported.
        /// </summary>
        private sealed record ExportSnapshot(bool IsExportable, string Result, string Xml)
        {
            public override string ToString() => $"exportable {IsExportable}, result {Result}:{Environment.NewLine}{Xml}";
        }

        /// <summary>
        /// The imports of a test case, once done by Excel and once by ClosedXML, each saved to its own file.
        /// </summary>
        private sealed class ComparisonRun
        {
            private readonly ExcelComparisonTests _fixture;
            private readonly ImportCase _testCase;
            private readonly string _template;
            private readonly List<string> _xmlFiles;

            internal ComparisonRun(ExcelComparisonTests fixture, ImportCase testCase)
            {
                _fixture = fixture;
                _testCase = testCase;
                var directory = TestFiles.OutputDirectory();
                _template = fixture.GetTemplate(testCase.Scenario);
                _xmlFiles = testCase.Imports.Select((import, i) =>
                {
                    var path = Path.Combine(directory, $"import{i + 1}.xml");
                    File.WriteAllText(path, import.Xml);
                    return path;
                }).ToList();
                ExcelPath = Path.Combine(directory, "excel.xlsx");
                ClosedXmlPath = Path.Combine(directory, "closedxml.xlsx");
            }

            internal string ExcelPath { get; }

            internal string ClosedXmlPath { get; }

            /// <summary>
            /// Export of each map by Excel after the imports.
            /// </summary>
            internal Dictionary<string, ExportSnapshot> ExcelExports { get; } = new();

            /// <summary>
            /// Export of each map by ClosedXML after the imports.
            /// </summary>
            internal Dictionary<string, ExportSnapshot> ClosedXmlExports { get; } = new();

            internal WorkbookSnapshot Expected { get; private set; } = null!;

            internal WorkbookSnapshot Actual { get; private set; } = null!;

            internal List<string> RunExcel(Action<dynamic>? afterImports = null)
            {
                var results = new List<string>();
                var workbook = _fixture.Excel.Open(_template, readOnly: true);
                try
                {
                    for (var i = 0; i < _xmlFiles.Count; i++)
                    {
                        var import = _testCase.Imports[i];
                        try
                        {
                            results.Add(((XLXmlImportResult)(int)workbook.XmlMaps.Item(import.MapName).Import(_xmlFiles[i], import.Overwrite)).ToString());
                        }
                        catch (COMException)
                        {
                            results.Add(ImportError);
                        }
                    }

                    if (afterImports is null)
                    {
                        foreach (var map in workbook.XmlMaps)
                        {
                            string name = map.Name;
                            var path = Path.Combine(Path.GetDirectoryName(ExcelPath)!, $"excel_export_{name}.xml");
                            string result;
                            try
                            {
                                result = ((XLXmlExportResult)(int)map.Export(path, true)).ToString();
                            }
                            catch (COMException)
                            {
                                result = ImportError;
                            }

                            ExcelExports[name] = new ExportSnapshot((bool)map.IsExportable, result, File.Exists(path) ? File.ReadAllText(path) : string.Empty);
                        }
                    }

                    afterImports?.Invoke(workbook);
                    ExcelSession.SaveAs(workbook, ExcelPath);
                }
                finally
                {
                    ExcelSession.Close(workbook);
                }

                return results;
            }

            internal List<string> RunClosedXml(Action<XLWorkbook>? afterImports = null)
            {
                var results = new List<string>();
                using var workbook = new XLWorkbook(_template);
                for (var i = 0; i < _xmlFiles.Count; i++)
                {
                    var import = _testCase.Imports[i];
                    try
                    {
                        results.Add(workbook.XmlMaps().Get(import.MapName).Import(_xmlFiles[i], import.Overwrite).ToString());
                    }
                    catch (Exception e) when (e is InvalidOperationException or System.Xml.XmlException)
                    {
                        results.Add(ImportError);
                    }
                }

                if (afterImports is null)
                {
                    foreach (var map in workbook.XmlMaps())
                    {
                        var path = Path.Combine(Path.GetDirectoryName(ClosedXmlPath)!, $"closedxml_export_{map.Name}.xml");
                        string result;
                        try
                        {
                            result = map.Export(path, true).ToString();
                        }
                        catch (InvalidOperationException)
                        {
                            result = ImportError;
                        }

                        ClosedXmlExports[map.Name] = new ExportSnapshot(map.IsExportable, result, File.Exists(path) ? File.ReadAllText(path) : string.Empty);
                    }
                }

                afterImports?.Invoke(workbook);
                workbook.SaveAs(ClosedXmlPath, TestFiles.SaveOptions);
                return results;
            }

            /// <summary>
            /// Compare the saved workbooks: maps, mappings and cells.
            /// </summary>
            internal List<string> Compare()
            {
                Expected = WorkbookSnapshot.Load(ExcelPath);
                Actual = WorkbookSnapshot.Load(ClosedXmlPath);
                var comparer = new SnapshotComparer(Expected, Actual);
                var cells = comparer.MappedArea();
                if (_testCase.Scenario.CompareAllCells)
                    cells = cells.Concat(comparer.AllCellsWithValue());

                comparer.CompareMaps().CompareMappings().CompareCells(cells);
                return comparer.Differences.ToList();
            }
        }
    }
}
