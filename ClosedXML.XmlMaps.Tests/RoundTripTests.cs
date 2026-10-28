using System.IO;
using System.Linq;
using ClosedXML.Excel;
using ClosedXML.XmlMaps.Tests.Infrastructure;
using NUnit.Framework;

namespace ClosedXML.XmlMaps.Tests
{
    /// <summary>
    /// Load and save workbooks with XML maps without Excel. The workbook saved by Excel is the expected result.
    /// </summary>
    [TestFixture]
    public class RoundTripTests
    {
        private const string RenderingTest = "RenderingTest.xlsx";
        private const string TwoMaps = "TwoMaps.xlsx";

        [Test]
        public void LoadAndSaveAs_KeepsMapsAndMappings()
        {
            var directory = TestFiles.OutputDirectory();
            var source = TestFiles.Extract(RenderingTest, directory);
            var target = Path.Combine(directory, "closedxml.xlsx");

            using (var wb = new XLWorkbook(source))
                wb.SaveAs(target, TestFiles.SaveOptions);

            AssertSameMappings(source, target);
        }

        [Test]
        public void LoadAndSaveToStream_KeepsMapsAndMappings()
        {
            var directory = TestFiles.OutputDirectory();
            var source = TestFiles.Extract(RenderingTest, directory);
            var target = Path.Combine(directory, "closedxml.xlsx");

            using (var input = TestFiles.Open(RenderingTest))
            using (var wb = new XLWorkbook(input))
            using (var output = File.Create(target))
                wb.SaveAs(output, TestFiles.SaveOptions);

            AssertSameMappings(source, target);
        }

        [Test]
        public void OpenFromTemplate_WritesMapsIntoNewPackage()
        {
            var directory = TestFiles.OutputDirectory();
            var source = TestFiles.Extract(RenderingTest, directory);
            var target = Path.Combine(directory, "closedxml.xlsx");

            using (var wb = XLWorkbook.OpenFromTemplate(source))
                wb.SaveAs(target, TestFiles.SaveOptions);

            AssertSameMappings(source, target);
        }

        [Test]
        public void SaveTwice_DoesNotDuplicateMappings()
        {
            var directory = TestFiles.OutputDirectory();
            var source = TestFiles.Extract(RenderingTest, directory);
            var first = Path.Combine(directory, "first.xlsx");
            var second = Path.Combine(directory, "second.xlsx");

            using (var wb = new XLWorkbook(source))
            {
                wb.SaveAs(first, TestFiles.SaveOptions);
                wb.SaveAs(second, TestFiles.SaveOptions);
            }

            AssertSameMappings(source, second);
        }

        [Test]
        public void XmlMaps_ExposesLoadedMap()
        {
            using var input = TestFiles.Open(RenderingTest);
            using var wb = new XLWorkbook(input);

            var maps = wb.XmlMaps();

            Assert.That(maps.Count, Is.EqualTo(1));
            var map = maps.Get("root_Map");
            Assert.Multiple(() =>
            {
                Assert.That(map.Id, Is.EqualTo(5));
                Assert.That(map.RootElement, Is.EqualTo("root"));
                Assert.That(map.SchemaId, Is.EqualTo("Schema3"));
                Assert.That(map.SchemaNamespace, Is.EqualTo("http://www.rubicon-it.com/Correspondence/RenderingTest/4.4"));
                Assert.That(map.SchemaXml, Does.Contain("<xs:element name=\"root\">"));
                Assert.That(map.SelectionNamespaces, Is.EqualTo("xmlns:ns2='http://www.rubicon-it.com/Correspondence/RenderingTest/4.4'"));
                Assert.That(map.ShowImportExportValidationErrors, Is.False);
                Assert.That(map.AdjustColumnWidth, Is.True);
                Assert.That(map.AppendOnImport, Is.False);
                Assert.That(map.PreserveColumnFilter, Is.True);
                Assert.That(map.PreserveNumberFormatting, Is.True);
            });
        }

        [Test]
        public void WorkbookWithoutMaps_GetsNoXmlMapParts()
        {
            var directory = TestFiles.OutputDirectory();
            var target = Path.Combine(directory, "closedxml.xlsx");

            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet();
                ws.Cell("A1").Value = "Name";
                ws.Range("A1:A2").CreateTable();
                Assert.That(wb.XmlMaps().Count, Is.Zero);
                wb.SaveAs(target, TestFiles.SaveOptions);
            }

            var snapshot = WorkbookSnapshot.Load(target);
            Assert.That(snapshot.HasXmlMapsPart, Is.False);
            Assert.That(snapshot.Sheets.Values.Single().Tables.Single().TableType, Is.EqualTo("worksheet"));
        }

        [Test]
        public void InsertedRowsAndColumns_MoveMappings()
        {
            var directory = TestFiles.OutputDirectory();
            var source = TestFiles.Extract(RenderingTest, directory);
            var target = Path.Combine(directory, "closedxml.xlsx");

            using (var wb = new XLWorkbook(source))
            {
                var ws = wb.Worksheet("First Tab");
                ws.Row(1).InsertRowsAbove(2);
                ws.Column(1).InsertColumnsBefore(1);
                wb.SaveAs(target, TestFiles.SaveOptions);
            }

            var sheet = WorkbookSnapshot.Load(target).Sheets["First Tab"];
            Assert.That(sheet.MappedCells.Select(c => c.Reference), Has.Member("C9").And.Member("C33").And.No.Member("B7"));
            Assert.That(sheet.MappedCells.Single(c => c.Reference == "C9").XPath, Is.EqualTo("/ns2:root/ns2:string"));
            var table = sheet.Tables.Single();
            Assert.That(table.Reference, Is.EqualTo("B44:F45"));
            Assert.That(table.TableType, Is.EqualTo("xml"));
        }

        [Test]
        public void DeletedTable_KeepsMapAndCells()
        {
            var directory = TestFiles.OutputDirectory();
            var source = TestFiles.Extract(RenderingTest, directory);
            var target = Path.Combine(directory, "closedxml.xlsx");

            using (var wb = new XLWorkbook(source))
            {
                wb.Worksheet("First Tab").Tables.Remove("List11");
                wb.SaveAs(target, TestFiles.SaveOptions);
            }

            var snapshot = WorkbookSnapshot.Load(target);
            Assert.That(snapshot.Maps.Select(m => m.Name), Is.EqualTo(new[] { "root_Map" }));
            Assert.That(snapshot.Sheets["First Tab"].Tables, Is.Empty);
            Assert.That(snapshot.Sheets["First Tab"].MappedCells, Has.Count.EqualTo(25));
        }

        [Test]
        public void DeletedWorksheet_KeepsMapWithoutMappings()
        {
            var directory = TestFiles.OutputDirectory();
            var source = TestFiles.Extract(TwoMaps, directory);
            var target = Path.Combine(directory, "closedxml.xlsx");

            using (var wb = new XLWorkbook(source))
            {
                wb.Worksheet("Second").Delete();
                wb.SaveAs(target, TestFiles.SaveOptions);
            }

            var snapshot = WorkbookSnapshot.Load(target);
            var first = snapshot.Sheets["First"];
            Assert.Multiple(() =>
            {
                Assert.That(snapshot.Maps.Select(m => m.Name), Is.EqualTo(new[] { "data_Map", "other_Map" }));
                Assert.That(first.MappedCells.Select(c => c.MapName), Is.All.EqualTo("data_Map"));
                Assert.That(first.Tables.Single().Columns.Select(c => c.MapName), Is.All.EqualTo("data_Map"));
            });
            SnapshotComparer.AssertUniqueIds(snapshot);
            XmlMapPartValidator.AssertValid(target, source);
        }

        [Test]
        public void TwoMaps_KeepMapIdsAndIdsAreUnique()
        {
            var directory = TestFiles.OutputDirectory();
            var source = TestFiles.Extract(TwoMaps, directory);
            var target = Path.Combine(directory, "closedxml.xlsx");

            using (var wb = new XLWorkbook(source))
                wb.SaveAs(target, TestFiles.SaveOptions);

            AssertSameMappings(source, target);
        }

        [Test]
        public void DeletedMap_RemovesAllXmlMapData()
        {
            var directory = TestFiles.OutputDirectory();
            var source = TestFiles.Extract(RenderingTest, directory);
            var target = Path.Combine(directory, "closedxml.xlsx");

            using (var wb = new XLWorkbook(source))
            {
                var map = wb.XmlMaps().Get("root_Map");
                map.Delete();
                Assert.That(wb.XmlMaps().Count, Is.Zero);
                Assert.That(() => map.ImportXml("<root/>"), Throws.InvalidOperationException);
                wb.SaveAs(target, TestFiles.SaveOptions);
            }

            var snapshot = WorkbookSnapshot.Load(target);
            var sheet = snapshot.Sheets["First Tab"];
            Assert.Multiple(() =>
            {
                Assert.That(snapshot.HasXmlMapsPart, Is.False);
                Assert.That(sheet.MappedCells, Is.Empty);
                Assert.That(sheet.Tables.Single().TableType, Is.EqualTo("worksheet"));
                Assert.That(sheet.Tables.Single().Columns.Select(c => c.XPath), Is.All.Null);
                Assert.That(sheet.GetCell(7, 1)?.Value, Is.EqualTo("string"));
            });
        }

        private static void AssertSameMappings(string expectedPath, string actualPath)
        {
            var expected = WorkbookSnapshot.Load(expectedPath);
            var actual = WorkbookSnapshot.Load(actualPath);
            var comparer = new SnapshotComparer(expected, actual).CompareMaps().CompareMappings();
            comparer.AssertNoDifferences();
            SnapshotComparer.AssertUniqueIds(actual);
            XmlMapPartValidator.AssertValid(actualPath, expectedPath);
        }
    }
}
