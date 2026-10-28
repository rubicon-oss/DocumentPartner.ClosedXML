using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using ClosedXML.Excel;
using ClosedXML.XmlMaps.Tests.Infrastructure;
using NUnit.Framework;
using static ClosedXML.XmlMaps.Tests.Scenarios.Scenarios;

namespace ClosedXML.XmlMaps.Tests
{
    /// <summary>
    /// Export without Excel. The exported XML itself is verified against Excel by <see cref="ExcelComparisonTests"/>,
    /// these tests cover the API and things Excel can't be asked about.
    /// </summary>
    [TestFixture]
    public class ExportTests
    {
        [Test]
        [SetCulture("de-AT")]
        public void Export_IsIndependentOfCulture()
        {
            using var wb = LoadResource("Types.xlsx");
            var map = wb.XmlMaps().Get(DataMap);
            map.ImportXml(DataXml(new Dictionary<string, string> { ["decimal"] = "1234.5", ["date"] = "2008-04-01", ["dateTime"] = "2008-04-01T13:14:15" }));

            map.ExportXml(out var xml);

            Assert.That(xml, Does.Contain("<ns1:v_decimal>1234.5</ns1:v_decimal>"));
            Assert.That(xml, Does.Contain("<ns1:v_date>2008-04-01</ns1:v_date>"));
            Assert.That(xml, Does.Contain("<ns1:v_dateTime>2008-04-01T13:14:15.000</ns1:v_dateTime>"));
        }

        [Test]
        public void Export_ImportOfExportedXml_GivesSameCells()
        {
            using var source = LoadResource("Tables.xlsx");
            source.XmlMaps().Get(DataMap).ImportXml(DataXml(rows: Rows(4), items: Items(3)));
            source.XmlMaps().Get(DataMap).ExportXml(out var xml);

            using var target = LoadResource("Tables.xlsx");
            var result = target.XmlMaps().Get(DataMap).ImportXml(xml);

            Assert.That(result, Is.EqualTo(XLXmlImportResult.Success));
            var sourceCells = source.Worksheet("Tables").Range("B2:K8");
            var targetCells = target.Worksheet("Tables").Range("B2:K8");
            foreach (var cell in sourceCells.Cells())
                Assert.That(targetCells.Cell(cell.Address.RowNumber - 1, cell.Address.ColumnNumber - 1).Value, Is.EqualTo(cell.Value), cell.Address.ToString());
        }

        [Test]
        public void Export_ToStreamAndFile_GiveSameXmlAsExportXml()
        {
            using var wb = LoadResource("Types.xlsx");
            var map = wb.XmlMaps().Get(DataMap);
            map.ImportXml(DataXml(new Dictionary<string, string> { ["string"] = "äöü €" }));
            map.ExportXml(out var xml);

            using var stream = new MemoryStream();
            map.Export(stream);
            var path = Path.Combine(TestFiles.OutputDirectory(), "export.xml");
            map.Export(path, overwrite: true);

            Assert.That(Encoding.UTF8.GetString(stream.ToArray()), Is.EqualTo(xml));
            Assert.That(File.ReadAllBytes(path), Is.EqualTo(stream.ToArray()), "UTF-8 without byte order mark");
        }

        [Test]
        public void Export_ToExistingFile_ThrowsWithoutOverwrite()
        {
            using var wb = LoadResource("Types.xlsx");
            var map = wb.XmlMaps().Get(DataMap);
            var path = Path.Combine(TestFiles.OutputDirectory(), "export.xml");
            File.WriteAllText(path, "existing");

            Assert.Throws<IOException>(() => map.Export(path));
            Assert.That(File.ReadAllText(path), Is.EqualTo("existing"));
        }

        [Test]
        public void Export_TablesWithoutData_Throws()
        {
            using var wb = LoadResource("Tables.xlsx");
            var map = wb.XmlMaps().Get(DataMap);
            map.ImportXml(DataXml(rows: Rows(2), items: Items(2)));
            map.ImportXml(DataXml());

            Assert.Throws<InvalidOperationException>(() => map.ExportXml(out _));
        }

        [Test]
        public void Export_DeletedMap_Throws()
        {
            using var wb = LoadResource("Types.xlsx");
            var map = wb.XmlMaps().Get(DataMap);
            map.Delete();

            Assert.That(map.IsExportable, Is.False);
            Assert.Throws<InvalidOperationException>(() => map.ExportXml(out _));
        }

        /// <summary>
        /// Load a workbook from a resource. ClosedXML needs the source stream again when saving, so it stays open.
        /// </summary>
        private static XLWorkbook LoadResource(string resource)
        {
            var memory = new MemoryStream();
            using (var stream = TestFiles.Open(resource))
                stream.CopyTo(memory);

            memory.Position = 0;
            return new XLWorkbook(memory);
        }
    }
}
