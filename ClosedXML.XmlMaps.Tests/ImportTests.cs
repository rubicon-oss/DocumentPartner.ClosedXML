using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using ClosedXML.Excel;
using ClosedXML.XmlMaps.Tests.Infrastructure;
using NUnit.Framework;
using static ClosedXML.XmlMaps.Tests.Scenarios.Scenarios;

namespace ClosedXML.XmlMaps.Tests
{
    /// <summary>
    /// Import without Excel. The behavior itself is verified against Excel by <see cref="ExcelComparisonTests"/>,
    /// these tests cover the API and things Excel can't be asked about, like the culture of the thread.
    /// </summary>
    [TestFixture]
    public class ImportTests
    {
        [Test]
        [SetCulture("de-AT")]
        [SetUICulture("de-AT")]
        public void Import_IsIndependentOfCulture()
        {
            using var wb = LoadResource("Types.xlsx");
            var result = wb.XmlMaps().Get(DataMap).ImportXml(DataXml(new Dictionary<string, string>
            {
                ["decimal"] = "1234.5", ["double"] = "1E3", ["date"] = "2008-04-01", ["dateTime"] = "2008-04-01T13:14:15",
                ["time"] = "14:20:00", ["string"] = "1,5",
            }));

            var ws = wb.Worksheet("Types");
            Assert.Multiple(() =>
            {
                Assert.That(result, Is.EqualTo(XLXmlImportResult.Success));
                Assert.That(TypeCell(ws, "decimal").Value, Is.EqualTo((XLCellValue)1234.5));
                Assert.That(TypeCell(ws, "double").Value, Is.EqualTo((XLCellValue)1000.0));
                Assert.That(TypeCell(ws, "date").GetDateTime(), Is.EqualTo(new DateTime(2008, 4, 1)));
                Assert.That(TypeCell(ws, "dateTime").GetDateTime(), Is.EqualTo(new DateTime(2008, 4, 1, 13, 14, 15)));
                Assert.That(TypeCell(ws, "time").GetTimeSpan(), Is.EqualTo(new TimeSpan(14, 20, 0)));
                Assert.That(TypeCell(ws, "string").Value, Is.EqualTo((XLCellValue)"1,5"));
            });
        }

        [Test]
        public void Import_FromStreamAndFile_GiveSameResult()
        {
            var xml = DataXml(new Dictionary<string, string> { ["string"] = "text", ["int"] = "7" });
            var path = Path.Combine(TestFiles.OutputDirectory(), "data.xml");
            File.WriteAllText(path, xml);

            using var fromFile = LoadResource("Types.xlsx");
            using var fromStream = LoadResource("Types.xlsx");
            fromFile.XmlMaps().Get(DataMap).Import(path);
            using (var stream = new MemoryStream(Encoding.UTF8.GetBytes(xml)))
                fromStream.XmlMaps().Get(DataMap).Import(stream);

            Assert.That(TypeCell(fromStream.Worksheet("Types"), "string").Value, Is.EqualTo(TypeCell(fromFile.Worksheet("Types"), "string").Value));
            Assert.That(TypeCell(fromStream.Worksheet("Types"), "int").Value, Is.EqualTo((XLCellValue)7.0));
        }

        [Test]
        public void Import_ValidationFailure_StillImportsData()
        {
            using var wb = LoadResource("Types.xlsx");

            var result = wb.XmlMaps().Get(DataMap).ImportXml(DataXml(new Dictionary<string, string> { ["int"] = "abc" }));

            Assert.That(result, Is.EqualTo(XLXmlImportResult.ValidationFailed));
            Assert.That(TypeCell(wb.Worksheet("Types"), "int").Value, Is.EqualTo((XLCellValue)"abc"));
        }

        [Test]
        public void Import_TooLongText_IsTruncated()
        {
            using var wb = LoadResource("Types.xlsx");

            var result = wb.XmlMaps().Get(DataMap).ImportXml(DataXml(new Dictionary<string, string> { ["string"] = new string('x', 40000) }));

            Assert.That(result, Is.EqualTo(XLXmlImportResult.ElementsTruncated));
            Assert.That(TypeCell(wb.Worksheet("Types"), "string").GetText(), Has.Length.EqualTo(32767));
        }

        [Test]
        public void Import_DtdIsNotProcessed()
        {
            using var wb = LoadResource("Types.xlsx");
            var xml = "<!DOCTYPE data [<!ENTITY e SYSTEM \"file:///c:/windows/win.ini\">]>" + DataXml(new Dictionary<string, string> { ["string"] = "&e;" });

            Assert.That(() => wb.XmlMaps().Get(DataMap).ImportXml(xml), Throws.InstanceOf<System.Xml.XmlException>());
        }

        [Test]
        public void Import_NotWellFormed_ThrowsAndKeepsCells()
        {
            using var wb = LoadResource("Types.xlsx");
            var map = wb.XmlMaps().Get(DataMap);
            map.ImportXml(DataXml(new Dictionary<string, string> { ["string"] = "kept" }));

            Assert.That(() => map.ImportXml($"<data xmlns=\"{DataNamespace}\"><v_string>a</data>"), Throws.InstanceOf<System.Xml.XmlException>());
            Assert.That(TypeCell(wb.Worksheet("Types"), "string").Value, Is.EqualTo((XLCellValue)"kept"));
        }

        [TestCase("<data xmlns=\"urn:other\"><v_string>a</v_string></data>")]
        [TestCase("<data><v_string>a</v_string></data>")]
        public void Import_RootNotInSchema_FailsValidationAndClearsCells(string xml)
        {
            using var wb = LoadResource("Types.xlsx");
            var map = wb.XmlMaps().Get(DataMap);
            map.ImportXml(DataXml(new Dictionary<string, string> { ["string"] = "cleared" }));

            var result = map.ImportXml(xml);

            Assert.That(result, Is.EqualTo(XLXmlImportResult.ValidationFailed));
            Assert.That(TypeCell(wb.Worksheet("Types"), "string").IsEmpty(), Is.True);
        }

        [Test]
        public void Import_KeepsWhitespaceOfCData()
        {
            using var wb = LoadResource("Types.xlsx");

            wb.XmlMaps().Get(DataMap).ImportXml(DataXml(new Dictionary<string, string>
            {
                ["string"] = "  a<![CDATA[ x ]]>b  ", ["normalizedString"] = "  <![CDATA[ x ]]>  ", ["int"] = "<![CDATA[ 5 ]]>",
            }));

            var ws = wb.Worksheet("Types");
            Assert.Multiple(() =>
            {
                Assert.That(TypeCell(ws, "string").Value, Is.EqualTo((XLCellValue)"a x b"));
                Assert.That(TypeCell(ws, "normalizedString").Value, Is.EqualTo((XLCellValue)" x "));
                Assert.That(TypeCell(ws, "int").Value, Is.EqualTo((XLCellValue)" 5 "));
            });
        }

        [Test]
        public void Append_WithoutXmlTable_Throws()
        {
            using var wb = LoadResource("Types.xlsx");

            Assert.That(() => wb.XmlMaps().Get(DataMap).ImportXml(DataXml(), overwrite: false), Throws.InvalidOperationException);
        }

        [Test]
        public void Import_GrowsTableAndKeepsMappingAfterSave()
        {
            var target = Path.Combine(TestFiles.OutputDirectory(), "closedxml.xlsx");
            using (var wb = LoadResource("Tables.xlsx"))
            {
                wb.XmlMaps().Get(DataMap).ImportXml(DataXml(rows: Rows(5)));
                var table = wb.Worksheet("Tables").Tables.Single(t => t.RangeAddress.FirstAddress.ColumnLetter == "B");
                Assert.That(table.RangeAddress.ToString(), Is.EqualTo("B2:F7"));
                wb.SaveAs(target, TestFiles.SaveOptions);
            }

            using (var wb = new XLWorkbook(target))
            {
                wb.XmlMaps().Get(DataMap).ImportXml(DataXml(rows: Rows(2)));
                var table = wb.Worksheet("Tables").Tables.Single(t => t.RangeAddress.FirstAddress.ColumnLetter == "B");
                Assert.That(table.RangeAddress.ToString(), Is.EqualTo("B2:F4"));
                Assert.That(wb.Worksheet("Tables").Cell("C4").Value, Is.EqualTo((XLCellValue)"text 2"));
            }
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

        private static IXLCell TypeCell(IXLWorksheet ws, string type)
        {
            return ws.Cell(Array.IndexOf(TypeNames, type) + 1, 2);
        }
    }
}
