using System;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using ClosedXML.XmlMaps.Tests.Infrastructure;
using NUnit.Framework;

namespace ClosedXML.XmlMaps.Tests
{
    /// <summary>
    /// Without <see cref="XLWorkbook.EnableXmlMaps"/>, XML maps are neither loaded nor saved.
    /// </summary>
    [TestFixture]
    [NonParallelizable]
    public class DisabledTests
    {
        private const string RenderingTest = "RenderingTest.xlsx";

        [SetUp]
        public void Disable() => XLWorkbook.EnableXmlMaps = false;

        [TearDown]
        public void Enable() => XLWorkbook.EnableXmlMaps = true;

        [Test]
        public void XmlMaps_Throws()
        {
            using var wb = new XLWorkbook();

            Assert.Throws<InvalidOperationException>(() => wb.XmlMaps());
        }

        /// <summary>
        /// Like ClosedXML without XML map support: the parts of the source package are passed through
        /// untouched, the tables written by ClosedXML lose their XML properties.
        /// </summary>
        [Test]
        public void SaveAs_DoesNotTouchXmlMapParts()
        {
            var directory = TestFiles.OutputDirectory();
            var source = TestFiles.Extract(RenderingTest, directory);
            var target = Path.Combine(directory, "closedxml.xlsx");

            using (var wb = new XLWorkbook(source))
                wb.SaveAs(target, TestFiles.SaveOptions);

            var expected = WorkbookSnapshot.Load(source);
            var actual = WorkbookSnapshot.Load(target);
            Assert.Multiple(() =>
            {
                Assert.That(actual.Maps.Select(m => m.Name), Is.EqualTo(expected.Maps.Select(m => m.Name)));
                Assert.That(actual.Sheets["First Tab"].MappedCells, Is.EqualTo(expected.Sheets["First Tab"].MappedCells));
                Assert.That(actual.Sheets["First Tab"].Tables.Select(t => t.TableType), Is.All.EqualTo("worksheet"));
            });
        }

        [Test]
        public void EnabledAfterLoad_DoesNotSeeMaps()
        {
            using var input = TestFiles.Open(RenderingTest);
            using var wb = new XLWorkbook(input);

            XLWorkbook.EnableXmlMaps = true;

            Assert.That(wb.XmlMaps().Count, Is.Zero);
        }
    }
}
