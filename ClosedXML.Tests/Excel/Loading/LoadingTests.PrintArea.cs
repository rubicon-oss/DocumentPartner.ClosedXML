using ClosedXML.Excel;
using NUnit.Framework;
using System.IO;

namespace ClosedXML.Tests.Excel
{
    /// <summary>
    /// Tests of the fork, kept out of <see cref="LoadingTests"/> to avoid merge conflicts with upstream.
    /// </summary>
    [TestFixture]
    public class PrintAreaLoadingTests
    {
        [Test]
        public void CanLoadAndSavePrintArea()
        {
            using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\PrintArea.xlsx"));
            using var wb = new XLWorkbook(stream);

            AssertPrintAreas(wb);

            using var ms = new MemoryStream();
            wb.SaveAs(ms, true);

            using var wb2 = new XLWorkbook(ms);
            AssertPrintAreas(wb2);

            static void AssertPrintAreas(XLWorkbook wb)
            {
                var table1PrintAreas = wb.Worksheets.Worksheet("Tabelle 1").PageSetup.PrintAreas.PrintArea;
                Assert.AreEqual("IF('Rohdaten !'!$A$2<5,'Tabelle 1'!$A$1:$C$40,'Tabelle 1'!$A$1:$C$80)", table1PrintAreas);

                var table2PrintAreas = wb.Worksheets.Worksheet("Tabelle 2").PageSetup.PrintAreas.PrintArea;
                Assert.AreEqual("IF('Rohdaten !'!$A$2<5,'Tabelle 2'!$A$1:$C$40,'Tabelle 2'!$A$1:$C$80),'Tabelle 2'!$J$17:$M$29", table2PrintAreas);

                var table3PrintAreas = wb.Worksheets.Worksheet("Tabelle 3").PageSetup.PrintAreas.PrintArea;
                Assert.AreEqual("'Tabelle 3'!$C$4:$G$12,'Tabelle 3'!$F$17:$J$23", table3PrintAreas);
            }
        }
    }
}
