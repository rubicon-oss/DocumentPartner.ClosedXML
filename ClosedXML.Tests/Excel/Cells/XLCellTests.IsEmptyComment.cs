using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests
{
    /// <summary>
    /// Tests of the fork, kept out of <see cref="XLCellTests"/> to avoid merge conflicts with upstream.
    /// </summary>
    [TestFixture]
    public class IsEmptyCommentCellTests
    {
        [Test]
        public void IsEmpty_Comment ()
        {
            IXLWorksheet ws = new XLWorkbook ().Worksheets.Add ("Sheet1");
            IXLCell cell = ws.Cell (1, 1);
            cell.GetComment ().AddText ("comment");
            bool actual = cell.IsEmpty ();
            bool expected = false;
            Assert.AreEqual (expected, actual);
        }

        [Test]
        public void IsEmpty_Comment_Value ()
        {
            IXLWorksheet ws = new XLWorkbook ().Worksheets.Add ("Sheet1");
            IXLCell cell = ws.Cell (1, 1);
            cell.GetComment ().AddText ("comment");
            cell.SetValue ("value");

            bool actual = cell.IsEmpty ();
            bool expected = false;
            Assert.AreEqual (expected, actual);
        }

        [Test]
        [TestCase (XLCellsUsedOptions.Contents, true)]
        [TestCase (XLCellsUsedOptions.DataType, true)]
        [TestCase (XLCellsUsedOptions.NormalFormats, true)]
        [TestCase (XLCellsUsedOptions.ConditionalFormats, true)]
        [TestCase (XLCellsUsedOptions.Comments, false)]
        [TestCase (XLCellsUsedOptions.DataValidation, true)]
        [TestCase (XLCellsUsedOptions.MergedRanges, true)]
        [TestCase (XLCellsUsedOptions.Sparklines, true)]
        [TestCase (XLCellsUsedOptions.AllFormats, true)]
        [TestCase (XLCellsUsedOptions.AllContents, false)]
        [TestCase (XLCellsUsedOptions.All, false)]
        public void IsEmpty_Comment_Options (XLCellsUsedOptions options, bool expected)
        {
            IXLWorksheet ws = new XLWorkbook ().Worksheets.Add ("Sheet1");
            IXLCell cell = ws.Cell (1, 1);
            cell.GetComment ().AddText ("comment");

            bool actual = cell.IsEmpty (options);

            Assert.AreEqual (expected, actual);
        }

        [Test]
        [TestCase (XLCellsUsedOptions.Contents, false)]
        [TestCase (XLCellsUsedOptions.DataType, true)]
        [TestCase (XLCellsUsedOptions.NormalFormats, true)]
        [TestCase (XLCellsUsedOptions.ConditionalFormats, true)]
        [TestCase (XLCellsUsedOptions.Comments, false)]
        [TestCase (XLCellsUsedOptions.DataValidation, true)]
        [TestCase (XLCellsUsedOptions.MergedRanges, true)]
        [TestCase (XLCellsUsedOptions.Sparklines, true)]
        [TestCase (XLCellsUsedOptions.AllFormats, true)]
        [TestCase (XLCellsUsedOptions.AllContents, false)]
        [TestCase (XLCellsUsedOptions.All, false)]
        public void IsEmpty_Comment_Options_Value (XLCellsUsedOptions options, bool expected) // see #1575
        {
            IXLWorksheet ws = new XLWorkbook ().Worksheets.Add ("Sheet1");
            IXLCell cell = ws.Cell (1, 1);
            cell.GetComment().AddText ("comment");
            cell.SetValue ("value");

            bool actual = cell.IsEmpty (options);

            Assert.AreEqual (expected, actual);
        }
    }
}
