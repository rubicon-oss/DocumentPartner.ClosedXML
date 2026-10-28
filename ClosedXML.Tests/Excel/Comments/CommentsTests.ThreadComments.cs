using ClosedXML.Excel;
using NUnit.Framework;
using System.IO;
using System.Linq;

namespace ClosedXML.Tests.Excel.Comments
{
    /// <summary>
    /// Tests of the fork, kept out of <see cref="CommentsTests"/> to avoid merge conflicts with upstream.
    /// </summary>
    public class ThreadCommentsTests
    {
        [Test]
        [TestCase(ThreadedCommentLoading.Skip, "")]
        [TestCase(ThreadedCommentLoading.ConvertToNotes, @"This is a threaded commentThis is a reply.")]
        public void CanReadThreadedCommentNote (ThreadedCommentLoading threadedCommentLoading, string expectedComments)
        {
            ThreadedCommentsLoadOptions.ThreadedCommentLoading = threadedCommentLoading;
            try
            {
                Assert.Multiple (() =>
                {
                    using var stream = TestHelper.GetStreamFromResource (TestHelper.GetResourcePath (@"TryToLoad\ThreadedComment.xlsx"));
                    using var wb = new XLWorkbook (stream);
                        AssertComment (expectedComments, wb);

                    using var ms1 = new MemoryStream ();
                    wb.SaveAs (ms1, true);

                    ThreadedCommentsLoadOptions.ThreadedCommentLoading = ThreadedCommentLoading.Skip;
                    using var wb2 = new XLWorkbook (ms1);
                    AssertComment (expectedComments, wb2);

                    using var ms2 = new MemoryStream ();
                    wb.SaveAs (ms2, true);

                    static void AssertComment (string expectedComments, XLWorkbook wb)
                    {
                        var ws = wb.Worksheets.First ();
                        var c1 = ws.Cell ("A1");
                        Assert.AreEqual (expectedComments, c1.GetComment ().Text);
                        Assert.AreEqual ("tc={49C52447-16DF-491E-8BD1-273F700714C6}", c1.GetComment ().Author);

                        var c2 = ws.Cell ("A2");
                        Assert.AreEqual ("Author:\r\nA note", c2.GetComment ().Text);
                        Assert.AreEqual ("tc={49C52447-16DF-491E-8BD1-273F700714C6}", c1.GetComment ().Author);
                    }
                });
            }
            finally
            {
                ThreadedCommentsLoadOptions.ThreadedCommentLoading = ThreadedCommentLoading.Skip;
            }
        }

        [Test]
        public void CanRemoveCommentsWithoutAddingOthers_Regression ()
        {
            Assert.Multiple (() =>
            {
                using (var stream = new MemoryStream ())
                {
                    // arange
                    using (var wb = new XLWorkbook ())
                    {
                        var sheet = wb.AddWorksheet ("sheet1");

                        var a1 = sheet.Cell ("A1");
                        var b5 = sheet.Cell ("B5");

                        a1.SetValue ("test a1");
                        b5.SetValue ("test b5");

                        a1.GetComment().AddText ("no comment");

                        var cellsWithComments3 = wb.Worksheets.SelectMany (_ => _.CellsUsed (XLCellsUsedOptions.Comments)).ToArray ();

                        Assert.That (cellsWithComments3.Length, Is.EqualTo (1));

                        wb.SaveAs (stream, true);
                    }

                    stream.Position = 0;

                    using (var wb = new XLWorkbook (stream))
                    {
                        var cellsWithComments = wb.Worksheets.SelectMany (_ => _.CellsUsed (XLCellsUsedOptions.Comments)).ToArray ();

                        Assert.That (cellsWithComments.Length, Is.EqualTo (1));

                        cellsWithComments.ForEach (_ => _.Clear (XLClearOptions.Comments));

                        wb.Save ();
                    }

                    // assert
                    stream.Position = 0;

                    using (var wb = new XLWorkbook (stream))
                    {
                        var cellsWithComments = wb.Worksheets.SelectMany (_ => _.CellsUsed (XLCellsUsedOptions.Comments)).ToArray ();

                        // BUG? when adding a1.SetValue ("test a1"); this will return at least one cell instead of none
                        Assert.That (cellsWithComments, Is.Empty);
                    }
                }
            });
        }
    }
}
