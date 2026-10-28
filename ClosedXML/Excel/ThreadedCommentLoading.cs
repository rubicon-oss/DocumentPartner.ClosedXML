namespace ClosedXML.Excel
{
    /// <summary>
    /// An enum to define how threaded comments should be handled during loading of spreadsheets.
    /// </summary>
    public enum ThreadedCommentLoading
    {
        /// <summary>
        /// Do not handle threaded comments at all.
        /// </summary>
        Skip,
        /// <summary>
        /// Opt into Converting threaded comments to the older format (also known as "Notes" in MS Excel).
        /// </summary>
        ConvertToNotes
    }
}
