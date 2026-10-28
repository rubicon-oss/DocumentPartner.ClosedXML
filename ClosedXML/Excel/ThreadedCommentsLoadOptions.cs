namespace ClosedXML.Excel
{
    /// <summary>
    /// Opt-in options for loading threaded comments, kept out of <see cref="LoadOptions"/> to avoid merge conflicts with upstream.
    /// </summary>
    public static class ThreadedCommentsLoadOptions
    {
        /// <summary>
        /// Define how threaded comments should be handled during loading of spreadsheets. Default value is <see cref="ThreadedCommentLoading.Skip"/>.
        /// </summary>
        public static ThreadedCommentLoading ThreadedCommentLoading { get; set; } = ThreadedCommentLoading.Skip;
    }
}
