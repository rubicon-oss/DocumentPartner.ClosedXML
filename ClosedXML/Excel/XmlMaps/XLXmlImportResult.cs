namespace ClosedXML.Excel
{
    /// <summary>
    /// Result of an XML import into an <see cref="IXLXmlMap"/>.
    /// </summary>
    public enum XLXmlImportResult
    {
        /// <summary>
        /// The XML data was imported successfully.
        /// </summary>
        Success = 0,

        /// <summary>
        /// The XML data was imported, but some data didn't fit into the worksheet and was truncated.
        /// </summary>
        ElementsTruncated = 1,

        /// <summary>
        /// The XML data didn't validate against the schema of the map. The data is imported anyway.
        /// </summary>
        ValidationFailed = 2,
    }
}
