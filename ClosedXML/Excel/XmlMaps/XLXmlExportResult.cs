namespace ClosedXML.Excel
{
    /// <summary>
    /// Result of an XML export of an <see cref="IXLXmlMap"/>.
    /// </summary>
    public enum XLXmlExportResult
    {
        /// <summary>
        /// The XML data was exported successfully.
        /// </summary>
        Success = 0,

        /// <summary>
        /// The exported XML data doesn't validate against the schema of the map. The data is exported anyway.
        /// </summary>
        ValidationFailed = 1,
    }
}
