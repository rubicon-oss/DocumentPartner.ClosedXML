using System.IO;

namespace ClosedXML.Excel
{
    /// <summary>
    /// An XML map of a workbook. The map connects an XML schema
    /// with cells and table columns of the workbook, so XML data can be imported into them.
    /// </summary>
    public interface IXLXmlMap
    {
        /// <summary>
        /// Id of the map, unique within the workbook.
        /// </summary>
        uint Id { get; }

        /// <summary>
        /// Name of the map, e.g. <c>root_Map</c>.
        /// </summary>
        string Name { get; }

        /// <summary>
        /// Local name of the root element of the mapped XML.
        /// </summary>
        string RootElement { get; }

        /// <summary>
        /// Id of the schema the map uses.
        /// </summary>
        string SchemaId { get; }

        /// <summary>
        /// Target namespace of the schema the map uses, if any.
        /// </summary>
        string? SchemaNamespace { get; }

        /// <summary>
        /// The XML schema (XSD) embedded in the workbook for this map.
        /// </summary>
        string SchemaXml { get; }

        /// <summary>
        /// Namespace prefixes used by the XPaths of the mappings, e.g. <c>xmlns:ns1='urn:test'</c>.
        /// </summary>
        string SelectionNamespaces { get; }

        /// <summary>
        /// Validate data against the schema on import and export and show validation errors.
        /// </summary>
        bool ShowImportExportValidationErrors { get; set; }

        /// <summary>
        /// Adjust column widths after an import.
        /// </summary>
        bool AdjustColumnWidth { get; set; }

        /// <summary>
        /// Append new data to existing XML tables on import instead of overwriting them.
        /// </summary>
        bool AppendOnImport { get; set; }

        /// <summary>
        /// Preserve column filter and sorting of XML tables on import.
        /// </summary>
        bool PreserveColumnFilter { get; set; }

        /// <summary>
        /// Keep the number formats of mapped cells on import.
        /// </summary>
        bool PreserveNumberFormatting { get; set; }

        /// <summary>
        /// Import XML data from a stream into the cells and tables mapped by this map.
        /// </summary>
        /// <param name="stream">The XML data.</param>
        /// <param name="overwrite"><c>true</c> to replace existing data, <c>false</c> to append rows to XML tables.</param>
        XLXmlImportResult Import(Stream stream, bool overwrite = true);

        /// <summary>
        /// Import XML data from a file into the cells and tables mapped by this map.
        /// </summary>
        /// <param name="path">Path of the XML file.</param>
        /// <param name="overwrite"><c>true</c> to replace existing data, <c>false</c> to append rows to XML tables.</param>
        XLXmlImportResult Import(string path, bool overwrite = true);

        /// <summary>
        /// Import XML data from a string into the cells and tables mapped by this map.
        /// </summary>
        /// <param name="xml">The XML data.</param>
        /// <param name="overwrite"><c>true</c> to replace existing data, <c>false</c> to append rows to XML tables.</param>
        XLXmlImportResult ImportXml(string xml, bool overwrite = true);

        /// <summary>
        /// Can the data of the map be exported? A map can't be exported, if the XML structure can't be restored
        /// from the cells: e.g. a table with a list of lists or denormalized data, or mapped elements inside of an <c>xs:choice</c>.
        /// </summary>
        bool IsExportable { get; }

        /// <summary>
        /// Export the data of the mapped cells and tables as XML.
        /// </summary>
        /// <param name="xml">The exported XML, including the XML declaration.</param>
        /// <exception cref="System.InvalidOperationException">The map isn't exportable or the mapped cells contain no data.</exception>
        XLXmlExportResult ExportXml(out string xml);

        /// <summary>
        /// Export the data of the mapped cells and tables as UTF-8 encoded XML to a stream.
        /// </summary>
        /// <param name="stream">The stream to write to.</param>
        /// <exception cref="System.InvalidOperationException">The map isn't exportable or the mapped cells contain no data.</exception>
        XLXmlExportResult Export(Stream stream);

        /// <summary>
        /// Export the data of the mapped cells and tables as UTF-8 encoded XML to a file.
        /// </summary>
        /// <param name="path">Path of the XML file.</param>
        /// <param name="overwrite"><c>true</c> to replace an existing file.</param>
        /// <exception cref="System.InvalidOperationException">The map isn't exportable or the mapped cells contain no data.</exception>
        /// <exception cref="IOException">The file exists and <paramref name="overwrite"/> is <c>false</c>.</exception>
        XLXmlExportResult Export(string path, bool overwrite = false);

        /// <summary>
        /// Remove the map and all its mappings from the workbook. Values of the cells stay,
        /// XML tables become normal tables.
        /// </summary>
        void Delete();
    }
}
