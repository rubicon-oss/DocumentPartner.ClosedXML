using DocumentFormat.OpenXml.Spreadsheet;

namespace ClosedXML.Excel
{
    /// <summary>
    /// A single cell mapped to an element or attribute of an XML map.
    /// </summary>
    internal sealed class XLXmlCellBinding
    {
        /// <param name="map">The map the cell is mapped to.</param>
        /// <param name="cell">The mapped cell. The range is tracked by the worksheet, so it moves with inserted or deleted rows and columns.</param>
        /// <param name="properties">The <c>xmlCellPr</c> element with unique name, XPath and data type.</param>
        /// <param name="connectionId">The connection id of the <c>singleXmlCell</c>.</param>
        internal XLXmlCellBinding(XLXmlMap map, XLRange cell, XmlCellProperties properties, uint connectionId)
        {
            Map = map;
            Cell = cell;
            Properties = properties;
            ConnectionId = connectionId;
        }

        internal XLXmlMap Map { get; }

        internal XLRange Cell { get; }

        internal XLWorksheet Worksheet => Cell.Worksheet;

        internal XmlCellProperties Properties { get; }

        internal uint ConnectionId { get; }

        internal string XPath => Properties.XmlProperties?.XPath?.Value ?? string.Empty;

        internal string? XmlDataType => Properties.XmlProperties?.XmlDataType?.InnerText;
    }
}
