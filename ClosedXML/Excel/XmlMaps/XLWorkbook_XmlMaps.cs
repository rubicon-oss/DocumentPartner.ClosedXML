using ClosedXML.Excel.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ClosedXML.Excel
{
    /// <summary>
    /// XML map support of the fork. Upstream files only contain one-line calls into this file,
    /// so merging upstream stays simple.
    /// </summary>
    public partial class XLWorkbook
    {
        /// <summary>
        /// Opt-in for XML map support. Default value is <c>false</c>. When disabled, XML maps are neither loaded
        /// nor saved and <see cref="XLWorkbookXmlMapExtensions.XmlMaps"/> throws, so the workbook behaves like
        /// ClosedXML without XML map support.
        /// </summary>
        public static bool EnableXmlMaps { get; set; }

        private XLXmlMapStore? _xmlMapStore;

        /// <summary>
        /// XML map data of the workbook, created on first access.
        /// </summary>
        internal XLXmlMapStore XmlMapStore => _xmlMapStore ??= new XLXmlMapStore(this);

        /// <summary>
        /// XML map data of the workbook or <c>null</c>, if the workbook never had any XML maps.
        /// </summary>
        internal XLXmlMapStore? XmlMapStoreOrNull => _xmlMapStore;

        private void LoadXmlMaps(WorkbookPart workbookPart)
        {
            var mappingsPart = workbookPart.CustomXmlMappingsPart;
            if (mappingsPart is null)
                return;

            XmlMapPartXml.Load(XmlMapStore, mappingsPart);
        }

        private void LoadXmlMapTable(XLTable xlTable, Table dTable)
        {
            if (_xmlMapStore is null)
                return;

            XmlMapReader.LoadTable(_xmlMapStore, xlTable, dTable);
        }

        private void LoadXmlMapCells(WorksheetPart worksheetPart, XLWorksheet ws)
        {
            if (_xmlMapStore is null || worksheetPart.SingleCellTablePart?.SingleXmlCells is not { } singleXmlCells)
                return;

            XmlMapReader.LoadCells(_xmlMapStore, ws, singleXmlCells);
        }
    }
}
