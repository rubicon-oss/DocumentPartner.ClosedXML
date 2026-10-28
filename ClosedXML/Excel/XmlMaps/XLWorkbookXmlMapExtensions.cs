using System;

namespace ClosedXML.Excel
{
    /// <summary>
    /// Access to XML maps of a workbook.
    /// </summary>
    public static class XLWorkbookXmlMapExtensions
    {
        /// <summary>
        /// Get XML maps of the workbook.
        /// </summary>
        /// <exception cref="InvalidOperationException">XML maps are not enabled by <see cref="XLWorkbook.EnableXmlMaps"/>.</exception>
        public static IXLXmlMaps XmlMaps(this IXLWorkbook workbook)
        {
            if (!XLWorkbook.EnableXmlMaps)
                throw new InvalidOperationException($"XML maps are disabled. Set {nameof(XLWorkbook)}.{nameof(XLWorkbook.EnableXmlMaps)} to true to use them.");

            return ((XLWorkbook)workbook).XmlMapStore.Maps;
        }
    }
}
