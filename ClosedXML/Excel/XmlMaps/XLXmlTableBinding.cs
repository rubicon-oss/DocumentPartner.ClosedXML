using System.Collections.Generic;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ClosedXML.Excel
{
    /// <summary>
    /// XML map related properties of a table that ClosedXML doesn't keep in <see cref="XLTable"/>.
    /// </summary>
    internal sealed class XLXmlTableBinding
    {
        internal uint? ConnectionId { get; set; }

        internal bool? InsertRow { get; set; }

        internal bool? InsertRowShift { get; set; }

        /// <summary>
        /// The sort state of the table. ClosedXML doesn't keep the sort state of tables, it is kept for XML tables,
        /// because an import keeps it.
        /// </summary>
        internal SortState? SortState { get; set; }

        internal List<XLXmlColumnBinding> Columns { get; } = new();

        /// <summary>
        /// The <c>uniqueName</c> of columns that aren't mapped (e.g. calculated columns). Every column of an XML table has a unique name.
        /// </summary>
        internal List<(int Index, string Name, string UniqueName)> UnmappedColumns { get; } = new();
    }

    /// <summary>
    /// A table column mapped to a repeating element or attribute of an XML map.
    /// </summary>
    internal sealed class XLXmlColumnBinding
    {
        /// <param name="map">The map the column is mapped to.</param>
        /// <param name="index">Zero-based index of the column in the table when loaded.</param>
        /// <param name="name">Name of the column when loaded.</param>
        /// <param name="uniqueName">The <c>uniqueName</c> attribute of the table column.</param>
        /// <param name="properties">The <c>xmlColumnPr</c> element with XPath and data type.</param>
        internal XLXmlColumnBinding(XLXmlMap map, int index, string name, string? uniqueName, XmlColumnProperties properties)
        {
            Map = map;
            Index = index;
            Name = name;
            UniqueName = uniqueName;
            Properties = properties;
        }

        internal XLXmlMap Map { get; }

        internal int Index { get; }

        internal string Name { get; }

        internal string? UniqueName { get; }

        internal XmlColumnProperties Properties { get; }

        internal string XPath => Properties.XPath?.Value ?? string.Empty;

        internal string? XmlDataType => Properties.XmlDataType?.InnerText;

        internal bool Denormalized => Properties.Denormalized?.Value ?? false;

        /// <summary>
        /// Style of data cells of the column. The column format is applied to rows filled on import,
        /// so the style is kept even if the rows of the table were cleared by an earlier import.
        /// </summary>
        internal XLStyleValue? DataStyle { get; set; }

        /// <summary>
        /// Find the current index of the column in the table. Columns are matched by name first,
        /// then by the index at load time.
        /// </summary>
        /// <returns>Zero-based index or -1, if the column no longer exists.</returns>
        internal int ResolveIndex(IReadOnlyList<string> fieldNames) => ResolveIndex(Name, Index, fieldNames);

        internal static int ResolveIndex(string name, int index, IReadOnlyList<string> fieldNames)
        {
            for (var i = 0; i < fieldNames.Count; i++)
            {
                if (string.Equals(fieldNames[i], name, System.StringComparison.OrdinalIgnoreCase))
                    return i;
            }

            return index < fieldNames.Count ? index : -1;
        }
    }
}
