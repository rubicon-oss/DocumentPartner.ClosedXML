using System.Collections.Generic;

namespace ClosedXML.Excel
{
    /// <summary>
    /// XML maps of a workbook.
    /// </summary>
    public interface IXLXmlMaps : IEnumerable<IXLXmlMap>
    {
        /// <summary>
        /// Number of maps in the workbook.
        /// </summary>
        int Count { get; }

        /// <summary>
        /// Get a map by its name.
        /// </summary>
        /// <exception cref="KeyNotFoundException">There is no map with the name.</exception>
        IXLXmlMap Get(string name);

        /// <summary>
        /// Try to get a map by its name.
        /// </summary>
        bool TryGet(string name, out IXLXmlMap map);

        /// <summary>
        /// Remove the map and all its mappings from the workbook. See <see cref="IXLXmlMap.Delete"/>.
        /// </summary>
        void Delete(IXLXmlMap map);
    }
}
