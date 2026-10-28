using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ClosedXML.Excel.IO
{
    /// <summary>
    /// Reads mapped single cells (<c>tableSingleCells</c> parts) and XML properties of tables into the
    /// <see cref="XLXmlMapStore"/>. The maps themselves are read by <see cref="XmlMapPartXml"/>.
    /// </summary>
    internal static class XmlMapReader
    {
        internal static void LoadTable(XLXmlMapStore store, XLTable xlTable, Table dTable)
        {
            var binding = new XLXmlTableBinding
            {
                ConnectionId = dTable.ConnectionId?.Value,
                InsertRow = dTable.InsertRow?.Value,
                InsertRowShift = dTable.InsertRowShift?.Value,
                SortState = (SortState?)dTable.SortState?.CloneNode(true),
            };

            var index = 0;
            foreach (var tableColumn in dTable.TableColumns?.Elements<TableColumn>() ?? Enumerable.Empty<TableColumn>())
            {
                var properties = tableColumn.XmlColumnProperties;
                var map = properties?.MapId is not null ? store.GetMap(properties.MapId.Value) : null;
                if (properties is not null && map is not null)
                {
                    binding.Columns.Add(new XLXmlColumnBinding(
                        map,
                        index,
                        GetTableColumnName(tableColumn.Name?.Value ?? string.Empty),
                        tableColumn.UniqueName?.Value,
                        (XmlColumnProperties)properties.CloneNode(true)));
                }
                else if (tableColumn.UniqueName?.Value is { } uniqueName)
                {
                    binding.UnmappedColumns.Add((index, GetTableColumnName(tableColumn.Name?.Value ?? string.Empty), uniqueName));
                }

                index++;
            }

            if (binding.Columns.Count > 0)
                store.SetTableBinding(xlTable, binding);
        }

        internal static void LoadCells(XLXmlMapStore store, XLWorksheet ws, SingleXmlCells singleXmlCells)
        {
            foreach (var singleXmlCell in singleXmlCells.Elements<SingleXmlCell>())
            {
                var reference = singleXmlCell.CellReference?.Value;
                var properties = singleXmlCell.XmlCellProperties;
                var mapId = properties?.XmlProperties?.MapId?.Value;
                var map = mapId is not null ? store.GetMap(mapId.Value) : null;
                if (reference is null || reference.Length == 0 || properties is null || map is null)
                    continue;

                store.AddCellBinding(new XLXmlCellBinding(
                    map,
                    (XLRange)ws.Range(reference)!,
                    (XmlCellProperties)properties.CloneNode(true),
                    singleXmlCell.ConnectionId?.Value ?? 0));
            }
        }

        /// <summary>
        /// Same decoding of table column names as the table loader does.
        /// </summary>
        private static string GetTableColumnName(string name)
        {
            return name.Replace("_x000a_", System.Environment.NewLine).Replace("_x005f_x000a_", "_x000a_");
        }
    }
}
