using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;

namespace ClosedXML.Excel
{
    /// <summary>
    /// All XML map data of a workbook: maps, schemas and mappings of cells and tables.
    /// The store is the source of truth for saving, the XML map parts of a loaded package are always regenerated.
    /// </summary>
    internal sealed class XLXmlMapStore
    {
        private readonly List<XLXmlMap> _maps = new();
        private readonly List<XLXmlCellBinding> _cellBindings = new();
        private readonly ConditionalWeakTable<XLTable, XLXmlTableBinding> _tableBindings = new();

        internal XLXmlMapStore(XLWorkbook workbook)
        {
            Workbook = workbook;
            Maps = new XLXmlMaps(this);
        }

        internal XLWorkbook Workbook { get; }

        internal XLXmlMaps Maps { get; }

        internal IReadOnlyList<XLXmlMap> MapList => _maps;

        internal string SelectionNamespaces { get; set; } = string.Empty;

        internal List<XLXmlSchema> Schemas { get; } = new();

        internal void AddMap(XLXmlMap map)
        {
            if (_maps.Any(m => m.Id == map.Id))
                throw new ArgumentException($"An XML map with id {map.Id} already exists.");

            _maps.Add(map);
        }

        internal XLXmlMap? GetMap(uint id) => _maps.FirstOrDefault(m => m.Id == id);

        internal void AddCellBinding(XLXmlCellBinding binding) => _cellBindings.Add(binding);

        internal IEnumerable<XLXmlCellBinding> GetCellBindings(XLWorksheet worksheet)
        {
            return _cellBindings.Where(b => ReferenceEquals(b.Worksheet, worksheet) && b.Cell.RangeAddress.IsValid);
        }

        internal IEnumerable<XLXmlCellBinding> GetCellBindings(XLXmlMap map)
        {
            return _cellBindings.Where(b => b.Map == map && b.Cell.RangeAddress.IsValid && IsLive(b.Worksheet));
        }

        internal void SetTableBinding(XLTable table, XLXmlTableBinding binding)
        {
            _tableBindings.Remove(table);
            _tableBindings.Add(table, binding);
        }

        internal XLXmlTableBinding? GetTableBinding(XLTable table)
        {
            return _tableBindings.TryGetValue(table, out var binding) ? binding : null;
        }

        /// <summary>
        /// Tables of live worksheets that have at least one column mapped to the <paramref name="map"/>.
        /// </summary>
        internal IEnumerable<(XLTable Table, XLXmlTableBinding Binding)> GetTableBindings(XLXmlMap map)
        {
            foreach (var worksheet in Workbook.WorksheetsInternal)
            {
                foreach (var table in worksheet.Tables.Cast<XLTable>())
                {
                    var binding = GetTableBinding(table);
                    if (binding is not null && binding.Columns.Any(c => c.Map == map))
                        yield return (table, binding);
                }
            }
        }

        /// <summary>
        /// Remove bindings of cells that no longer exist, e.g. because the worksheet or the cell was deleted.
        /// </summary>
        internal void RemoveDeadBindings()
        {
            _cellBindings.RemoveAll(b => !b.Cell.RangeAddress.IsValid || !IsLive(b.Worksheet));
        }

        internal void Delete(XLXmlMap map)
        {
            if (!_maps.Remove(map))
                return;

            _cellBindings.RemoveAll(b => b.Map == map);
            foreach (var worksheet in Workbook.WorksheetsInternal)
            {
                foreach (var table in worksheet.Tables.Cast<XLTable>())
                {
                    var binding = GetTableBinding(table);
                    if (binding is null)
                        continue;

                    binding.Columns.RemoveAll(c => c.Map == map);
                    if (binding.Columns.Count == 0)
                        _tableBindings.Remove(table);
                }
            }

            map.IsDeleted = true;
        }

        /// <summary>
        /// Schemas used by the remaining maps, including schemas referenced by them.
        /// </summary>
        internal IEnumerable<XLXmlSchema> GetUsedSchemas()
        {
            var used = new HashSet<string>(_maps.Select(m => m.SchemaId));
            var added = true;
            while (added)
            {
                added = false;
                foreach (var schema in Schemas.Where(s => used.Contains(s.Id) && s.SchemaReference is not null))
                    added |= used.Add(schema.SchemaReference!);
            }

            return Schemas.Where(s => used.Contains(s.Id));
        }

        private bool IsLive(XLWorksheet worksheet)
        {
            return Workbook.WorksheetsInternal.Cast<XLWorksheet>().Any(ws => ReferenceEquals(ws, worksheet));
        }
    }
}
