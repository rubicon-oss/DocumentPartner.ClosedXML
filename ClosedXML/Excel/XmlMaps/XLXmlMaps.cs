using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal sealed class XLXmlMaps : IXLXmlMaps
    {
        private readonly XLXmlMapStore _store;

        internal XLXmlMaps(XLXmlMapStore store)
        {
            _store = store;
        }

        public int Count => _store.MapList.Count;

        public IXLXmlMap Get(string name)
        {
            if (!TryGet(name, out var map))
                throw new KeyNotFoundException($"There is no XML map with name '{name}'.");

            return map;
        }

        public bool TryGet(string name, out IXLXmlMap map)
        {
            map = _store.MapList.FirstOrDefault(m => string.Equals(m.Name, name, StringComparison.OrdinalIgnoreCase))!;
            return map is not null;
        }

        public void Delete(IXLXmlMap map)
        {
            _store.Delete((XLXmlMap)map);
        }

        public IEnumerator<IXLXmlMap> GetEnumerator() => _store.MapList.Cast<IXLXmlMap>().GetEnumerator();

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
