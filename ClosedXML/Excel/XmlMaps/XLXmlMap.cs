using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace ClosedXML.Excel
{
    internal sealed class XLXmlMap : IXLXmlMap
    {
        private readonly XLXmlMapStore _store;

        internal XLXmlMap(XLXmlMapStore store, uint id, string name, string rootElement, string schemaId)
        {
            _store = store;
            Id = id;
            Name = name;
            RootElement = rootElement;
            SchemaId = schemaId;
        }

        public uint Id { get; }

        public string Name { get; }

        public string RootElement { get; }

        public string SchemaId { get; }

        public string? SchemaNamespace => Schema?.Namespace;

        public string SchemaXml => Schema?.SchemaXml ?? string.Empty;

        public string SelectionNamespaces => _store.SelectionNamespaces;

        public bool ShowImportExportValidationErrors { get; set; }

        public bool AdjustColumnWidth { get; set; } = true;

        public bool AppendOnImport { get; set; }

        public bool PreserveColumnFilter { get; set; } = true;

        public bool PreserveNumberFormatting { get; set; } = true;

        /// <summary>
        /// The <c>DataBinding</c> element of the map, kept as loaded.
        /// </summary>
        internal XElement? DataBinding { get; set; }

        internal XLXmlMapStore Store => _store;

        internal XLXmlSchema? Schema => _store.Schemas.FirstOrDefault(s => s.Id == SchemaId);

        internal bool IsDeleted { get; set; }

        public XLXmlImportResult Import(Stream stream, bool overwrite = true)
        {
            CheckNotDeleted();
            return XmlMapImporter.Import(this, stream, overwrite);
        }

        public XLXmlImportResult Import(string path, bool overwrite = true)
        {
            using var stream = File.OpenRead(path);
            return Import(stream, overwrite);
        }

        public XLXmlImportResult ImportXml(string xml, bool overwrite = true)
        {
            using var reader = new StringReader(xml);
            CheckNotDeleted();
            return XmlMapImporter.Import(this, reader, overwrite);
        }

        public bool IsExportable => !IsDeleted && XmlMapExporter.IsExportable(this);

        public XLXmlExportResult ExportXml(out string xml)
        {
            CheckNotDeleted();
            return XmlMapExporter.Export(this, out xml);
        }

        public XLXmlExportResult Export(Stream stream)
        {
            var result = ExportXml(out var xml);
            var bytes = new System.Text.UTF8Encoding(false).GetBytes(xml);
            stream.Write(bytes, 0, bytes.Length);
            return result;
        }

        public XLXmlExportResult Export(string path, bool overwrite = false)
        {
            // Export first, so no file is created when the map can't be exported.
            var result = ExportXml(out var xml);
            using var stream = new FileStream(path, overwrite ? FileMode.Create : FileMode.CreateNew, FileAccess.Write);
            var bytes = new System.Text.UTF8Encoding(false).GetBytes(xml);
            stream.Write(bytes, 0, bytes.Length);
            return result;
        }

        public void Delete()
        {
            _store.Delete(this);
        }

        private void CheckNotDeleted()
        {
            if (IsDeleted)
                throw new InvalidOperationException($"The XML map '{Name}' has been deleted.");
        }
    }
}
