using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace ClosedXML.Excel.IO
{
    /// <summary>
    /// Reads and writes the XML maps part (<c>xl/xmlMaps.xml</c>) as plain XML. The OpenXML SDK DOM
    /// loses namespace declarations of the embedded schemas, so the part is processed with LINQ to XML.
    /// </summary>
    internal static class XmlMapPartXml
    {
        private static readonly XNamespace Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

        internal static void Load(XLXmlMapStore store, CustomXmlMappingsPart part)
        {
            XDocument document;
            using (var stream = part.GetStream(FileMode.Open, FileAccess.Read))
            using (var reader = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null }))
                document = XDocument.Load(reader);

            var mapInfo = document.Root;
            if (mapInfo is null || mapInfo.Name != Main + "MapInfo")
                return;

            store.SelectionNamespaces = (string?)mapInfo.Attribute("SelectionNamespaces") ?? string.Empty;
            foreach (var schema in mapInfo.Elements(Main + "Schema"))
            {
                store.Schemas.Add(new XLXmlSchema(
                    (string?)schema.Attribute("ID") ?? string.Empty,
                    (string?)schema.Attribute("SchemaRef"),
                    (string?)schema.Attribute("Namespace"),
                    schema.Elements().Select(e => new XElement(e)).ToArray()));
            }

            foreach (var element in mapInfo.Elements(Main + "Map"))
            {
                var map = new XLXmlMap(
                    store,
                    (uint?)element.Attribute("ID") ?? 0,
                    (string?)element.Attribute("Name") ?? string.Empty,
                    (string?)element.Attribute("RootElement") ?? string.Empty,
                    (string?)element.Attribute("SchemaID") ?? string.Empty)
                {
                    ShowImportExportValidationErrors = GetBoolean(element, "ShowImportExportValidationErrors"),
                    AdjustColumnWidth = GetBoolean(element, "AutoFit"),
                    AppendOnImport = GetBoolean(element, "Append"),
                    PreserveColumnFilter = GetBoolean(element, "PreserveSortAFLayout"),
                    PreserveNumberFormatting = GetBoolean(element, "PreserveFormat"),
                    DataBinding = element.Element(Main + "DataBinding") is { } dataBinding ? new XElement(dataBinding) : null,
                };
                store.AddMap(map);
            }
        }

        /// <param name="store">The maps to write.</param>
        /// <param name="part">The part to write to.</param>
        /// <param name="hasConnection">Does the package contain a connection with the id? A data binding without its connection isn't written.</param>
        internal static void Write(XLXmlMapStore store, CustomXmlMappingsPart part, Func<uint, bool> hasConnection)
        {
            var mapInfo = new XElement(Main + "MapInfo", new XAttribute("SelectionNamespaces", store.SelectionNamespaces));
            foreach (var schema in store.GetUsedSchemas())
            {
                mapInfo.Add(new XElement(Main + "Schema",
                    new XAttribute("ID", schema.Id),
                    schema.SchemaReference is null ? null : new XAttribute("SchemaRef", schema.SchemaReference),
                    schema.Namespace is null ? null : new XAttribute("Namespace", schema.Namespace),
                    schema.Content.Select(e => new XElement(e))));
            }

            foreach (var map in store.MapList.OrderBy(m => m.Id))
            {
                mapInfo.Add(new XElement(Main + "Map",
                    new XAttribute("ID", map.Id.ToString(CultureInfo.InvariantCulture)),
                    new XAttribute("Name", map.Name),
                    new XAttribute("RootElement", map.RootElement),
                    new XAttribute("SchemaID", map.SchemaId),
                    new XAttribute("ShowImportExportValidationErrors", ToXml(map.ShowImportExportValidationErrors)),
                    new XAttribute("AutoFit", ToXml(map.AdjustColumnWidth)),
                    new XAttribute("Append", ToXml(map.AppendOnImport)),
                    new XAttribute("PreserveSortAFLayout", ToXml(map.PreserveColumnFilter)),
                    new XAttribute("PreserveFormat", ToXml(map.PreserveNumberFormatting)),
                    HasValidConnection(map.DataBinding, hasConnection) ? new XElement(map.DataBinding!) : null));
            }

            using var stream = new MemoryStream();
            using (var writer = XmlWriter.Create(stream, new XmlWriterSettings { Encoding = new System.Text.UTF8Encoding(false) }))
                new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), mapInfo).Save(writer);

            stream.Position = 0;
            part.FeedData(stream);
        }

        /// <summary>
        /// A data binding refers to a connection of <c>connections.xml</c>, which ClosedXML only keeps when it saves into the loaded package.
        /// </summary>
        private static bool HasValidConnection(XElement? dataBinding, Func<uint, bool> hasConnection)
        {
            if (dataBinding is null)
                return false;

            var connectionId = (string?)dataBinding.Attribute("ConnectionID");
            return connectionId is null ||
                   (uint.TryParse(connectionId, NumberStyles.None, CultureInfo.InvariantCulture, out var id) && hasConnection(id));
        }

        private static bool GetBoolean(XElement element, string name)
        {
            var value = (string?)element.Attribute(name);
            return value is not null && (value == "1" || value.Equals("true", StringComparison.OrdinalIgnoreCase));
        }

        private static string ToXml(bool value) => value ? "true" : "false";
    }
}
