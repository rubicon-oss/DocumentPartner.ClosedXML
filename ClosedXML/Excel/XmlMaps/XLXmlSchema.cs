using System.Xml.Linq;

namespace ClosedXML.Excel
{
    /// <summary>
    /// A schema embedded in the XML maps part. A map references a schema by <see cref="Id"/>.
    /// </summary>
    internal sealed class XLXmlSchema
    {
        /// <param name="id">The <c>ID</c> attribute of the schema.</param>
        /// <param name="schemaReference">The <c>SchemaRef</c> attribute, id of another schema this schema depends on.</param>
        /// <param name="namespace">The <c>Namespace</c> attribute, target namespace of the schema.</param>
        /// <param name="content">The content of the <c>Schema</c> element, usually a single <c>xs:schema</c> element.</param>
        internal XLXmlSchema(string id, string? schemaReference, string? @namespace, XElement[] content)
        {
            Id = id;
            SchemaReference = schemaReference;
            Namespace = @namespace;
            Content = content;
        }

        internal string Id { get; }

        internal string? SchemaReference { get; }

        internal string? Namespace { get; }

        /// <summary>
        /// Content of the schema element, kept as loaded including all namespace declarations.
        /// </summary>
        internal XElement[] Content { get; }

        internal string SchemaXml => string.Concat((object[])Content);
    }
}
