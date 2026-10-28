using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using static ClosedXML.Excel.XmlMapImporter;

namespace ClosedXML.Excel
{
    /// <summary>
    /// Exports the data of the cells and tables of an XML map as XML: elements are written
    /// in the order of the schema, empty optional elements and attributes are left out and values are formatted by <see cref="XmlDataTypeConverter"/>.
    /// </summary>
    internal static class XmlMapExporter
    {
        private const string XsiNamespace = "http://www.w3.org/2001/XMLSchema-instance";

        /// <summary>
        /// Can the data of the map be exported? Maps whose structure can't be restored from the cells can't be exported,
        /// see <see cref="GetNotExportableReason"/>.
        /// </summary>
        internal static bool IsExportable(XLXmlMap map) => GetNotExportableReason(map) is null;

        /// <returns>The reason why the map can't be exported or <c>null</c>, if it can be exported.</returns>
        internal static string? GetNotExportableReason(XLXmlMap map) => Analyze(map).Reason;

        internal static XLXmlExportResult Export(XLXmlMap map, out string xml)
        {
            var model = Analyze(map);
            if (model.Reason is not null)
                throw new InvalidOperationException($"The XML map '{map.Name}' can't be exported: {model.Reason}");

            // A map without mapped single cells needs at least one table with data rows.
            if (model.Root is null || model.AllNodes.All(n => n.Cell is null && (n.Table is null || !n.Table.Rows().Any())))
                throw new InvalidOperationException($"The XML map '{map.Name}' can't be exported, because it contains no data.");

            xml = Serialize(model, Build(model, model.Root, null)!);
            var document = new XmlDocument { XmlResolver = null };
            document.LoadXml(xml);
            return model.Schema.Validate(document) ? XLXmlExportResult.Success : XLXmlExportResult.ValidationFailed;
        }

        private static Model Analyze(XLXmlMap map)
        {
            var schema = XmlMapSchema.Create(map);
            var namespaces = CreateNamespaceManager(new NameTable(), map.SelectionNamespaces);
            var model = new Model(map, schema, namespaces);
            if (!schema.IsAvailable)
                return model.Fail("the schema of the map can't be used.");

            foreach (var binding in map.Store.GetCellBindings(map))
            {
                var node = model.Add(XPathStep.Parse(binding.XPath));
                if (node is null)
                    return model.Fail($"the XPath '{binding.XPath}' isn't part of the schema.");

                node.Cell = binding;
            }

            var rowNodes = new HashSet<Node>();
            foreach (var (table, binding) in map.Store.GetTableBindings(map))
            {
                var fieldNames = table.Fields.OrderBy(f => f.Index).Select(f => f.Name).ToList();
                var columns = binding.Columns
                    .Where(c => c.Map == map)
                    .Select(c => (Index: c.ResolveIndex(fieldNames), c.XPath, c.XmlDataType, Steps: XPathStep.Parse(c.XPath)))
                    .Where(c => c.Index >= 0 && c.Steps.Count > 0)
                    .ToList();
                if (columns.Count == 0)
                    continue;

                var levels = GetRowLevels(schema, namespaces, columns.Select(c => c.Steps).ToList());
                if (levels.Count != 1)
                    return model.Fail($"the table '{table.Name}' contains a list of lists.");

                var rowPath = levels[0];
                var rowNode = model.Add(rowPath);
                if (rowNode is null)
                    return model.Fail($"the rows of the table '{table.Name}' aren't part of the schema.");

                // Nonrepeating siblings under the same repeating parent can't be in different tables.
                if (!rowNodes.Add(rowNode))
                    return model.Fail($"elements of the same repeating element are mapped to different tables.");

                rowNode.Table = new TableSource(table);
                foreach (var column in columns)
                {
                    if (column.Steps.Count < rowPath.Count || !column.Steps.Take(rowPath.Count).SequenceEqual(rowPath))
                        return model.Fail($"the column '{fieldNames[column.Index]}' of the table '{table.Name}' contains denormalized data.");

                    var node = model.Add(column.Steps);
                    if (node is null)
                        return model.Fail($"the XPath '{column.XPath}' isn't part of the schema.");

                    node.Column = column.Index;
                    node.ColumnDataType = column.XmlDataType;
                }
            }

            foreach (var node in model.AllNodes)
            {
                if (node.Schema is { InChoice: true })
                    return model.Fail($"the element '{node.Name}' is part of a choice.");
                if (node.Schema is { InRepeatingGroup: true })
                    return model.Fail($"the element '{node.Name}' is part of a repeating group.");
                if (node.Schema is { Repeating: true } && node.Table is null && node.HasMappedDescendant())
                    return model.Fail($"the repeating element '{node.Name}' isn't the row of a table.");
            }

            return model;
        }

        /// <summary>
        /// Build the element of a node, <c>null</c> if it is optional and has no data.
        /// </summary>
        /// <param name="model">The analyzed map.</param>
        /// <param name="node">The node to build.</param>
        /// <param name="row">The row of the table of the innermost row element, if any.</param>
        private static XmlOutElement? Build(Model model, Node node, TableRow? row)
        {
            var element = new XmlOutElement(node.Name);
            if (node.Cell is not null || node.Column is not null)
                element.Text = GetText(model, node, row);

            foreach (var (name, required) in node.Schema is { } step ? XmlMapSchema.GetAttributes(step.Element) : new List<(XmlQualifiedName, bool)>())
            {
                var attribute = node.Children.FirstOrDefault(c => c.IsAttribute && c.Name == name);
                if (attribute is null)
                    continue;

                var text = GetText(model, attribute, row);
                if (text is not null || required)
                    element.Attributes.Add((name, text ?? string.Empty));
            }

            var childNames = node.Schema is { } schemaStep ? XmlMapSchema.GetChildElementNames(schemaStep.Element) : new List<XmlQualifiedName>();
            foreach (var child in node.Children.Where(c => !c.IsAttribute).OrderBy(c => childNames.IndexOf(c.Name)))
            {
                if (child.Table is { } table)
                {
                    foreach (var tableRow in table.Rows())
                    {
                        if (Build(model, child, tableRow) is { } rowElement)
                            element.Children.Add(rowElement);
                    }
                }
                else if (Build(model, child, row) is { } childElement)
                {
                    element.Children.Add(childElement);
                }
            }

            if (element.Text is not null || element.Attributes.Count > 0 || element.Children.Count > 0 || node.Parent is null)
                return element;

            // Elements containing rows of tables are written even when the rows are empty, a row element too, unless
            // it is mapped itself. An empty required element is written empty or as nil.
            if (node.HasTableRows(row) && (node.Table is null || (node.Cell is null && node.Column is null)))
                return element;

            if (node.Schema is not { Optional: false } requiredElement)
                return null;

            if (requiredElement.Element.IsNillable && (node.Cell is not null || node.Column is not null))
                element.Attributes.Add((new XmlQualifiedName("nil", XsiNamespace), "true"));

            return element;
        }

        private static string? GetText(Model model, Node node, TableRow? row)
        {
            IXLCell? cell = null;
            if (node.Cell is { } binding)
                cell = binding.Worksheet.Cell(binding.Cell.RangeAddress.FirstAddress.RowNumber, binding.Cell.RangeAddress.FirstAddress.ColumnNumber);
            else if (node.Column is { } column && row is not null)
                cell = row.Table.Table.Worksheet.Cell(row.RowNumber, row.Table.Table.RangeAddress.FirstAddress.ColumnNumber + column);

            if (cell is null)
                return null;

            var xmlDataType = node.Cell is not null ? node.Cell.XmlDataType : node.ColumnDataType;
            return XmlDataTypeConverter.ToXmlText(GetValue(cell), xmlDataType, model.Map.Store.Workbook.Use1904DateSystem);
        }

        private static XLCellValue GetValue(IXLCell cell)
        {
            try
            {
                return cell.Value;
            }
            catch (Exception e) when (e is not OutOfMemoryException)
            {
                // ClosedXML can't evaluate all formulas, e.g. structured references. The cached value is exported.
                return ((XLCell)cell).CachedValue;
            }
        }

        /// <summary>
        /// Write the XML with declaration, indented by tabs, lines separated by CR LF and namespaces declared on the root element.
        /// </summary>
        private static string Serialize(Model model, XmlOutElement root)
        {
            var used = new List<string>();
            void CollectNamespaces(XmlOutElement element)
            {
                if (element.Name.Namespace.Length > 0 && !used.Contains(element.Name.Namespace))
                    used.Add(element.Name.Namespace);
                foreach (var (name, _) in element.Attributes)
                {
                    if (name.Namespace.Length > 0 && !used.Contains(name.Namespace))
                        used.Add(name.Namespace);
                }

                foreach (var child in element.Children)
                    CollectNamespaces(child);
            }

            CollectNamespaces(root);
            if (model.AllNodes.Any(n => n.Schema is { Element.IsNillable: true }) && !used.Contains(XsiNamespace))
                used.Add(XsiNamespace);

            var prefixes = new Dictionary<string, string>();
            foreach (var ns in used)
                prefixes[ns] = ns == XsiNamespace ? "xsi" : model.Namespaces.LookupPrefix(ns) ?? $"ns{prefixes.Count + 1}";

            var xml = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n");
            var declarations = string.Concat(used.Select(ns => $" xmlns:{prefixes[ns]}=\"{Escape(ns, true)}\""));
            WriteElement(xml, root, 0, prefixes, declarations);

            // The last line isn't ended.
            return xml.ToString(0, xml.Length - 2);
        }

        private static void WriteElement(StringBuilder xml, XmlOutElement element, int depth, Dictionary<string, string> prefixes, string declarations)
        {
            string QualifiedName(XmlQualifiedName name) => name.Namespace.Length == 0 ? name.Name : prefixes[name.Namespace] + ":" + name.Name;

            var indent = new string('\t', depth);
            var name = QualifiedName(element.Name);
            xml.Append(indent).Append('<').Append(name).Append(declarations);
            foreach (var (attributeName, value) in element.Attributes)
                xml.Append(' ').Append(QualifiedName(attributeName)).Append("=\"").Append(Escape(value, true)).Append('"');

            if (element.Children.Count == 0 && string.IsNullOrEmpty(element.Text))
            {
                xml.Append("/>\r\n");
                return;
            }

            xml.Append('>');
            if (element.Children.Count == 0)
            {
                xml.Append(Escape(element.Text!, false)).Append("</").Append(name).Append(">\r\n");
                return;
            }

            xml.Append("\r\n");
            foreach (var child in element.Children)
                WriteElement(xml, child, depth + 1, prefixes, string.Empty);
            xml.Append(indent).Append("</").Append(name).Append(">\r\n");
        }

        private static string Escape(string text, bool attribute)
        {
            var escaped = text.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;");
            return attribute ? escaped.Replace("\"", "&quot;") : escaped;
        }

        /// <summary>
        /// The mapped elements and attributes of a map as a tree, with the schema information of each element.
        /// </summary>
        private sealed class Model
        {
            internal Model(XLXmlMap map, XmlMapSchema schema, XmlNamespaceManager namespaces)
            {
                Map = map;
                Schema = schema;
                Namespaces = namespaces;
            }

            internal XLXmlMap Map { get; }

            internal XmlMapSchema Schema { get; }

            internal XmlNamespaceManager Namespaces { get; }

            internal Node? Root { get; private set; }

            internal string? Reason { get; private set; }

            internal IEnumerable<Node> AllNodes => Root?.Descendants() ?? Enumerable.Empty<Node>();

            internal Model Fail(string reason)
            {
                Reason = reason;
                return this;
            }

            /// <summary>
            /// Add the nodes of a path.
            /// </summary>
            /// <returns>The node of the last step or <c>null</c>, if the path isn't in the schema.</returns>
            internal Node? Add(List<XPathStep> steps)
            {
                var elementNames = steps.Where(s => !s.IsAttribute).Select(s => s.ToQualifiedName(Namespaces)).ToList();
                var schemaSteps = Schema.ResolvePath(elementNames);
                if (schemaSteps is null || steps.Count == 0 || steps[0].IsAttribute)
                    return null;

                Root ??= new Node(elementNames[0], false, null, schemaSteps[0]);
                if (Root.Name != elementNames[0])
                    return null;

                var node = Root;
                for (var i = 1; i < steps.Count; i++)
                {
                    var name = steps[i].ToQualifiedName(Namespaces);
                    var isAttribute = steps[i].IsAttribute;
                    var child = node.Children.FirstOrDefault(c => c.Name == name && c.IsAttribute == isAttribute);
                    if (child is null)
                    {
                        child = new Node(name, isAttribute, node, isAttribute ? null : schemaSteps[i]);
                        node.Children.Add(child);
                    }

                    node = child;
                }

                return node;
            }
        }

        private sealed class Node
        {
            internal Node(XmlQualifiedName name, bool isAttribute, Node? parent, SchemaStep? schema)
            {
                Name = name;
                IsAttribute = isAttribute;
                Parent = parent;
                Schema = schema;
            }

            internal XmlQualifiedName Name { get; }

            internal bool IsAttribute { get; }

            internal Node? Parent { get; }

            internal SchemaStep? Schema { get; }

            internal List<Node> Children { get; } = new();

            internal XLXmlCellBinding? Cell { get; set; }

            /// <summary>
            /// Index of the mapped column in the table of the innermost row element.
            /// </summary>
            internal int? Column { get; set; }

            internal string? ColumnDataType { get; set; }

            /// <summary>
            /// The table, if this node is the row element of a table.
            /// </summary>
            internal TableSource? Table { get; set; }

            internal IEnumerable<Node> Descendants()
            {
                yield return this;
                foreach (var descendant in Children.SelectMany(c => c.Descendants()))
                    yield return descendant;
            }

            internal bool HasMappedDescendant()
            {
                return Descendants().Any(n => n.Cell is not null || n.Column is not null);
            }

            /// <summary>
            /// Does the node contain rows of a table or is it a row element of a table with rows?
            /// </summary>
            /// <param name="row">The row being built, if the node is inside of a row element.</param>
            internal bool HasTableRows(TableRow? row)
            {
                return (Table is not null && row is not null && row.Table == Table) || Descendants().Any(n => n.Table is not null && n.Table.Rows().Any());
            }
        }

        private sealed class TableSource
        {
            internal TableSource(XLTable table)
            {
                Table = table;
            }

            internal XLTable Table { get; }

            /// <summary>
            /// The data rows of the table, without header and totals row. A table emptied by an import has no data rows,
            /// only the insert row.
            /// </summary>
            internal IEnumerable<TableRow> Rows()
            {
                if (Table.Worksheet.Workbook.XmlMapStoreOrNull?.GetTableBinding(Table)?.InsertRow == true)
                    yield break;

                var first = Table.RangeAddress.FirstAddress.RowNumber + (Table.ShowHeaderRow ? 1 : 0);
                var last = Table.RangeAddress.LastAddress.RowNumber - (Table.ShowTotalsRow ? 1 : 0);
                for (var row = first; row <= last; row++)
                    yield return new TableRow(this, row);
            }
        }

        private sealed record TableRow(TableSource Table, int RowNumber);

        private sealed class XmlOutElement
        {
            internal XmlOutElement(XmlQualifiedName name)
            {
                Name = name;
            }

            internal XmlQualifiedName Name { get; }

            internal string? Text { get; set; }

            internal List<(XmlQualifiedName Name, string Value)> Attributes { get; } = new();

            internal List<XmlOutElement> Children { get; } = new();
        }
    }
}
