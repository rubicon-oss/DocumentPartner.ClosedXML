using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Schema;

namespace ClosedXML.Excel
{
    /// <summary>
    /// Compiled schema of an XML map. Used to validate imported XML and to find repeating elements of XML tables.
    /// </summary>
    internal sealed class XmlMapSchema
    {
        private readonly XmlSchemaSet? _schemaSet;

        private XmlMapSchema(XmlSchemaSet? schemaSet)
        {
            _schemaSet = schemaSet;
        }

        internal static XmlMapSchema Create(XLXmlMap map)
        {
            var schemas = map.Store.Schemas;
            var ids = new HashSet<string> { map.SchemaId };
            var queue = new Queue<string>(ids);
            while (queue.Count > 0)
            {
                var id = queue.Dequeue();
                var schema = schemas.FirstOrDefault(s => s.Id == id);
                foreach (var reference in schema?.SchemaReference?.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries) ?? Array.Empty<string>())
                {
                    if (ids.Add(reference))
                        queue.Enqueue(reference);
                }
            }

            try
            {
                var schemaSet = new XmlSchemaSet { XmlResolver = null };
                var settings = new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null };
                foreach (var schema in schemas.Where(s => ids.Contains(s.Id)))
                {
                    using var reader = XmlReader.Create(new StringReader(schema.SchemaXml), settings);
                    schemaSet.Add(XmlSchema.Read(reader, null)!);
                }

                schemaSet.Compile();
                return new XmlMapSchema(schemaSet);
            }
            catch (Exception e) when (e is XmlException or XmlSchemaException)
            {
                return new XmlMapSchema(null);
            }
        }

        /// <summary>
        /// Validate the document against the schema.
        /// </summary>
        /// <returns><c>true</c>, if the document is valid or there is no usable schema.</returns>
        internal bool Validate(XmlDocument document)
        {
            if (_schemaSet is null)
                return true;

            // The validator only warns about a root element without declaration (e.g. in another namespace), it is reported as invalid.
            var root = document.DocumentElement;
            if (root is null || !_schemaSet.GlobalElements.Contains(new XmlQualifiedName(root.LocalName, root.NamespaceURI)))
                return false;

            var valid = true;
            document.Schemas = _schemaSet;
            document.Validate((_, e) =>
            {
                if (e.Severity == XmlSeverityType.Error)
                    valid = false;
            });
            return valid;
        }

        /// <summary>
        /// Is the last element of the <paramref name="path"/> a repeating element, i.e. it can occur more than once in its parent?
        /// </summary>
        /// <param name="path">Qualified names of elements from the root element.</param>
        /// <returns><c>null</c>, if the element can't be found in the schema.</returns>
        internal bool? IsRepeating(IReadOnlyList<XmlQualifiedName> path)
        {
            if (_schemaSet is null || path.Count == 0)
                return null;

            if (_schemaSet.GlobalElements[path[0]] is not XmlSchemaElement element)
                return null;

            var repeating = false;
            for (var i = 1; i < path.Count; i++)
            {
                var particle = (element.ElementSchemaType as XmlSchemaComplexType)?.ContentTypeParticle;
                if (particle is null || !TryFindChild(particle, path[i], false, out element!, out repeating))
                    return null;
            }

            return repeating;
        }

        /// <summary>
        /// Can the schema be used, i.e. could it be compiled?
        /// </summary>
        internal bool IsAvailable => _schemaSet is not null;

        /// <summary>
        /// Find the declarations of the elements of a path and how they are used in the content of their parents.
        /// </summary>
        /// <param name="path">Qualified names of elements from the root element.</param>
        /// <returns>One step for each element of the path or <c>null</c>, if an element can't be found in the schema.</returns>
        internal IReadOnlyList<SchemaStep>? ResolvePath(IReadOnlyList<XmlQualifiedName> path)
        {
            if (_schemaSet is null || path.Count == 0 || _schemaSet.GlobalElements[path[0]] is not XmlSchemaElement root)
                return null;

            var steps = new List<SchemaStep> { new(root, false, false, false, false) };
            for (var i = 1; i < path.Count; i++)
            {
                var particle = (steps[i - 1].Element.ElementSchemaType as XmlSchemaComplexType)?.ContentTypeParticle;
                if (particle is null || FindChild(particle, path[i], default) is not { } step)
                    return null;

                steps.Add(step);
            }

            return steps;
        }

        /// <summary>
        /// Qualified names of the child elements of an element in the order of the schema.
        /// </summary>
        internal static List<XmlQualifiedName> GetChildElementNames(XmlSchemaElement element)
        {
            var names = new List<XmlQualifiedName>();
            void Collect(XmlSchemaParticle? particle)
            {
                switch (particle)
                {
                    case XmlSchemaElement child:
                        if (!names.Contains(child.QualifiedName))
                            names.Add(child.QualifiedName);
                        break;
                    case XmlSchemaGroupBase group:
                        foreach (var item in group.Items.OfType<XmlSchemaParticle>())
                            Collect(item);
                        break;
                    case XmlSchemaGroupRef groupRef:
                        Collect(groupRef.Particle);
                        break;
                }
            }

            Collect((element.ElementSchemaType as XmlSchemaComplexType)?.ContentTypeParticle);
            return names;
        }

        /// <summary>
        /// Qualified names of the attributes of an element in the order of the schema and whether they are required.
        /// </summary>
        internal static List<(XmlQualifiedName Name, bool Required)> GetAttributes(XmlSchemaElement element)
        {
            return (element.ElementSchemaType as XmlSchemaComplexType)?.AttributeUses.Values.Cast<XmlSchemaAttribute>()
                .Select(a => (a.QualifiedName, a.Use == XmlSchemaUse.Required))
                .ToList() ?? new List<(XmlQualifiedName, bool)>();
        }

        private static SchemaStep? FindChild(XmlSchemaParticle particle, XmlQualifiedName name, SchemaStep context)
        {
            switch (particle)
            {
                case XmlSchemaElement candidate when candidate.QualifiedName == name:
                    return new SchemaStep(
                        candidate,
                        context.Optional || candidate.MinOccurs == 0,
                        context.Repeating || candidate.MaxOccurs > 1,
                        context.InChoice,
                        context.InRepeatingGroup);

                case XmlSchemaGroupBase group:
                    var groupContext = new SchemaStep(
                        null!,
                        context.Optional || group.MinOccurs == 0,
                        context.Repeating || group.MaxOccurs > 1,
                        context.InChoice || group is XmlSchemaChoice,
                        context.InRepeatingGroup || (group.MaxOccurs > 1 && group.Items.Count > 1));
                    foreach (var item in group.Items.OfType<XmlSchemaParticle>())
                    {
                        if (FindChild(item, name, groupContext) is { } step)
                            return step;
                    }

                    break;

                case XmlSchemaGroupRef groupRef when groupRef.Particle is not null:
                    var refContext = context with
                    {
                        Optional = context.Optional || groupRef.MinOccurs == 0,
                        Repeating = context.Repeating || groupRef.MaxOccurs > 1,
                    };
                    return FindChild(groupRef.Particle, name, refContext);
            }

            return null;
        }

        private static bool TryFindChild(XmlSchemaParticle particle, XmlQualifiedName name, bool outerRepeating, out XmlSchemaElement? element, out bool repeating)
        {
            switch (particle)
            {
                case XmlSchemaElement candidate when candidate.QualifiedName == name:
                    element = candidate;
                    repeating = outerRepeating || candidate.MaxOccurs > 1;
                    return true;

                case XmlSchemaGroupBase group:
                    var groupRepeating = outerRepeating || group.MaxOccurs > 1;
                    foreach (var item in group.Items.OfType<XmlSchemaParticle>())
                    {
                        if (TryFindChild(item, name, groupRepeating, out element, out repeating))
                            return true;
                    }

                    break;

                case XmlSchemaGroupRef groupRef when groupRef.Particle is not null:
                    return TryFindChild(groupRef.Particle, name, outerRepeating || groupRef.MaxOccurs > 1, out element, out repeating);
            }

            element = null;
            repeating = false;
            return false;
        }
    }

    /// <summary>
    /// An element of a mapped path and how it is used in the content of its parent element.
    /// </summary>
    /// <param name="Element">Declaration of the element.</param>
    /// <param name="Optional">The element may be missing: it or a group it is in has <c>minOccurs="0"</c>.</param>
    /// <param name="Repeating">The element can occur more than once in its parent.</param>
    /// <param name="InChoice">The element is part of an <c>xs:choice</c>.</param>
    /// <param name="InRepeatingGroup">The element is part of a repeating group with more than one particle.</param>
    internal readonly record struct SchemaStep(XmlSchemaElement Element, bool Optional, bool Repeating, bool InChoice, bool InRepeatingGroup);
}
