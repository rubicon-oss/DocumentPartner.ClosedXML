using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using ClosedXML.XmlMaps.Tests.Infrastructure;

namespace ClosedXML.XmlMaps.Tests.Scenarios
{
    /// <summary>
    /// Templates with XML maps and XML data for them. The templates are built by Excel, so they contain
    /// exactly what Excel writes for the mappings (e.g. number formats Excel sets when a cell is mapped).
    /// </summary>
    internal static class Scenarios
    {
        internal const string DataNamespace = "urn:scenario";
        internal const string DataSelection = "xmlns:ns1='urn:scenario'";
        internal const string DataMap = "data_Map";
        internal const string OtherNamespace = "urn:other";
        internal const string OtherSelection = "xmlns:ns2='urn:other'";
        internal const string OtherMap = "other_Map";

        /// <summary>
        /// XSD built-in types mapped to single cells of the <see cref="Types"/> scenario.
        /// </summary>
        internal static readonly string[] TypeNames =
        {
            "string", "normalizedString", "token", "int", "integer", "positiveInteger", "negativeInteger",
            "nonPositiveInteger", "nonNegativeInteger", "long", "unsignedLong", "short", "unsignedShort", "byte",
            "unsignedByte", "unsignedInt", "decimal", "float", "double", "boolean", "date", "dateTime", "time",
            "duration", "gYear", "gYearMonth", "gMonth", "gMonthDay", "gDay", "base64Binary", "hexBinary", "anyURI",
            "language", "NMTOKEN", "Name", "NCName",
        };

        /// <summary>
        /// Real-world workbook with pivot tables and charts. ClosedXML doesn't keep the output cells of
        /// pivot tables, so only the mapped cells are compared.
        /// </summary>
        internal static readonly XmlMapScenario RenderingTest = XmlMapScenario.FromResource("RenderingTest.xlsx", compareAllCells: false);

        /// <summary>
        /// Every type of <see cref="TypeNames"/> mapped to a single cell in column B, label in column A.
        /// </summary>
        internal static readonly XmlMapScenario Types = XmlMapScenario.FromExcel(nameof(Types), b =>
        {
            var map = b.AddMap(DataSchema, "data");
            for (var i = 0; i < TypeNames.Length; i++)
            {
                b.Set("Types", $"A{i + 1}", TypeNames[i]);
                b.MapCell("Types", $"B{i + 1}", map, $"/ns1:data/ns1:v_{TypeNames[i]}", DataSelection);
            }

            b.Set("Types", $"A{TypeNames.Length + 1}", "nil");
            b.MapCell("Types", $"B{TypeNames.Length + 1}", map, "/ns1:data/ns1:nil", DataSelection);
        });

        /// <summary>
        /// Same as <see cref="Types"/>, but the workbook uses the 1904 date system.
        /// </summary>
        internal static readonly XmlMapScenario Types1904 = XmlMapScenario.FromExcel(nameof(Types1904), b =>
        {
            b.Workbook.Date1904 = true;
            Types.Build!(b);
        });

        /// <summary>
        /// A table with element and attribute columns and cells below and right of it,
        /// and a second table with a totals row.
        /// </summary>
        internal static readonly XmlMapScenario Tables = XmlMapScenario.FromExcel(nameof(Tables), b =>
        {
            var map = b.AddMap(DataSchema, "data");
            b.MapTable("Tables", "B2:F3", map, DataSelection,
                "/ns1:data/ns1:rows/ns1:row/@id",
                "/ns1:data/ns1:rows/ns1:row/ns1:a",
                "/ns1:data/ns1:rows/ns1:row/ns1:d",
                "/ns1:data/ns1:rows/ns1:row/ns1:n",
                "/ns1:data/ns1:rows/ns1:row/ns1:b");
            b.Set("Tables", "B6", "below B6");
            b.Set("Tables", "D12", "below D12");
            b.Set("Tables", "H3", "right H3");
            b.Set("Tables", "A5", "left A5");

            var items = b.MapTable("Tables", "J2:K3", map, DataSelection,
                "/ns1:data/ns1:items/ns1:item/ns1:name",
                "/ns1:data/ns1:items/ns1:item/ns1:qty");
            items.ShowTotals = true;
            b.Set("Tables", "J8", "below J8");
        });

        /// <summary>
        /// Two maps with different schemas on two sheets.
        /// </summary>
        internal static readonly XmlMapScenario TwoMaps = XmlMapScenario.FromExcel(nameof(TwoMaps), b =>
        {
            var data = b.AddMap(DataSchema, "data");
            b.MapCell("First", "B1", data, "/ns1:data/ns1:v_string", DataSelection);
            b.MapCell("First", "B2", data, "/ns1:data/ns1:v_int", DataSelection);
            b.MapTable("First", "D1:E2", data, DataSelection,
                "/ns1:data/ns1:rows/ns1:row/ns1:a",
                "/ns1:data/ns1:rows/ns1:row/ns1:n");

            var other = b.AddMap(OtherSchema, "other");
            b.MapCell("Second", "B1", other, "/ns2:other/ns2:title", OtherSelection);
            b.MapTable("Second", "D1:E2", other, OtherSelection,
                "/ns2:other/ns2:entries/ns2:entry/ns2:key",
                "/ns2:other/ns2:entries/ns2:entry/ns2:value");
        });

        /// <summary>
        /// Mapped cells with number formats set by the user after mapping.
        /// </summary>
        internal static readonly XmlMapScenario Formats = XmlMapScenario.FromExcel(nameof(Formats), b => BuildFormats(b, true));

        /// <summary>
        /// Same as <see cref="Formats"/>, but the map doesn't preserve number formats.
        /// </summary>
        internal static readonly XmlMapScenario FormatsNotPreserved = XmlMapScenario.FromExcel(nameof(FormatsNotPreserved), b => BuildFormats(b, false));

        /// <summary>
        /// Same as <see cref="Types"/>, but the formats Excel sets when mapping are reset to General afterwards.
        /// Shows which formats Excel applies on import.
        /// </summary>
        internal static readonly XmlMapScenario TypesGeneral = XmlMapScenario.FromExcel(nameof(TypesGeneral), b =>
        {
            Types.Build!(b);
            b.Format("Types", $"B1:B{TypeNames.Length + 1}", "General");
        });

        /// <summary>
        /// Single elements of the <see cref="Advanced"/> scenario in schema order, mapped to column B.
        /// <c>@currency</c> is the attribute of <c>amount</c>.
        /// </summary>
        internal static readonly string[] AdvancedSingles =
        {
            "title", "color", "level", "percent", "code", "due", "numbers", "intOrText", "amount", "@currency", "email", "phone",
        };

        /// <summary>
        /// Schema features beyond built-in types (enumerations, facets, list and union types, references, choice, all,
        /// simple content with attributes) and an XML table with calculated columns and a totals row.
        /// </summary>
        internal static readonly XmlMapScenario Advanced = XmlMapScenario.FromExcel(nameof(Advanced), b =>
        {
            var map = b.AddMap(AdvancedSchema, "order");
            for (var i = 0; i < AdvancedSingles.Length; i++)
            {
                b.Set("Order", $"A{i + 1}", AdvancedSingles[i]);
                b.MapCell("Order", $"B{i + 1}", map, "/ns1:order/" + AdvancedXPath(AdvancedSingles[i]), AdvancedSelection);
            }

            var lines = b.MapTable("Order", "D1:H2", map, AdvancedSelection,
                "/ns1:order/ns1:lines/ns1:line/ns1:product",
                "/ns1:order/ns1:lines/ns1:line/ns1:qty",
                "/ns1:order/ns1:lines/ns1:line/ns1:price",
                "/ns1:order/ns1:lines/ns1:line/ns1:color",
                "/ns1:order/ns1:lines/ns1:line/@status");
            lines.Name = "Lines";
            var names = new[] { "product", "qty", "price", "color", "status" };
            for (var i = 0; i < names.Length; i++)
                lines.ListColumns.Item(i + 1).Name = names[i];

            b.AddCalculatedColumn(lines, 3, "Double", "=[@qty]*2");
            b.AddCalculatedColumn(lines, null, "Total", "=[@qty]*[@price]");
            lines.ShowTotals = true;
            b.SetTotals(lines, "product", ExcelWorkbookBuilder.TotalsCount);
            b.SetTotals(lines, "qty", ExcelWorkbookBuilder.TotalsSum);
            b.SetTotals(lines, "Double", ExcelWorkbookBuilder.TotalsMax);
            b.SetTotals(lines, "price", ExcelWorkbookBuilder.TotalsAverage);
            b.SetTotals(lines, "Total", ExcelWorkbookBuilder.TotalsSum);
            lines.ListColumns.Item("color").Total.Value2 = "Totals";
            b.Set("Order", $"A{AdvancedSingles.Length + 2}", "sum of totals");
            b.Formula("Order", $"B{AdvancedSingles.Length + 2}", "=SUM(Lines[Total])");
            b.Set("Order", "D10", "below D10");
            b.Set("Order", "L2", "right L2");
        });

        /// <summary>
        /// Tables with denormalized columns: elements above the repeating row element (the customer of all orders,
        /// the id and date of an order) are repeated in each row. The second table has a column of an element whose
        /// repeating parent (the order) isn't mapped.
        /// </summary>
        internal static readonly XmlMapScenario Nested = XmlMapScenario.FromExcel(nameof(Nested), b =>
        {
            var map = b.AddMap(NestedSchema, "orders");
            b.MapTable("Nested", "A1:E2", map, NestedSelection,
                "/ns1:orders/ns1:customer",
                "/ns1:orders/ns1:order/@id",
                "/ns1:orders/ns1:order/ns1:date",
                "/ns1:orders/ns1:order/ns1:line/ns1:product",
                "/ns1:orders/ns1:order/ns1:line/ns1:qty");
            b.Set("Nested", "A12", "below A12");
            b.MapTable("Nested", "H1:H2", map, NestedSelection,
                "/ns1:orders/ns1:order/ns1:note/ns1:text");
        });

        internal const string FirstMap = "first_Map";
        internal const string SecondMap = "second_Map";

        /// <summary>
        /// Two maps of the same schema. Each map has its own cells and tables, an import only changes the cells of its map.
        /// </summary>
        internal static readonly XmlMapScenario SharedSchema = XmlMapScenario.FromExcel(nameof(SharedSchema), b =>
        {
            var first = b.AddMap(DataSchema, "data");
            first.Name = FirstMap;
            var second = b.AddMap(DataSchema, "data");
            second.Name = SecondMap;
            b.MapCell("Shared", "B1", first, "/ns1:data/ns1:v_string", DataSelection);
            b.MapCell("Shared", "B2", second, "/ns1:data/ns1:v_string", DataSelection);
            b.MapTable("Shared", "D1:E2", first, DataSelection, "/ns1:data/ns1:rows/ns1:row/ns1:a", "/ns1:data/ns1:rows/ns1:row/ns1:n");
            b.MapTable("Shared", "G1:H2", second, DataSelection, "/ns1:data/ns1:rows/ns1:row/ns1:a", "/ns1:data/ns1:rows/ns1:row/ns1:n");
        });

        /// <summary>
        /// Two maps of schemas without target namespace, so the XPaths have no prefixes.
        /// </summary>
        internal static readonly XmlMapScenario NoNamespace = XmlMapScenario.FromExcel(nameof(NoNamespace), b =>
        {
            var alpha = b.AddMap(NoNamespaceSchema("alpha"), "alpha");
            var beta = b.AddMap(NoNamespaceSchema("beta"), "beta");
            b.MapCell("Plain", "B1", alpha, "/alpha/name", string.Empty);
            b.MapTable("Plain", "D1:E2", alpha, string.Empty, "/alpha/items/item/@id", "/alpha/items/item/value");
            b.MapCell("Plain", "B2", beta, "/beta/name", string.Empty);
            b.MapTable("Plain", "G1:H2", beta, string.Empty, "/beta/items/item/@id", "/beta/items/item/value");
        });

        /// <summary>
        /// One map with single cells and two tables on different sheets.
        /// </summary>
        internal static readonly XmlMapScenario Sheets = XmlMapScenario.FromExcel(nameof(Sheets), b =>
        {
            var map = b.AddMap(DataSchema, "data");
            b.MapCell("Cells", "B1", map, "/ns1:data/ns1:v_string", DataSelection);
            b.MapCell("Cells", "B2", map, "/ns1:data/ns1:v_date", DataSelection);
            b.MapTable("Rows", "A1:B2", map, DataSelection, "/ns1:data/ns1:rows/ns1:row/ns1:a", "/ns1:data/ns1:rows/ns1:row/ns1:n");
            b.MapTable("Items", "C3:D4", map, DataSelection, "/ns1:data/ns1:items/ns1:item/ns1:name", "/ns1:data/ns1:items/ns1:item/ns1:qty");
        });

        /// <summary>
        /// Tables in the same columns above each other: a table of the data map above a table of the other map and a table of
        /// the data map that is only partly in the columns of the table above it.
        /// </summary>
        internal static readonly XmlMapScenario Stacked = XmlMapScenario.FromExcel(nameof(Stacked), b =>
        {
            var data = b.AddMap(DataSchema, "data");
            var other = b.AddMap(OtherSchema, "other");
            b.MapTable("Stacked", "B2:C3", data, DataSelection, "/ns1:data/ns1:rows/ns1:row/ns1:a", "/ns1:data/ns1:rows/ns1:row/ns1:n");
            b.MapTable("Stacked", "B6:C7", other, OtherSelection, "/ns2:other/ns2:entries/ns2:entry/ns2:key", "/ns2:other/ns2:entries/ns2:entry/ns2:value");
            b.MapTable("Stacked", "C12:D13", data, DataSelection, "/ns1:data/ns1:items/ns1:item/ns1:name", "/ns1:data/ns1:items/ns1:item/ns1:qty");
            b.Set("Stacked", "B20", "below B20");
        });

        /// <summary>
        /// Tables that can't grow without breaking something below them: merged cells across the edge of the table columns
        /// and the end of the sheet.
        /// </summary>
        internal static readonly XmlMapScenario Blocked = XmlMapScenario.FromExcel(nameof(Blocked), b =>
        {
            var map = b.AddMap(DataSchema, "data");
            b.MapTable("Blocked", "B2:C3", map, DataSelection, "/ns1:data/ns1:rows/ns1:row/ns1:a", "/ns1:data/ns1:rows/ns1:row/ns1:n");
            b.Set("Blocked", "C6", "merged C6:D6");
            b.Sheet("Blocked").Range("C6:D6").Merge();
            b.MapTable("Blocked", "F1048574:G1048575", map, DataSelection, "/ns1:data/ns1:items/ns1:item/ns1:name", "/ns1:data/ns1:items/ns1:item/ns1:qty");
        });

        /// <summary>
        /// A table with a cell in the way of its growth and content below it that is shifted with the cells:
        /// a table, merged cells and an array formula, all inside of the columns of the table.
        /// </summary>
        internal static readonly XmlMapScenario Shifted = XmlMapScenario.FromExcel(nameof(Shifted), b =>
        {
            var map = b.AddMap(DataSchema, "data");
            b.MapTable("Shifted", "B2:C3", map, DataSelection, "/ns1:data/ns1:rows/ns1:row/ns1:a", "/ns1:data/ns1:rows/ns1:row/ns1:n");
            b.Set("Shifted", "B5", "in the way");
            var ws = b.Sheet("Shifted");
            ws.Range("B10").Value2 = "h1";
            ws.Range("C10").Value2 = "h2";
            ws.Range("B11").Value2 = 1;
            ws.Range("C11").Value2 = 2;
            ws.ListObjects.Add(1, ws.Range("B10:C11"), System.Reflection.Missing.Value, 1);
            ws.Range("B14").Value2 = "merged";
            ws.Range("B14:C14").Merge();
            ws.Range("B16:C16").FormulaArray = "=B11+1";
            b.Set("Shifted", "A20", "left of the table");
        });

        /// <summary>
        /// A table that can't grow, because an array formula is across the edge of the table columns.
        /// </summary>
        internal static readonly XmlMapScenario BlockedArray = XmlMapScenario.FromExcel(nameof(BlockedArray), b =>
        {
            var map = b.AddMap(DataSchema, "data");
            b.MapTable("Blocked", "B2:C3", map, DataSelection, "/ns1:data/ns1:rows/ns1:row/ns1:a", "/ns1:data/ns1:rows/ns1:row/ns1:n");
            b.Set("Blocked", "B5", "in the way");
            b.Sheet("Blocked").Range("C8:D8").FormulaArray = "=1";
        });

        /// <summary>
        /// A table that can't grow, because a cell would be shifted off the sheet.
        /// </summary>
        internal static readonly XmlMapScenario BlockedEnd = XmlMapScenario.FromExcel(nameof(BlockedEnd), b =>
        {
            var map = b.AddMap(DataSchema, "data");
            b.MapTable("Blocked", "B2:C3", map, DataSelection, "/ns1:data/ns1:rows/ns1:row/ns1:a", "/ns1:data/ns1:rows/ns1:row/ns1:n");
            b.Set("Blocked", "B5", "in the way");
            b.Set("Blocked", "C1048576", "last row");
        });

        /// <summary>
        /// A table with a filter and a table with a filter and a sort order.
        /// </summary>
        internal static readonly XmlMapScenario Filtered = XmlMapScenario.FromExcel(nameof(Filtered), b => BuildFiltered(b, true));

        /// <summary>
        /// Same as <see cref="Filtered"/>, but the map doesn't preserve the filters.
        /// </summary>
        internal static readonly XmlMapScenario FilteredNotPreserved = XmlMapScenario.FromExcel(nameof(FilteredNotPreserved), b => BuildFiltered(b, false));

        private static void BuildFiltered(ExcelWorkbookBuilder b, bool preserve)
        {
            var map = b.AddMap(DataSchema, "data");
            map.PreserveColumnFilter = preserve;
            var rows = b.MapTable("Filtered", "B2:C3", map, DataSelection, "/ns1:data/ns1:rows/ns1:row/ns1:a", "/ns1:data/ns1:rows/ns1:row/ns1:n");
            rows.Range.AutoFilter(2, ">3");
            b.Set("Filtered", "A12", "below");

            // Rows are hidden for the whole sheet, the tables are below each other, so they don't hide rows of each other.
            var items = b.MapTable("Filtered", "F20:G21", map, DataSelection, "/ns1:data/ns1:items/ns1:item/ns1:name", "/ns1:data/ns1:items/ns1:item/ns1:qty");
            items.Sort.SortFields.Add(items.ListColumns.Item(2).DataBodyRange, 0, 2);
            items.Sort.Apply();
            items.Range.AutoFilter(2, "<>30");
        }

        internal static readonly string[] UnsupportedKinds = { "any", "anyAttribute", "recursive", "abstract", "mixed" };

        /// <summary>
        /// Schemas with constructs Excel doesn't support for mapping: <c>xs:any</c>, <c>xs:anyAttribute</c>, recursion,
        /// abstract elements with substitution groups and mixed content. The elements Excel can map are mapped.
        /// Each schema has its own namespace, the maps are named after <see cref="UnsupportedKinds"/>.
        /// </summary>
        internal static readonly XmlMapScenario Unsupported = XmlMapScenario.FromExcel(nameof(Unsupported), b =>
        {
            var maps = new Dictionary<string, dynamic>();
            for (var i = 0; i < UnsupportedKinds.Length; i++)
            {
                var kind = UnsupportedKinds[i];
                var map = b.AddMap(UnsupportedSchema(kind), "root");
                map.Name = kind + "_Map";
                var selection = UnsupportedSelection(kind);
                var prefix = $"ns{i + 1}";
                b.MapCell(kind, "B1", map, $"/{prefix}:root/{prefix}:name", selection);
                if (kind == "recursive")
                {
                    b.MapTable(kind, "D1:D2", map, selection, $"/{prefix}:root/{prefix}:node/{prefix}:label");
                    b.MapTable(kind, "F1:F2", map, selection, $"/{prefix}:root/{prefix}:node/{prefix}:node/{prefix}:label");
                }
                else if (kind == "mixed")
                {
                    b.MapTable(kind, "D1:D2", map, selection, $"/{prefix}:root/{prefix}:text/{prefix}:b");
                }
            }
        });

        internal static string UnsupportedNamespace(string kind) => "urn:" + kind;

        /// <summary>
        /// Excel numbers the prefixes of the namespaces of all maps of a workbook.
        /// </summary>
        private static string UnsupportedSelection(string kind) => $"xmlns:ns{System.Array.IndexOf(UnsupportedKinds, kind) + 1}='{UnsupportedNamespace(kind)}'";

        internal static string UnsupportedSchema(string kind)
        {
            const string name = "<xs:element name=\"name\" type=\"xs:string\" minOccurs=\"0\"/>";
            var body = kind switch
            {
                "any" => $"<xs:element name=\"root\"><xs:complexType><xs:sequence>{name}<xs:any namespace=\"urn:strict\" processContents=\"strict\" minOccurs=\"0\"/>" +
                         "<xs:any namespace=\"urn:lax\" processContents=\"lax\" minOccurs=\"0\" maxOccurs=\"unbounded\"/></xs:sequence></xs:complexType></xs:element>",
                "anyAttribute" => $"<xs:element name=\"root\"><xs:complexType><xs:sequence>{name}</xs:sequence><xs:anyAttribute processContents=\"skip\"/></xs:complexType></xs:element>",
                "recursive" => "<xs:complexType name=\"Node\"><xs:sequence><xs:element name=\"label\" type=\"xs:string\" minOccurs=\"0\"/>" +
                               "<xs:element name=\"node\" type=\"Node\" minOccurs=\"0\" maxOccurs=\"unbounded\"/></xs:sequence></xs:complexType>" +
                               $"<xs:element name=\"root\"><xs:complexType><xs:sequence>{name}<xs:element name=\"node\" type=\"Node\" minOccurs=\"0\" maxOccurs=\"unbounded\"/></xs:sequence></xs:complexType></xs:element>",
                "abstract" => "<xs:element name=\"shape\" type=\"xs:string\" abstract=\"true\"/><xs:element name=\"circle\" type=\"xs:string\" substitutionGroup=\"shape\"/>" +
                              $"<xs:element name=\"root\"><xs:complexType><xs:sequence>{name}<xs:element ref=\"shape\" minOccurs=\"0\"/></xs:sequence></xs:complexType></xs:element>",
                "mixed" => $"<xs:element name=\"root\"><xs:complexType><xs:sequence>{name}<xs:element name=\"text\" minOccurs=\"0\"><xs:complexType mixed=\"true\"><xs:sequence>" +
                           "<xs:element name=\"b\" type=\"xs:string\" minOccurs=\"0\" maxOccurs=\"unbounded\"/></xs:sequence></xs:complexType></xs:element></xs:sequence></xs:complexType></xs:element>",
                _ => throw new System.ArgumentOutOfRangeException(nameof(kind)),
            };
            return "<?xml version=\"1.0\" encoding=\"UTF-8\"?><xs:schema xmlns:xs=\"http://www.w3.org/2001/XMLSchema\" " +
                   $"targetNamespace=\"{UnsupportedNamespace(kind)}\" xmlns=\"{UnsupportedNamespace(kind)}\" elementFormDefault=\"qualified\">{body}</xs:schema>";
        }

        /// <summary>
        /// XML data for a map of the <see cref="Unsupported"/> scenario: the root element with the given attributes and raw content.
        /// </summary>
        internal static string UnsupportedXml(string kind, string content, string attributes = "") => $"<root xmlns=\"{UnsupportedNamespace(kind)}\"{attributes}>{content}</root>";

        /// <summary>
        /// Maps with schemas Excel inferred from XML data without schema: nested data without namespace (Excel creates a table
        /// with denormalized columns) and flat data with a namespace. The maps are named by Excel.
        /// </summary>
        internal static readonly XmlMapScenario Inferred = XmlMapScenario.FromExcel(nameof(Inferred), b =>
        {
            b.ImportInferred("Orders", "A1", InferredOrdersXml(2, 2));
            b.ImportInferred("Contacts", "B2", InferredContactsXml(2));
        });

        internal const string InferredOrdersMap = "shop_Map";
        internal const string InferredContactsMap = "contacts_Map";

        internal static string InferredOrdersXml(int orders, int lines)
        {
            var xml = new StringBuilder("<shop><owner>me</owner>");
            for (var o = 1; o <= orders; o++)
            {
                xml.Append($"<order id=\"{o}\"><date>2024-0{o % 9 + 1}-01</date>");
                for (var l = 1; l <= lines; l++)
                    xml.Append($"<line no=\"{l}\"><product>p{o}{l}</product><qty>{o * l}</qty><price>{l}.25</price></line>");
                xml.Append("</order>");
            }

            return xml.Append("</shop>").ToString();
        }

        internal static string InferredContactsXml(int count)
        {
            var xml = new StringBuilder("<contacts xmlns=\"urn:contacts\">");
            for (var i = 1; i <= count; i++)
                xml.Append($"<contact><name>n{i}</name><phone>0{i}123</phone><active>{(i % 2 == 0 ? "true" : "false")}</active></contact>");
            return xml.Append("</contacts>").ToString();
        }

        /// <summary>
        /// Formulas in mapped cells and mapped table columns, which the import replaces, and formulas that depend on mapped data.
        /// </summary>
        internal static readonly XmlMapScenario Formulas = XmlMapScenario.FromExcel(nameof(Formulas), b =>
        {
            var map = b.AddMap(DataSchema, "data");
            var cells = new (string Type, string Formula)[] { ("int", "=1+1"), ("string", "=\"x\""), ("date", "=DATE(2020,1,1)"), ("boolean", "=TRUE"), ("double", "=B1*2") };
            for (var i = 0; i < cells.Length; i++)
            {
                b.Set("Formulas", $"A{i + 1}", cells[i].Type);
                b.MapCell("Formulas", $"B{i + 1}", map, $"/ns1:data/ns1:v_{cells[i].Type}", DataSelection);
                b.Formula("Formulas", $"B{i + 1}", cells[i].Formula);
            }

            var rows = b.MapTable("Formulas", "D1:E3", map, DataSelection, "/ns1:data/ns1:rows/ns1:row/ns1:a", "/ns1:data/ns1:rows/ns1:row/ns1:n");
            rows.Name = "Rows";
            rows.ListColumns.Item(1).Name = "a";
            rows.ListColumns.Item(2).Name = "n";
            rows.ListColumns.Item(2).DataBodyRange.Formula = "=1";
            b.AddCalculatedColumn(rows, null, "Twice", "=[@n]*2");
            b.Set("Formulas", "H1", "dependent");
            b.Formula("Formulas", "H2", "=B1*10");
            b.Formula("Formulas", "H3", "=SUM(Rows[n])");
            b.Formula("Formulas", "H4", "=COUNTA(Rows[a])");
            b.Formula("Formulas", "H5", "=B3+1");
            b.Formula("Formulas", "H6", "=IF(B4,\"yes\",\"no\")");
            b.Formula("Formulas", "H7", "=LEN(B2)");
        });

        /// <summary>
        /// A mapped cell and a table on a protected sheet.
        /// </summary>
        internal static readonly XmlMapScenario Protected = XmlMapScenario.FromExcel(nameof(Protected), b =>
        {
            var map = b.AddMap(DataSchema, "data");
            b.MapCell("Protected", "B1", map, "/ns1:data/ns1:v_string", DataSelection);
            b.MapTable("Protected", "D1:E2", map, DataSelection, "/ns1:data/ns1:rows/ns1:row/ns1:a", "/ns1:data/ns1:rows/ns1:row/ns1:n");
            b.Sheet("Protected").Protect();
        });

        internal static IEnumerable<XmlMapScenario> All => new[]
        {
            RenderingTest, Types, Types1904, TypesGeneral, Tables, TwoMaps, Formats, FormatsNotPreserved, Advanced, Nested,
            SharedSchema, NoNamespace, Sheets, Stacked, Shifted, Blocked, BlockedArray, BlockedEnd, Protected, Filtered, FilteredNotPreserved,
            Unsupported, Inferred, Formulas,
        };

        internal static string NoNamespaceSchema(string root) => $@"<?xml version=""1.0"" encoding=""UTF-8""?>
<xs:schema xmlns:xs=""http://www.w3.org/2001/XMLSchema"">
  <xs:element name=""{root}"">
    <xs:complexType>
      <xs:sequence>
        <xs:element name=""name"" type=""xs:string"" minOccurs=""0""/>
        <xs:element name=""items"" minOccurs=""0"">
          <xs:complexType>
            <xs:sequence>
              <xs:element name=""item"" minOccurs=""0"" maxOccurs=""unbounded"">
                <xs:complexType>
                  <xs:sequence>
                    <xs:element name=""value"" type=""xs:double"" minOccurs=""0""/>
                  </xs:sequence>
                  <xs:attribute name=""id"" type=""xs:string""/>
                </xs:complexType>
              </xs:element>
            </xs:sequence>
          </xs:complexType>
        </xs:element>
      </xs:sequence>
    </xs:complexType>
  </xs:element>
</xs:schema>";

        internal static string NoNamespaceXml(string root, string? name, int items, string attributes = "")
        {
            var xml = new StringBuilder($"<{root}{attributes}>");
            if (name is not null)
                xml.Append($"<name>{SecurityElement.Escape(name)}</name>");
            xml.Append("<items>");
            for (var i = 1; i <= items; i++)
                xml.Append($"<item id=\"{root}{i}\"><value>{i}.5</value></item>");
            xml.Append($"</items></{root}>");
            return xml.ToString();
        }

        internal const string NestedNamespace = "urn:nested";
        internal const string NestedSelection = "xmlns:ns1='urn:nested'";
        internal const string NestedMap = "orders_Map";

        internal const string NestedSchema = @"<?xml version=""1.0"" encoding=""UTF-8""?>
<xs:schema xmlns:xs=""http://www.w3.org/2001/XMLSchema"" targetNamespace=""urn:nested"" xmlns=""urn:nested"" elementFormDefault=""qualified"">
  <xs:element name=""orders"">
    <xs:complexType>
      <xs:sequence>
        <xs:element name=""customer"" type=""xs:string"" minOccurs=""0""/>
        <xs:element name=""order"" minOccurs=""0"" maxOccurs=""unbounded"">
          <xs:complexType>
            <xs:sequence>
              <xs:element name=""date"" type=""xs:date"" minOccurs=""0""/>
              <xs:element name=""note"" minOccurs=""0"" maxOccurs=""unbounded"">
                <xs:complexType>
                  <xs:sequence>
                    <xs:element name=""text"" type=""xs:string""/>
                  </xs:sequence>
                </xs:complexType>
              </xs:element>
              <xs:element name=""line"" minOccurs=""0"" maxOccurs=""unbounded"">
                <xs:complexType>
                  <xs:sequence>
                    <xs:element name=""product"" type=""xs:string""/>
                    <xs:element name=""qty"" type=""xs:int""/>
                  </xs:sequence>
                </xs:complexType>
              </xs:element>
            </xs:sequence>
            <xs:attribute name=""id"" type=""xs:int""/>
          </xs:complexType>
        </xs:element>
      </xs:sequence>
    </xs:complexType>
  </xs:element>
</xs:schema>";

        /// <summary>
        /// XML data for the <see cref="Nested"/> scenario: the root element with the given raw content.
        /// </summary>
        internal static string NestedXml(string content) => $"<orders xmlns=\"{NestedNamespace}\">{content}</orders>";

        internal const string AdvancedNamespace = "urn:advanced";
        internal const string AdvancedSelection = "xmlns:ns1='urn:advanced'";
        internal const string AdvancedMap = "order_Map";

        private static string AdvancedXPath(string single) => single == "@currency" ? "ns1:amount/@currency" : "ns1:" + single;

        internal const string AdvancedSchema = @"<?xml version=""1.0"" encoding=""UTF-8""?>
<xs:schema xmlns:xs=""http://www.w3.org/2001/XMLSchema"" targetNamespace=""urn:advanced"" xmlns=""urn:advanced"" elementFormDefault=""qualified"">
  <xs:simpleType name=""Color"">
    <xs:restriction base=""xs:string"">
      <xs:enumeration value=""red""/>
      <xs:enumeration value=""green""/>
      <xs:enumeration value=""blue""/>
    </xs:restriction>
  </xs:simpleType>
  <xs:simpleType name=""Level"">
    <xs:restriction base=""xs:int"">
      <xs:enumeration value=""1""/>
      <xs:enumeration value=""2""/>
      <xs:enumeration value=""3""/>
    </xs:restriction>
  </xs:simpleType>
  <xs:simpleType name=""Status"">
    <xs:restriction base=""xs:token"">
      <xs:enumeration value=""open""/>
      <xs:enumeration value=""done""/>
    </xs:restriction>
  </xs:simpleType>
  <xs:simpleType name=""Percent"">
    <xs:restriction base=""xs:decimal"">
      <xs:minInclusive value=""0""/>
      <xs:maxInclusive value=""100""/>
      <xs:fractionDigits value=""2""/>
    </xs:restriction>
  </xs:simpleType>
  <xs:simpleType name=""Code"">
    <xs:restriction base=""xs:string"">
      <xs:pattern value=""[A-Z]{3}-\d{2}""/>
    </xs:restriction>
  </xs:simpleType>
  <xs:simpleType name=""DueDate"">
    <xs:restriction base=""xs:date"">
      <xs:minInclusive value=""2000-01-01""/>
    </xs:restriction>
  </xs:simpleType>
  <xs:simpleType name=""Numbers"">
    <xs:list itemType=""xs:int""/>
  </xs:simpleType>
  <xs:simpleType name=""IntOrColor"">
    <xs:union memberTypes=""xs:int Color""/>
  </xs:simpleType>
  <xs:complexType name=""Amount"">
    <xs:simpleContent>
      <xs:extension base=""xs:decimal"">
        <xs:attribute name=""currency"" type=""xs:string"" default=""EUR""/>
      </xs:extension>
    </xs:simpleContent>
  </xs:complexType>
  <xs:element name=""title"" type=""xs:string""/>
  <xs:element name=""order"">
    <xs:complexType>
      <xs:sequence>
        <xs:element ref=""title"" minOccurs=""0""/>
        <xs:element name=""color"" type=""Color"" minOccurs=""0""/>
        <xs:element name=""level"" type=""Level"" minOccurs=""0""/>
        <xs:element name=""percent"" type=""Percent"" minOccurs=""0""/>
        <xs:element name=""code"" type=""Code"" minOccurs=""0""/>
        <xs:element name=""due"" type=""DueDate"" minOccurs=""0""/>
        <xs:element name=""numbers"" type=""Numbers"" minOccurs=""0""/>
        <xs:element name=""intOrText"" type=""IntOrColor"" minOccurs=""0""/>
        <xs:element name=""amount"" type=""Amount"" minOccurs=""0""/>
        <xs:choice minOccurs=""0"">
          <xs:element name=""email"" type=""xs:string""/>
          <xs:element name=""phone"" type=""xs:string""/>
        </xs:choice>
        <xs:element name=""lines"" minOccurs=""0"">
          <xs:complexType>
            <xs:sequence>
              <xs:element name=""line"" minOccurs=""0"" maxOccurs=""unbounded"">
                <xs:complexType>
                  <xs:all>
                    <xs:element name=""product"" type=""xs:string""/>
                    <xs:element name=""qty"" type=""xs:int""/>
                    <xs:element name=""price"" type=""xs:decimal"" minOccurs=""0""/>
                    <xs:element name=""color"" type=""Color"" minOccurs=""0""/>
                  </xs:all>
                  <xs:attribute name=""status"" type=""Status"" default=""open""/>
                </xs:complexType>
              </xs:element>
            </xs:sequence>
          </xs:complexType>
        </xs:element>
      </xs:sequence>
    </xs:complexType>
  </xs:element>
</xs:schema>";

        /// <summary>
        /// XML data for the <see cref="Advanced"/> scenario.
        /// </summary>
        /// <param name="singles">Raw content of the single elements, in schema order (key: element name). <c>amount</c> may contain attributes: the value is inserted after the element name.</param>
        /// <param name="lines">Raw content of the <c>line</c> elements, each including the <c>line</c> element itself.</param>
        internal static string AdvancedXml(IReadOnlyDictionary<string, string>? singles = null, IEnumerable<string>? lines = null)
        {
            var xml = new StringBuilder($"<order xmlns=\"{AdvancedNamespace}\">");
            foreach (var name in AdvancedSingles.Where(n => !n.StartsWith("@")))
            {
                if (singles is null || !singles.TryGetValue(name, out var value))
                    continue;

                xml.Append(name == "amount" ? $"<amount{value}</amount>" : $"<{name}>{value}</{name}>");
            }

            if (lines is not null)
                xml.Append("<lines>").Append(string.Concat(lines)).Append("</lines>");

            xml.Append("</order>");
            return xml.ToString();
        }

        /// <summary>
        /// Valid lines with the elements of <c>xs:all</c> in varying order.
        /// </summary>
        internal static IEnumerable<string> AdvancedLines(int count, int start = 1)
        {
            return Enumerable.Range(start, count).Select(i => (i % 3) switch
            {
                0 => $"<line status=\"done\"><product>p{i}</product><qty>{i}</qty><price>{i}.25</price><color>red</color></line>",
                1 => $"<line><price>{i}.5</price><qty>{i}</qty><product>p{i}</product></line>",
                _ => $"<line status=\"open\"><color>blue</color><qty>{i}</qty><product>p{i}</product></line>",
            });
        }

        internal static string DataSchema
        {
            get
            {
                var singles = string.Concat(TypeNames.Select(t => $"<xs:element name=\"v_{t}\" type=\"xs:{t}\" minOccurs=\"0\"/>"));
                return $@"<?xml version=""1.0"" encoding=""UTF-8""?>
<xs:schema xmlns:xs=""http://www.w3.org/2001/XMLSchema"" targetNamespace=""{DataNamespace}"" xmlns=""{DataNamespace}"" elementFormDefault=""qualified"">
  <xs:element name=""data"">
    <xs:complexType>
      <xs:sequence>
        {singles}
        <xs:element name=""nil"" type=""xs:string"" nillable=""true"" minOccurs=""0""/>
        <xs:element name=""rows"" minOccurs=""0"">
          <xs:complexType>
            <xs:sequence>
              <xs:element name=""row"" minOccurs=""0"" maxOccurs=""unbounded"">
                <xs:complexType>
                  <xs:sequence>
                    <xs:element name=""a"" type=""xs:string"" minOccurs=""0""/>
                    <xs:element name=""d"" type=""xs:date"" minOccurs=""0""/>
                    <xs:element name=""n"" type=""xs:double"" minOccurs=""0""/>
                    <xs:element name=""b"" type=""xs:boolean"" minOccurs=""0""/>
                  </xs:sequence>
                  <xs:attribute name=""id"" type=""xs:int""/>
                </xs:complexType>
              </xs:element>
            </xs:sequence>
          </xs:complexType>
        </xs:element>
        <xs:element name=""items"" minOccurs=""0"">
          <xs:complexType>
            <xs:sequence>
              <xs:element name=""item"" minOccurs=""0"" maxOccurs=""unbounded"">
                <xs:complexType>
                  <xs:sequence>
                    <xs:element name=""name"" type=""xs:string""/>
                    <xs:element name=""qty"" type=""xs:int""/>
                  </xs:sequence>
                </xs:complexType>
              </xs:element>
            </xs:sequence>
          </xs:complexType>
        </xs:element>
      </xs:sequence>
    </xs:complexType>
  </xs:element>
</xs:schema>";
            }
        }

        internal const string OtherSchema = @"<?xml version=""1.0"" encoding=""UTF-8""?>
<xs:schema xmlns:xs=""http://www.w3.org/2001/XMLSchema"" targetNamespace=""urn:other"" xmlns=""urn:other"" elementFormDefault=""qualified"">
  <xs:element name=""other"">
    <xs:complexType>
      <xs:sequence>
        <xs:element name=""title"" type=""xs:string"" minOccurs=""0""/>
        <xs:element name=""entries"" minOccurs=""0"">
          <xs:complexType>
            <xs:sequence>
              <xs:element name=""entry"" minOccurs=""0"" maxOccurs=""unbounded"">
                <xs:complexType>
                  <xs:sequence>
                    <xs:element name=""key"" type=""xs:string""/>
                    <xs:element name=""value"" type=""xs:decimal""/>
                  </xs:sequence>
                </xs:complexType>
              </xs:element>
            </xs:sequence>
          </xs:complexType>
        </xs:element>
      </xs:sequence>
    </xs:complexType>
  </xs:element>
</xs:schema>";

        /// <summary>
        /// XML data for the data map.
        /// </summary>
        /// <param name="singles">Values of single elements, key is the type name of <see cref="TypeNames"/>. The value is inserted as is (not escaped).</param>
        /// <param name="rows">Rows of the <c>rows</c> table: id, a, d, n, b. <c>null</c> omits the element or attribute.</param>
        /// <param name="items">Rows of the <c>items</c> table: name, qty.</param>
        /// <param name="nil">Add a nil element.</param>
        /// <param name="extra">Additional raw XML appended to the root element.</param>
        internal static string DataXml(
            IReadOnlyDictionary<string, string>? singles = null,
            IEnumerable<(string? Id, string? A, string? D, string? N, string? B)>? rows = null,
            IEnumerable<(string Name, string Qty)>? items = null,
            bool nil = false,
            string extra = "")
        {
            var xml = new StringBuilder($"<data xmlns=\"{DataNamespace}\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">");
            foreach (var type in TypeNames)
            {
                if (singles is not null && singles.TryGetValue(type, out var value))
                    xml.Append($"<v_{type}>{value}</v_{type}>");
            }

            if (nil)
                xml.Append("<nil xsi:nil=\"true\"/>");

            if (rows is not null)
            {
                xml.Append("<rows>");
                foreach (var row in rows)
                {
                    xml.Append(row.Id is null ? "<row>" : $"<row id=\"{row.Id}\">");
                    AppendElement(xml, "a", row.A);
                    AppendElement(xml, "d", row.D);
                    AppendElement(xml, "n", row.N);
                    AppendElement(xml, "b", row.B);
                    xml.Append("</row>");
                }

                xml.Append("</rows>");
            }

            if (items is not null)
            {
                xml.Append("<items>");
                foreach (var (name, qty) in items)
                    xml.Append($"<item><name>{SecurityElement.Escape(name)}</name><qty>{qty}</qty></item>");
                xml.Append("</items>");
            }

            xml.Append(extra);
            xml.Append("</data>");
            return xml.ToString();
        }

        internal static IEnumerable<(string? Id, string? A, string? D, string? N, string? B)> Rows(int count, int start = 1)
        {
            return Enumerable.Range(start, count).Select(i => ((string?)i.ToString(), (string?)$"text {i}", (string?)$"2020-01-{i % 28 + 1:00}", (string?)$"{i}.5", (string?)(i % 2 == 0 ? "true" : "false")));
        }

        internal static IEnumerable<(string Name, string Qty)> Items(int count)
        {
            return Enumerable.Range(1, count).Select(i => ($"item {i}", (i * 10).ToString()));
        }

        internal static string OtherXml(string? title, params (string Key, string Value)[] entries)
        {
            var xml = new StringBuilder($"<other xmlns=\"{OtherNamespace}\">");
            if (title is not null)
                xml.Append($"<title>{SecurityElement.Escape(title)}</title>");

            xml.Append("<entries>");
            foreach (var (key, value) in entries)
                xml.Append($"<entry><key>{SecurityElement.Escape(key)}</key><value>{value}</value></entry>");
            xml.Append("</entries></other>");
            return xml.ToString();
        }

        private static void BuildFormats(ExcelWorkbookBuilder b, bool preserveFormat)
        {
            var map = b.AddMap(DataSchema, "data");
            map.PreserveNumberFormatting = preserveFormat;
            var cells = new (string Type, string Format)[]
            {
                ("int", "0.00"), ("date", "yyyy-mm-dd"), ("string", "General"), ("double", "0%"),
                ("dateTime", "yyyy-mm-dd hh:mm:ss"), ("time", "[h]:mm"), ("decimal", "#,##0.000"), ("boolean", "@"),
            };
            for (var i = 0; i < cells.Length; i++)
            {
                b.Set("Formats", $"A{i + 1}", cells[i].Type);
                b.MapCell("Formats", $"B{i + 1}", map, $"/ns1:data/ns1:v_{cells[i].Type}", DataSelection);
                b.Format("Formats", $"B{i + 1}", cells[i].Format);
            }

            b.MapTable("Formats", "D1:E2", map, DataSelection,
                "/ns1:data/ns1:rows/ns1:row/ns1:d",
                "/ns1:data/ns1:rows/ns1:row/ns1:n");
            b.Format("Formats", "D2", "dd.mm.yyyy");
            b.Format("Formats", "E2", "0.000");
        }

        private static void AppendElement(StringBuilder xml, string name, string? value)
        {
            if (value is not null)
                xml.Append($"<{name}>{SecurityElement.Escape(value)}</{name}>");
        }
    }
}
