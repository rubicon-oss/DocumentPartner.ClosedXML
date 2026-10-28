using System.Collections.Generic;
using System.Linq;
using ClosedXML.XmlMaps.Tests.Infrastructure;
using static ClosedXML.XmlMaps.Tests.Scenarios.Scenarios;

namespace ClosedXML.XmlMaps.Tests.Scenarios
{
    /// <summary>
    /// Imports that are done by Excel and by ClosedXML and must give the same workbook.
    /// </summary>
    public static class ImportCases
    {
        private static readonly Dictionary<string, string> ValidValues = new()
        {
            ["string"] = "text", ["normalizedString"] = "norm", ["token"] = "tok", ["int"] = "2", ["integer"] = "3",
            ["positiveInteger"] = "4", ["negativeInteger"] = "-5", ["nonPositiveInteger"] = "-6", ["nonNegativeInteger"] = "7",
            ["long"] = "2147483648", ["unsignedLong"] = "4294967296", ["short"] = "15", ["unsignedShort"] = "20",
            ["byte"] = "127", ["unsignedByte"] = "255", ["unsignedInt"] = "10", ["decimal"] = "25.00", ["float"] = "30.30",
            ["double"] = "35.35", ["boolean"] = "true", ["date"] = "2008-04-01", ["dateTime"] = "2008-04-01T00:01:02",
            ["time"] = "14:20:00", ["duration"] = "P1Y2M3DT10H30M0S", ["gYear"] = "2020", ["gYearMonth"] = "2020-05",
            ["gMonth"] = "--05", ["gMonthDay"] = "--05-01", ["gDay"] = "---05", ["base64Binary"] = "QmluYXJ5",
            ["hexBinary"] = "0FB7", ["anyURI"] = "http://example.com/a?b=c", ["language"] = "de-AT", ["NMTOKEN"] = "tok-1",
            ["Name"] = "name1", ["NCName"] = "ncname",
        };

        public static IEnumerable<ImportCase> All => RenderingTestCases
            .Concat(TypeCases)
            .Concat(TableCases)
            .Concat(TwoMapCases)
            .Concat(FormatCases)
            .Concat(InvalidXmlCases)
            .Concat(AdvancedCases)
            .Concat(NestedCases)
            .Concat(MultipleMapCases)
            .Concat(FilterCases)
            .Concat(UnsupportedCases)
            .Concat(InferredCases)
            .Concat(FormulaCases);

        public static IEnumerable<ImportCase> RenderingTestCases
        {
            get
            {
                var content = new XmlImport("root_Map", TestFiles.ReadText("ContentForRenderingTest.xml"));
                yield return new ImportCase("RenderingTest", RenderingTest, content);
                yield return new ImportCase("RenderingTest twice", RenderingTest, content, content);
                yield return new ImportCase("RenderingTest append", RenderingTest, content, content with { Overwrite = false });
            }
        }

        public static IEnumerable<ImportCase> TypeCases
        {
            get
            {
                yield return Types("Types valid", ValidValues);
                yield return Types("Types numbers A", new()
                {
                    ["int"] = "007", ["integer"] = "+12", ["decimal"] = "1.50", ["float"] = "1E3", ["double"] = ".5",
                    ["long"] = "12345678901234567", ["unsignedLong"] = "18446744073709551615", ["short"] = "-0",
                    ["byte"] = "  7  ", ["unsignedByte"] = "300", ["positiveInteger"] = "1,5", ["negativeInteger"] = "-1e2",
                    ["nonNegativeInteger"] = "0x10", ["unsignedShort"] = "1.", ["unsignedInt"] = "5.0E-3",
                });
                yield return Types("Types numbers B", new()
                {
                    ["double"] = "INF", ["float"] = "-INF", ["decimal"] = "NaN", ["int"] = "2147483648",
                    ["integer"] = "99999999999999999999999", ["long"] = "-9223372036854775808", ["short"] = "1 000",
                    ["unsignedByte"] = "+0", ["positiveInteger"] = "00001", ["unsignedShort"] = "1.5E+300",
                    ["byte"] = "1e400", ["nonPositiveInteger"] = "-.5", ["unsignedInt"] = " 42\n",
                });
                yield return Types("Types dates", new()
                {
                    ["date"] = "2020-02-29", ["dateTime"] = "2020-02-29T13:14:15.5", ["time"] = "23:59:59.999",
                    ["gYear"] = "1999", ["gYearMonth"] = "1999-12", ["duration"] = "PT1H",
                });
                yield return Types("Types dates time zones", new()
                {
                    ["date"] = "2020-02-29Z", ["dateTime"] = "2020-02-29T13:14:15+02:00", ["time"] = "10:00:00Z",
                });
                yield return Types("Types dates invalid", new()
                {
                    ["date"] = "2020-13-01", ["dateTime"] = "2020-02-30T00:00:00", ["time"] = "25:00:00", ["gYear"] = "20x0",
                });
                yield return Types("Types dates limits", new()
                {
                    ["date"] = "1899-12-31", ["dateTime"] = "1900-01-01T00:00:00",
                });
                yield return Types("Types gYear negative", new() { ["gYear"] = "-0001" }) with
                {
                    KnownDeviation = "Excel accepts negative years as valid xs:gYear, the schema validation of .NET rejects them.",
                };
                yield return Types("Types time 24:00", new() { ["time"] = "24:00:00" }) with
                {
                    KnownDeviation = "Excel accepts 24:00:00 as valid xs:time, the schema validation of .NET rejects it. " +
                                     "The value and the date format are the same as in Excel.",
                };
                yield return Types("Types dates early 1900", new()
                {
                    ["date"] = "1900-01-01", ["dateTime"] = "1900-02-28T23:59:59", ["time"] = "00:00:00",
                });
                yield return Types("Types dates far", new()
                {
                    ["date"] = "9999-12-31", ["dateTime"] = "2008-04-01T00:00:00.1234567", ["time"] = "12:00:00.5",
                });
                foreach (var boolean in new[] { "false", "1", "0", "TRUE", "yes" })
                    yield return Types($"Types boolean {boolean}", new() { ["boolean"] = boolean });

                yield return Types("Types strings", new()
                {
                    ["string"] = "  padded  ", ["normalizedString"] = "a&#9;b", ["token"] = "a  b", ["NCName"] = "'quoted",
                    ["Name"] = "=1+1", ["language"] = "123", ["anyURI"] = "  http://x  ", ["base64Binary"] = "QmluYXJ5&#10;RGF0YQ==",
                    ["NMTOKEN"] = "&#160;nbsp&#160;", ["duration"] = "  P1D ", ["gMonth"] = "line1&#10;line2",
                });
                yield return Types("Types long string", new() { ["string"] = new string('x', 40000) });
                yield return Types("Types empty elements", new() { ["string"] = string.Empty, ["int"] = string.Empty, ["date"] = string.Empty });
                yield return new ImportCase("Types missing elements", Scenarios.Types,
                    new XmlImport(DataMap, DataXml(ValidValues)),
                    new XmlImport(DataMap, DataXml(new Dictionary<string, string> { ["string"] = "only", ["int"] = "1" })));
                yield return new ImportCase("Types nil", Scenarios.Types,
                    new XmlImport(DataMap, DataXml(ValidValues)),
                    new XmlImport(DataMap, DataXml(new Dictionary<string, string> { ["string"] = "x" }, nil: true)));
                yield return new ImportCase("Types append", Scenarios.Types,
                    new XmlImport(DataMap, DataXml(ValidValues)),
                    new XmlImport(DataMap, DataXml(new Dictionary<string, string> { ["string"] = "appended" }), Overwrite: false));
                yield return new ImportCase("Types not valid", Scenarios.Types,
                    new XmlImport(DataMap, DataXml(ValidValues, extra: "<unknown>1</unknown>")));
                yield return new ImportCase("Types 1904", Types1904,
                    new XmlImport(DataMap, DataXml(ValidValues)));
                yield return new ImportCase("Types 1904 limits", Types1904, new XmlImport(DataMap, DataXml(new Dictionary<string, string>
                {
                    ["date"] = "1903-12-31", ["dateTime"] = "1904-01-01T00:00:00", ["time"] = "00:00:00.001", ["gYearMonth"] = "1904-01",
                    ["gMonthDay"] = "--02-29",
                })));
                yield return Types("Types g spellings", new()
                {
                    ["gMonth"] = "--05--", ["gDay"] = "---32", ["gMonthDay"] = "--02-29", ["gYearMonth"] = "1999-12Z", ["gYear"] = "0099",
                });
                yield return Types("Types g spellings invalid", new()
                {
                    ["gMonth"] = "--13", ["gDay"] = "---05Z", ["gMonthDay"] = "--02-30", ["gYearMonth"] = "1999-1", ["gYear"] = "99",
                });
                yield return Types("Types dates time zones shifting the date", new()
                {
                    ["date"] = "2020-01-01-05:00", ["dateTime"] = "2020-01-01T23:00:00-05:00", ["time"] = "23:00:00+14:00",
                });
                yield return Types("Types dates loose spellings", new()
                {
                    ["date"] = "2020-1-1", ["dateTime"] = "2020-01-01 10:00:00", ["time"] = "10:00", ["gYear"] = "+2020",
                });
                yield return Types("Types dates loose spellings B", new()
                {
                    ["date"] = "01.01.2020", ["dateTime"] = "2020-01-01T10:00", ["time"] = "10:00:00.", ["gYearMonth"] = "12-1999",
                });
                yield return Types("Types numbers C", new()
                {
                    ["double"] = "1.5E+308", ["float"] = "3.4028235E+38", ["decimal"] = "0.1234567890123456789",
                    ["int"] = "-0", ["integer"] = "１２", ["long"] = "1E-400", ["short"] = "+1.5", ["unsignedLong"] = "0042",
                    ["byte"] = "4.9E-324", ["unsignedInt"] = "1234.5000",
                });
                yield return Types("Types numbers out of range", new()
                {
                    ["double"] = "0.0E-999", ["float"] = "0e5", ["decimal"] = "1e400", ["int"] = "-1E-400",
                });
                yield return Types("Types g spellings B", new()
                {
                    ["gYear"] = "+999", ["gMonth"] = "--05", ["gDay"] = "---5", ["gMonthDay"] = "--02-31", ["gYearMonth"] = "2020-01",
                });
                yield return Types("Types g spellings C", new()
                {
                    ["gYear"] = "-999", ["gMonth"] = "--12--", ["gDay"] = "---31", ["gMonthDay"] = "--04-31",
                });
                yield return new ImportCase("Types 1904 first day of year", Types1904, new XmlImport(DataMap, DataXml(new Dictionary<string, string>
                {
                    ["gMonthDay"] = "--01-01", ["date"] = "1904-01-01", ["dateTime"] = "2020-01-01T24:00", ["time"] = "23:59",
                })));
                foreach (var boolean in new[] { "True", "-1", "2", "1.5", "1e1", "0x10" })
                    yield return Types($"Types boolean {boolean}", new() { ["boolean"] = boolean });
                yield return new ImportCase("Types general format", TypesGeneral,
                    new XmlImport(DataMap, DataXml(ValidValues)));
            }
        }

        public static IEnumerable<ImportCase> TableCases
        {
            get
            {
                yield return Table("Tables 1 row", DataXml(rows: Rows(1)));
                yield return Table("Tables grow 5 rows", DataXml(rows: Rows(5)));
                yield return Table("Tables grow 20 rows", DataXml(rows: Rows(20)));
                yield return Table("Tables shrink", DataXml(rows: Rows(5)), DataXml(rows: Rows(2)));
                yield return Table("Tables no rows", DataXml(rows: Rows(5)), DataXml(rows: Rows(0)));
                yield return Table("Tables no rows element", DataXml(rows: Rows(5)), DataXml());
                yield return new ImportCase("Tables append", Scenarios.Tables,
                    new XmlImport(DataMap, DataXml(rows: Rows(2))),
                    new XmlImport(DataMap, DataXml(rows: Rows(3, 10)), Overwrite: false));
                yield return new ImportCase("Tables append to empty", Scenarios.Tables,
                    new XmlImport(DataMap, DataXml(rows: Rows(2)), Overwrite: false));
                yield return new ImportCase("Tables append after clear", Scenarios.Tables,
                    new XmlImport(DataMap, DataXml(rows: Rows(3))),
                    new XmlImport(DataMap, DataXml(rows: Rows(0))),
                    new XmlImport(DataMap, DataXml(rows: Rows(2)), Overwrite: false));
                yield return Table("Tables missing values", DataXml(rows: new (string?, string?, string?, string?, string?)[]
                {
                    (null, "x", null, "1", null),
                    ("2", null, "2020-01-01", null, "true"),
                    (null, null, null, null, null),
                }));
                yield return Table("Tables invalid values", DataXml(rows: new (string?, string?, string?, string?, string?)[]
                {
                    ("x", "  a  ", "not a date", "abc", "maybe"),
                    ("1.5", "", "2020-01-01Z", "1E3", "1"),
                }));
                yield return Table("Tables totals row grow", DataXml(items: Items(4)));
                yield return Table("Tables totals row shrink", DataXml(items: Items(4)), DataXml(items: Items(1)));
                yield return Table("Tables totals row no rows", DataXml(items: Items(3)), DataXml(items: Items(0)));
                yield return Table("Tables both", DataXml(rows: Rows(3), items: Items(2)));
            }
        }

        public static IEnumerable<ImportCase> TwoMapCases
        {
            get
            {
                var data = new XmlImport(DataMap, DataXml(new Dictionary<string, string> { ["string"] = "first", ["int"] = "1" }, Rows(3)));
                var other = new XmlImport(OtherMap, OtherXml("title", ("a", "1.5"), ("b", "2"), ("c", "-3")));
                yield return new ImportCase("TwoMaps first", TwoMaps, data);
                yield return new ImportCase("TwoMaps both", TwoMaps, data, other);
                yield return new ImportCase("TwoMaps second shrink", TwoMaps, other, new XmlImport(OtherMap, OtherXml(null, ("z", "9"))));
            }
        }

        public static IEnumerable<ImportCase> FormatCases
        {
            get
            {
                var xml = DataXml(ValidValues, Rows(2));
                yield return new ImportCase("Formats", Formats, new XmlImport(DataMap, xml));
                yield return new ImportCase("Formats not preserved", FormatsNotPreserved, new XmlImport(DataMap, xml));
            }
        }

        /// <summary>
        /// XML that doesn't match the schema of the map, or isn't even well-formed. Valid data is imported first,
        /// so the cases also show which cells are cleared or kept.
        /// </summary>
        public static IEnumerable<ImportCase> InvalidXmlCases
        {
            get
            {
                const string valid = "<v_string>a</v_string><v_int>1</v_int>";
                yield return InvalidTypes("Invalid unknown attribute on root", Data(valid, " foo=\"1\""));
                yield return InvalidTypes("Invalid unknown attribute on element", Data("<v_string lang=\"x\">a</v_string><v_int>1</v_int>"));
                yield return InvalidTypes("Invalid attribute in other namespace", Data(valid, " xmlns:x=\"urn:x\" x:foo=\"1\""));
                yield return InvalidTypes("Invalid xsi schemaLocation", Data(valid, $" xsi:schemaLocation=\"{DataNamespace} missing.xsd\""));
                yield return InvalidTypes("Invalid wrong order", Data("<v_int>1</v_int><v_string>a</v_string>"));
                yield return InvalidTypes("Invalid duplicate element", Data("<v_string>first</v_string><v_string>second</v_string><v_int>1</v_int>"));
                yield return InvalidTypes("Invalid unknown element between", Data("<v_string>a</v_string><unknown>u</unknown><v_int>1</v_int>"));
                yield return InvalidTypes("Invalid unknown element at start", Data("<unknown>u</unknown>" + valid));
                yield return InvalidTypes("Invalid child element in simple element", Data("<v_string>a<b>bold</b>c</v_string><v_int>1<x/>2</v_int>"));
                yield return InvalidTypes("Invalid element in other namespace", Data("<v_string xmlns=\"urn:x\">a</v_string><v_int>1</v_int>"));
                yield return InvalidTypes("Invalid element without namespace", Data("<v_string xmlns=\"\">a</v_string><v_int>1</v_int>"));
                yield return InvalidTypes("Invalid nil on element not nillable", Data("<v_string xsi:nil=\"true\"/><v_int>1</v_int>"));
                yield return InvalidTypes("Invalid root in other namespace", "<data xmlns=\"urn:x\"><v_string>a</v_string></data>");
                yield return InvalidTypes("Invalid root without namespace", "<data><v_string>a</v_string></data>");
                yield return InvalidTypes("Invalid wrong root element", $"<root xmlns=\"{DataNamespace}\"><v_string>a</v_string></root>");
                yield return InvalidTypes("Invalid XML of other map", OtherXml("title", ("a", "1")));
                yield return InvalidTypes("Invalid not well-formed", $"<data xmlns=\"{DataNamespace}\"><v_string>a</data>");
                yield return InvalidTypes("Invalid undeclared prefix", $"<data xmlns=\"{DataNamespace}\"><p:v_string>a</p:v_string></data>");
                yield return InvalidTypes("Valid empty root", Data(string.Empty));
                yield return InvalidTypes("Valid other prefix", $"<p:data xmlns:p=\"{DataNamespace}\"><p:v_string>a</p:v_string><p:v_int>1</p:v_int></p:data>");
                yield return InvalidTypes("Valid default namespace redeclared", $"<p:data xmlns:p=\"{DataNamespace}\"><v_string xmlns=\"{DataNamespace}\">a</v_string><p:v_int>1</p:v_int></p:data>");
                yield return InvalidTypes("Valid comments CDATA processing instructions", Data("<v_string><!-- c --><![CDATA[ <x> ]]>b</v_string><?pi x?><v_int><!-- c -->1</v_int>"));
                yield return InvalidTypes("Valid CDATA whitespace", Data(
                    "<v_string>  a<![CDATA[ x ]]>b  </v_string><v_normalizedString>  a<!-- c -->  b  </v_normalizedString>" +
                    "<v_int><![CDATA[ 5 ]]></v_int><v_integer>  <![CDATA[6]]>  </v_integer>" +
                    "<v_boolean><![CDATA[ true ]]></v_boolean><v_date> <![CDATA[2020-01-01]]> </v_date><v_Name>  <![CDATA[ x ]]></v_Name>"));
                yield return InvalidTypes("Token whitespace", Data("<v_token> a  b </v_token>"));
                yield return InvalidTypes("Valid token only whitespace", Data("<v_token><![CDATA[   ]]></v_token>")) with
                {
                    KnownDeviation = "A token of only whitespace collapses to an empty token, which Excel accepts. " +
                                     "The schema validation of .NET rejects it.",
                };

                const string row = "<a>t</a><n>1</n>";
                yield return InvalidTables("Invalid table unknown attribute on row", RowsXml($"<row id=\"1\" extra=\"x\">{row}</row><row id=\"2\">{row}</row>"));
                yield return InvalidTables("Invalid table attribute in other namespace", RowsXml($"<row x:id=\"1\" xmlns:x=\"urn:x\">{row}</row><row id=\"2\">{row}</row>"));
                yield return InvalidTables("Invalid table unknown element in row", RowsXml($"<row id=\"1\"><a>t</a><z>zz</z><n>1</n></row>"));
                yield return InvalidTables("Invalid table attribute as element", RowsXml($"<row><id>5</id>{row}</row>"));
                yield return InvalidTables("Invalid table wrong order in row", RowsXml("<row id=\"1\"><n>1</n><a>t</a></row>"));
                yield return InvalidTables("Invalid table duplicate element in row", RowsXml("<row id=\"1\"><a>first</a><a>second</a><n>1</n><n>2</n></row>"));
                yield return InvalidTables("Invalid table row in other namespace", RowsXml($"<row id=\"1\">{row}</row><row id=\"2\" xmlns=\"urn:x\">{row}</row><row id=\"3\">{row}</row>"));
                yield return InvalidTables("Invalid table rows in unknown wrapper", RowsXml($"<group><row id=\"1\">{row}</row></group>"));
                yield return InvalidTables("Invalid table rows twice", Data($"<rows><row id=\"1\">{row}</row></rows><rows><row id=\"2\">{row}</row><row id=\"3\">{row}</row></rows>"));
                yield return InvalidTables("Invalid table row outside rows", Data($"<row id=\"1\">{row}</row>"));
                yield return InvalidTables("Invalid table missing required element", Data("<items><item><name>a</name></item><item><qty>2</qty></item></items>"));
                yield return InvalidTables("Invalid table text in rows", RowsXml($"text<row id=\"1\">{row}</row>more"));
            }
        }

        private static readonly Dictionary<string, string> AdvancedValues = new()
        {
            ["title"] = "Order 1", ["color"] = "green", ["level"] = "2", ["percent"] = "12.5", ["code"] = "ABC-12",
            ["due"] = "2024-06-30", ["numbers"] = "1 2 3", ["intOrText"] = "42", ["amount"] = " currency=\"USD\">99.95", ["email"] = "a@b.c",
        };

        public static IEnumerable<ImportCase> AdvancedCases
        {
            get
            {
                yield return Advanced("Advanced valid", AdvancedXml(AdvancedValues, AdvancedLines(3)));
                yield return Advanced("Advanced enum invalid", AdvancedXml(
                    new Dictionary<string, string> { ["color"] = "purple", ["level"] = "5", ["intOrText"] = "yellow" },
                    new[] { "<line status=\"weird\"><product>x</product><qty>1</qty><color>pink</color></line>" }));
                yield return Advanced("Advanced enum case and whitespace", AdvancedXml(
                    new Dictionary<string, string> { ["color"] = "Red", ["level"] = " 3 ", ["intOrText"] = "blue" },
                    new[] { "<line status=\" done \"><product>x</product><qty>1</qty><color> red </color></line>" }));
                yield return Advanced("Advanced facets invalid", AdvancedXml(
                    new Dictionary<string, string> { ["percent"] = "150", ["code"] = "abc-1", ["due"] = "1999-12-31", ["numbers"] = "1 x 3" }));
                yield return Advanced("Advanced fraction digits", AdvancedXml(new Dictionary<string, string> { ["percent"] = "1.234" }));
                yield return Advanced("Advanced list and union", AdvancedXml(new Dictionary<string, string> { ["numbers"] = "  7   8 ", ["intOrText"] = "-5" }));
                yield return Advanced("Advanced choice second", AdvancedXml(new Dictionary<string, string> { ["phone"] = "+43 1 234" }));
                yield return Advanced("Advanced choice both", AdvancedXml(new Dictionary<string, string> { ["email"] = "a@b.c", ["phone"] = "+43 1 234" }));
                yield return Advanced("Advanced defaults missing", AdvancedXml(
                    new Dictionary<string, string> { ["amount"] = ">5" },
                    new[] { "<line><product>x</product><qty>2</qty></line>", "<line status=\"\"><product>y</product><qty>3</qty></line>" }));
                yield return Advanced("Advanced all missing required", AdvancedXml(lines: new[] { "<line><price>1</price></line>", "<line><qty>1</qty><qty>2</qty><product>z</product></line>" }));
                yield return Advanced("Advanced grow", AdvancedXml(lines: AdvancedLines(1)), AdvancedXml(lines: AdvancedLines(7)));
                yield return Advanced("Advanced shrink", AdvancedXml(lines: AdvancedLines(6)), AdvancedXml(lines: AdvancedLines(2)));
                yield return Advanced("Advanced no lines", AdvancedXml(lines: AdvancedLines(4)), AdvancedXml(AdvancedValues));
                yield return new ImportCase("Advanced append", Scenarios.Advanced,
                    new XmlImport(AdvancedMap, AdvancedXml(lines: AdvancedLines(2))),
                    new XmlImport(AdvancedMap, AdvancedXml(lines: AdvancedLines(3, 10)), Overwrite: false));
            }
        }

        public static IEnumerable<ImportCase> NestedCases
        {
            get
            {
                const string order1 = "<order id=\"1\"><date>2024-01-01</date><note><text>n1</text></note><note><text>n2</text></note>" +
                                      "<line><product>a</product><qty>1</qty></line><line><product>b</product><qty>2</qty></line></order>";
                const string order2 = "<order id=\"2\"><date>2024-02-01</date></order>";
                const string order3 = "<order id=\"3\"><line><product>c</product><qty>3</qty></line><line><product>d</product><qty>4</qty></line>" +
                                      "<line><product>e</product><qty>5</qty></line></order>";
                yield return Nested("Nested orders", NestedXml("<customer>ACME</customer>" + order1 + order2 + order3));
                yield return Nested("Nested without customer", NestedXml(order1 + order3));
                yield return Nested("Nested only customer", NestedXml("<customer>ACME</customer>"));
                yield return Nested("Nested empty root", NestedXml("<customer>ACME</customer>" + order1), NestedXml(string.Empty));
                yield return Nested("Nested empty elements", NestedXml("<customer>ACME</customer><order id=\"1\"/><order/><order><line/><note/></order>" +
                                                                       "<order id=\"3\"><line/><line><product>x</product><qty>1</qty></line></order>"));
                yield return Nested("Nested customer after orders", NestedXml(order1 + "<customer>LATE</customer>"));
                yield return Nested("Nested shrink", NestedXml("<customer>ACME</customer>" + order1 + order2 + order3), NestedXml(order2));
                yield return new ImportCase("Nested append", Scenarios.Nested,
                    new XmlImport(NestedMap, NestedXml("<customer>ACME</customer>" + order1)),
                    new XmlImport(NestedMap, NestedXml("<customer>OTHER</customer>" + order3), Overwrite: false));
            }
        }

        /// <summary>
        /// Several maps, maps without namespace, tables on several sheets and tables that are in the way of each other.
        /// </summary>
        public static IEnumerable<ImportCase> MultipleMapCases
        {
            get
            {
                var first = new XmlImport(FirstMap, DataXml(new Dictionary<string, string> { ["string"] = "first" }, Rows(3)));
                var second = new XmlImport(SecondMap, DataXml(new Dictionary<string, string> { ["string"] = "second" }, Rows(2, 10)));
                yield return new ImportCase("Shared schema first", SharedSchema, first);
                yield return new ImportCase("Shared schema both", SharedSchema, first, second);
                yield return new ImportCase("Shared schema second clears only its cells", SharedSchema, first, second, new XmlImport(SecondMap, DataXml()));

                var alpha = new XmlImport("alpha_Map", NoNamespaceXml("alpha", "a", 3));
                var beta = new XmlImport("beta_Map", NoNamespaceXml("beta", "b", 2));
                yield return new ImportCase("No namespace both", NoNamespace, alpha, beta);
                yield return new ImportCase("No namespace wrong root", NoNamespace, alpha, new XmlImport("alpha_Map", NoNamespaceXml("beta", "b", 2)));
                yield return new ImportCase("No namespace data with namespace", NoNamespace, alpha, new XmlImport("alpha_Map", NoNamespaceXml("alpha", "ns", 1, " xmlns=\"urn:x\"")));

                yield return new ImportCase("Sheets", Sheets, new XmlImport(DataMap, DataXml(ValidValues, Rows(4), Items(3))));
                yield return new ImportCase("Sheets shrink", Sheets,
                    new XmlImport(DataMap, DataXml(ValidValues, Rows(4), Items(3))),
                    new XmlImport(DataMap, DataXml(rows: Rows(1))));

                var otherEntries = new XmlImport(OtherMap, OtherXml("t", ("a", "1"), ("b", "2")));
                yield return new ImportCase("Stacked grow into gap", Stacked, new XmlImport(DataMap, DataXml(rows: Rows(2))));
                yield return new ImportCase("Stacked grow into table below", Stacked, otherEntries, new XmlImport(DataMap, DataXml(rows: Rows(4))));
                yield return new ImportCase("Stacked grow far", Stacked, otherEntries, new XmlImport(DataMap, DataXml(rows: Rows(12), items: Items(2))));
                yield return new ImportCase("Stacked grow lower table", Stacked, otherEntries, new XmlImport(OtherMap, OtherXml("t", ("a", "1"), ("b", "2"), ("c", "3"), ("d", "4"), ("e", "5"), ("f", "6"))));
                yield return new ImportCase("Stacked shrink", Stacked, new XmlImport(DataMap, DataXml(rows: Rows(2), items: Items(5))), new XmlImport(DataMap, DataXml(rows: Rows(1), items: Items(1))));

                yield return new ImportCase("Blocked by merged cells", Blocked, new XmlImport(DataMap, DataXml(rows: Rows(2))), new XmlImport(DataMap, DataXml(rows: Rows(5))));
                yield return new ImportCase("Blocked by end of sheet", Blocked, new XmlImport(DataMap, DataXml(items: Items(2))), new XmlImport(DataMap, DataXml(items: Items(3))));

                yield return new ImportCase("Shifted grow", Shifted, new XmlImport(DataMap, DataXml(rows: Rows(4))));
                yield return new ImportCase("Shifted grow and shrink", Shifted, new XmlImport(DataMap, DataXml(rows: Rows(8))), new XmlImport(DataMap, DataXml(rows: Rows(1))));
                yield return new ImportCase("Blocked by array formula", BlockedArray, new XmlImport(DataMap, DataXml(rows: Rows(2))), new XmlImport(DataMap, DataXml(rows: Rows(4))));
                yield return new ImportCase("Blocked by last row", BlockedEnd, new XmlImport(DataMap, DataXml(rows: Rows(2))), new XmlImport(DataMap, DataXml(rows: Rows(4))));
                yield return new ImportCase("Protected sheet", Protected, new XmlImport(DataMap, DataXml(new Dictionary<string, string> { ["string"] = "x" }, Rows(3))));
            }
        }

        /// <summary>
        /// Tables with filters and sort orders. Excel applies the filters to the imported rows, but doesn't sort them.
        /// </summary>
        public static IEnumerable<ImportCase> FilterCases
        {
            get
            {
                var data = DataXml(rows: new[] { 5, 1, 4, 2, 7, 3 }.Select(i => ((string?)null, (string?)$"t{i}", (string?)null, (string?)i.ToString(), (string?)null)),
                                   items: new[] { ("b", "30"), ("c", "10"), ("a", "20"), ("d", "30") });
                yield return new ImportCase("Filtered", Filtered, new XmlImport(DataMap, data));
                yield return new ImportCase("Filtered shrink", Filtered, new XmlImport(DataMap, data), new XmlImport(DataMap, DataXml(rows: Rows(2), items: Items(1))));
                yield return new ImportCase("Filtered append", Filtered, new XmlImport(DataMap, data), new XmlImport(DataMap, data, Overwrite: false));
                yield return new ImportCase("Filtered not preserved", FilteredNotPreserved, new XmlImport(DataMap, data));
            }
        }

        /// <summary>
        /// Data for schemas with constructs Excel doesn't support for mapping.
        /// </summary>
        public static IEnumerable<ImportCase> UnsupportedCases
        {
            get
            {
                ImportCase Case(string name, string kind, string content, string attributes = "") =>
                    new(name, Unsupported, new XmlImport(kind + "_Map", UnsupportedXml(kind, content, attributes)));

                yield return Case("Unsupported any none", "any", "<name>n</name>");
                yield return Case("Unsupported any declared lax", "any", "<name>n</name><x:a xmlns:x=\"urn:lax\">1</x:a><x:b xmlns:x=\"urn:lax\">2</x:b>");
                yield return Case("Unsupported any undeclared strict", "any", "<name>n</name><x:a xmlns:x=\"urn:strict\">1</x:a>");
                yield return Case("Unsupported any other namespace", "any", "<name>n</name><x:a xmlns:x=\"urn:x\">1</x:a>");
                yield return Case("Unsupported any same namespace", "any", "<name>n</name><name>second</name>");
                yield return Case("Unsupported any before name", "any", "<x:a xmlns:x=\"urn:lax\">1</x:a><name>n</name>");
                yield return Case("Unsupported anyAttribute", "anyAttribute", "<name a=\"1\">n</name>", " foo=\"1\" xmlns:x=\"urn:x\" x:bar=\"2\"");
                yield return Case("Unsupported anyAttribute on root only", "anyAttribute", "<name>n</name>", " foo=\"1\" xmlns:x=\"urn:x\" x:bar=\"2\"");
                yield return Case("Unsupported recursive", "recursive",
                    "<name>n</name><node><label>a</label><node><label>a1</label></node><node><label>a2</label><node><label>a21</label></node></node></node>" +
                    "<node><label>b</label></node><node><node><label>c1</label></node></node>");
                yield return Case("Unsupported abstract substitution", "abstract", "<name>n</name><circle>c</circle>");
                yield return Case("Unsupported abstract element", "abstract", "<name>n</name><shape>s</shape>");
                yield return Case("Unsupported mixed", "mixed", "<name>n</name><text>before <b>one</b> between <b>two</b> after</text>");
                yield return Case("Unsupported mixed text only", "mixed", "<name>n</name><text>only text</text>");
            }
        }

        /// <summary>
        /// Maps with schemas inferred by Excel.
        /// </summary>
        public static IEnumerable<ImportCase> InferredCases
        {
            get
            {
                yield return new ImportCase("Inferred orders grow", Inferred, new XmlImport(InferredOrdersMap, InferredOrdersXml(3, 4)));
                yield return new ImportCase("Inferred orders shrink", Inferred, new XmlImport(InferredOrdersMap, InferredOrdersXml(1, 1)));
                yield return new ImportCase("Inferred orders other values", Inferred, new XmlImport(InferredOrdersMap,
                    "<shop><owner>x</owner><order id=\"a\"><date>soon</date><line no=\"1\"><product>p</product><qty>1.5</qty><price>x</price></line></order></shop>"));
                yield return new ImportCase("Inferred orders unknown element", Inferred, new XmlImport(InferredOrdersMap,
                    "<shop><owner>x</owner><extra>e</extra><order id=\"1\"><line no=\"1\"><product>p</product><qty>1</qty><price>2</price><note>n</note></line></order></shop>"));
                yield return new ImportCase("Inferred contacts", Inferred, new XmlImport(InferredContactsMap, InferredContactsXml(5)));
                yield return new ImportCase("Inferred contacts append", Inferred, new XmlImport(InferredContactsMap, InferredContactsXml(3), Overwrite: false));
            }
        }

        /// <summary>
        /// The import replaces formulas in mapped cells, formulas depending on mapped cells get the imported values.
        /// </summary>
        public static IEnumerable<ImportCase> FormulaCases
        {
            get
            {
                yield return new ImportCase("Formulas replaced", Formulas, new XmlImport(DataMap, DataXml(ValidValues, Rows(4))));
                yield return new ImportCase("Formulas missing elements", Formulas, new XmlImport(DataMap, DataXml(new Dictionary<string, string> { ["int"] = "5" })));
                yield return new ImportCase("Formulas append", Formulas, new XmlImport(DataMap, DataXml(rows: Rows(2)), Overwrite: false));
            }
        }

        private static ImportCase Nested(string name, params string[] xmls)
        {
            return new ImportCase(name, Scenarios.Nested, xmls.Select(x => new XmlImport(NestedMap, x)).ToArray());
        }

        private static ImportCase Advanced(string name, params string[] xmls)
        {
            return new ImportCase(name, Scenarios.Advanced, xmls.Select(x => new XmlImport(AdvancedMap, x)).ToArray());
        }

        private static ImportCase Types(string name, Dictionary<string, string> values)
        {
            return new ImportCase(name, Scenarios.Types, new XmlImport(DataMap, DataXml(values)));
        }

        private static ImportCase Table(string name, params string[] xmls)
        {
            return new ImportCase(name, Scenarios.Tables, xmls.Select(x => new XmlImport(DataMap, x)).ToArray());
        }

        private static ImportCase InvalidTypes(string name, string xml)
        {
            return new ImportCase(name, Scenarios.Types, new XmlImport(DataMap, DataXml(ValidValues)), new XmlImport(DataMap, xml));
        }

        private static ImportCase InvalidTables(string name, string xml)
        {
            return Table(name, DataXml(rows: Rows(3), items: Items(2)), xml);
        }

        /// <summary>
        /// Raw XML for the data map: the root element with the given content.
        /// </summary>
        private static string Data(string content, string attributes = "")
        {
            return $"<data xmlns=\"{DataNamespace}\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"{attributes}>{content}</data>";
        }

        private static string RowsXml(string content) => Data($"<rows>{content}</rows>");
    }
}
