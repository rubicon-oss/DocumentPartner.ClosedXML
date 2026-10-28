# XML maps

This fork of ClosedXML supports XML maps (the custom XML mappings of SpreadsheetML, ECMA-376). An XML map connects an
XML schema (XSD) with single cells and table columns of a workbook. With the map, XML data can be imported into the
cells and the content of the cells can be exported as XML.

ClosedXML can:

- load and save workbooks with XML maps without losing them, also when saving to a new file or after changing the
  workbook (inserted rows, deleted tables or sheets, …);
- import XML data into the mapped cells and tables;
- export the mapped cells and tables as XML;
- delete a map.

Maps can't be created with ClosedXML. The map has to exist in the loaded workbook, e.g. a template. The model is
prepared for map creation (see [Architecture](#architecture)).

The file format follows ECMA-376. Where the standard leaves the behavior open, this document specifies the behavior
of ClosedXML. The behavior is covered by tests, see [Tests](#tests).

> This feature isn't part of upstream ClosedXML and won't be accepted there. All code is in this folder, upstream
> files only contain a few one-line hook calls, see [Merging upstream](#merging-upstream).

## Contents

- [XML maps](#xml-maps)
  - [Contents](#contents)
  - [Quick start](#quick-start)
  - [API](#api)
    - [`IXLXmlMaps`](#ixlxmlmaps)
    - [`IXLXmlMap`](#ixlxmlmap)
    - [Results and errors](#results-and-errors)
  - [Import](#import)
    - [Validation](#validation)
    - [Single cells](#single-cells)
    - [Values](#values)
    - [Number formats](#number-formats)
    - [Tables](#tables)
    - [Growing tables](#growing-tables)
    - [Filters and sort order](#filters-and-sort-order)
    - [Column widths](#column-widths)
    - [Formulas](#formulas)
  - [Export](#export)
    - [Maps that can't be exported](#maps-that-cant-be-exported)
  - [Delete](#delete)
  - [Loading and saving](#loading-and-saving)
  - [Limitations](#limitations)
  - [Architecture](#architecture)
    - [Merging upstream](#merging-upstream)
  - [Tests](#tests)

## Quick start

XML maps are opt-in. Unless `XLWorkbook.EnableXmlMaps` is set to `true` (once, before workbooks are loaded),
XML maps are neither loaded nor saved and `XmlMaps()` throws an `InvalidOperationException`.

```csharp
using ClosedXML.Excel;

XLWorkbook.EnableXmlMaps = true;

using var workbook = new XLWorkbook("Template.xlsx");
var map = workbook.XmlMaps().Get("root_Map");

// Import: the result tells whether the data was valid, the data is imported anyway.
var result = map.Import("data.xml");
if (result == XLXmlImportResult.ValidationFailed)
    Console.WriteLine("The XML doesn't match the schema of the map.");

// Export: the XML is produced from the current content of the mapped cells.
if (map.IsExportable)
    map.Export("export.xml", overwrite: true);

// Remove the map after the import, the values stay in the cells.
map.Delete();

workbook.SaveAs("Result.xlsx");
```

The name of a map is the `Name` attribute of its `Map` element in `xmlMaps.xml`. Maps are usually named after the
root element of the schema, e.g. `root_Map`, with further maps of the same root numbered (`root_Map1`).

## API

The maps of a workbook are reached through the extension method `IXLWorkbook.XmlMaps()`. It is an extension method,
so `IXLWorkbook` stays unchanged.

### `IXLXmlMaps`

| Member | Description |
|---|---|
| `int Count` | Number of maps. |
| `IXLXmlMap Get(string name)` | Map by name. Throws `KeyNotFoundException` if there is none. |
| `bool TryGet(string name, out IXLXmlMap map)` | Map by name. |
| `void Delete(IXLXmlMap map)` | Same as `map.Delete()`. |
| `IEnumerator<IXLXmlMap>` | Enumerates the maps. |

### `IXLXmlMap`

| Member | Description |
|---|---|
| `uint Id` | Id of the map, unique in the workbook. |
| `string Name` | Name of the map. |
| `string RootElement` | Local name of the root element. |
| `string SchemaId`, `string? SchemaNamespace`, `string SchemaXml` | The schema embedded in the workbook. |
| `string SelectionNamespaces` | Prefixes used in the XPaths of the mappings, e.g. `xmlns:ns1='urn:a'`. There is one list for all maps of a workbook. |
| `bool ShowImportExportValidationErrors` | Kept in the file. It doesn't change the import or export: data is always validated, see [Validation](#validation). |
| `bool AdjustColumnWidth` | Adjust the width of mapped columns after an import. |
| `bool AppendOnImport` | Append rows to XML tables instead of replacing them. |
| `bool PreserveColumnFilter` | Kept in the file. Filters are applied again in any case, see [Filters and sort order](#filters-and-sort-order). |
| `bool PreserveNumberFormatting` | Keep the number formats of mapped cells. |
| `XLXmlImportResult Import(Stream stream, bool overwrite = true)` | Import XML from a stream. |
| `XLXmlImportResult Import(string path, bool overwrite = true)` | Import XML from a file. |
| `XLXmlImportResult ImportXml(string xml, bool overwrite = true)` | Import XML from a string. |
| `bool IsExportable` | Can the map be exported? See [Maps that can't be exported](#maps-that-cant-be-exported). |
| `XLXmlExportResult ExportXml(out string xml)` | Export as a string. |
| `XLXmlExportResult Export(Stream stream)` | Export as UTF-8 (without byte order mark) into a stream. |
| `XLXmlExportResult Export(string path, bool overwrite = false)` | Export into a file. Without `overwrite`, an existing file causes an `IOException`. |
| `void Delete()` | Remove the map, see [Delete](#delete). |

`overwrite: false` on import is the same as `AppendOnImport = true`.

### Results and errors

`XLXmlImportResult`:

| Value | Meaning |
|---|---|
| `Success` (0) | Imported. |
| `ElementsTruncated` (1) | Imported, but a text was longer than the 32,767 characters of a cell and was truncated. |
| `ValidationFailed` (2) | The XML doesn't match the schema. **The data is imported anyway**. |

`XLXmlExportResult`:

| Value | Meaning |
|---|---|
| `Success` (0) | Exported. |
| `ValidationFailed` (1) | The exported XML doesn't match the schema (e.g. a text in a number column or an empty required element). **The XML is exported anyway**. |

When an operation is refused, ClosedXML throws an exception and leaves the workbook unchanged:

| Situation | Exception |
|---|---|
| The XML isn't well-formed, e.g. an unclosed element or an undeclared prefix. | `System.Xml.XmlException` |
| Appending (`overwrite: false` or `AppendOnImport`) to a map without XML tables. | `InvalidOperationException` |
| A sheet with mapped cells or tables of the map is protected. | `InvalidOperationException` |
| An XML table can't grow, see [Growing tables](#growing-tables). | `InvalidOperationException` |
| Export of a map that isn't exportable. | `InvalidOperationException` |
| Export of a map without single cells whose tables have no data rows. | `InvalidOperationException` |
| Import, export or delete of a deleted map. | `InvalidOperationException` |

The import is checked completely before anything changes: when one table can't grow, none of the cells and tables
of the map is changed.

## Import

### Validation

The data is always validated against the schema. The setting `ShowImportExportValidationErrors` is kept in the file,
but doesn't change the validation. A validation error
doesn't stop the import: everything that can be found is imported and the result is `ValidationFailed`.

ClosedXML validates with `System.Xml.Schema` against the schemas embedded in the workbook. The XML is parsed without
DTD processing and without resolving external resources. Notable cases:

- **Root element:** a root element in another namespace, without namespace or with another name is invalid. The cells
  of the map are cleared, because none of the mapped elements are found.
- **Unknown elements and attributes, wrong order, duplicates:** invalid. Elements are found by their path anyway. Of
  duplicated single elements, the first one is used.
- **Prefixes:** prefixes in the data don't matter, only the namespaces.
- **Default values** of attributes in the schema are ignored. A missing attribute gives an empty cell, not the
  default value.
- **XML that isn't well-formed** is refused completely.

### Single cells

A mapped cell gets the value of the element or attribute of its XPath:

- A missing element and an element with `xsi:nil="true"` clear the cell.
- An existing formula in the cell is replaced.
- On append, single cells aren't changed.

### Values

Values are converted according to the XSD data type of the mapping (`xmlDataType`). For derived types (enumerations,
lists, unions, facets), the mapping stores the built-in base type, e.g. `string` for an enumeration of strings.

| XSD types | Cell value |
|---|---|
| `byte`, `short`, `int`, `long`, `integer` and their unsigned / positive / negative variants, `decimal`, `float`, `double` | Number |
| `boolean` | `true` / `false` become logical values, any other number (`1`, `0`, `-1`, `1.5`, `1e1`) stays a number |
| `date`, `dateTime` | Date (respecting the 1904 date system of the workbook) |
| `time` | Time |
| `gYear` | Number, if it has at least four digits (`0099`, `-0001`, not `99` or `0000`) or a plus sign (`+999`) |
| `gYearMonth` | Date of the first day of the month |
| `gMonthDay` | Day of the year 1900, see below |
| `gDay`, `gMonth` | Number for `---01` … `---31` and `--01--` … `--12--`; `--05` (XML Schema 1.1 notation) stays text |
| everything else (`string`, `token`, `duration`, `base64Binary`, …) | Text |

- A value that can't be converted (e.g. `abc` in an `int` element) stays text and the import fails validation.
- Number spellings like `+12`, `.5`, `1.`, `1E3`, `007` are numbers. `0x10`, `INF`, `NaN`, `1 000`, `1,5`, full-width
  digits stay text. So do numbers out of the range of a double: `1e400`, `1e-400`, `0.0E-999`. A negative underflow like
  `-1E-400` is taken as 0.
- Times may leave out the seconds (`10:00`, `2020-01-01T10:00`). `24:00` is the end of the day: the time `1`, or the
  next day in a `dateTime`.
- Dates and times with a time zone, and times with a non-zero fraction of a second, stay text.
- `gMonthDay` is the day in the year 1900 with a February of 29 days (the 1900 date system of ECMA-376 treats 1900
  as a leap year), and the day isn't checked against the month: `--02-29` is 60, `--02-31` is 62 (the 2nd of March).
  In the 1904 date system it is one less, so `--01-01` is 0.
- Leading and trailing whitespace is removed, but whitespace inside a CDATA section is kept.
- Numbers are parsed independently of the culture of the thread.

### Number formats

- With `PreserveNumberFormatting`, the existing number format of a cell is kept.
- A cell with the *General* format gets the default format of the data type: `mm-dd-yy` for `date`,
  `m/d/yy h:mm` for `dateTime`, `h:mm:ss` for `time`, `@` (text) for string types, and so on.
- For dates and times, a *built-in* format that doesn't fit the value is replaced as well. A value of 0 (only possible
  in the 1904 date system, e.g. `1904-01-01`) gets `h:mm:ss` instead of a date format. A value of a day or more in a
  time format (`24:00:00`, or a date in a cell formatted `h:mm`) gets `m/d/yyyy`, or `m/d/yyyy h:mm` when it has a time
  of day. Custom formats (`yyyy-mm-dd`, `[h]:mm`) and non-date formats (`0.00`) are kept.
- New rows of a table get the format of the first data row of their column.

### Tables

The rows of an XML table are the repeating element of the column XPaths, e.g. `/root/persons/person`. A column gets
the value of the element or attribute relative to the row element. A missing element gives an empty cell. Repeating
elements without any mapped data don't become rows.

**Denormalized columns and nested lists.** A table can contain columns of elements above the row element, e.g. the
customer of all orders or the id of an order next to the lines of the order. The nesting is flattened like an outer
join:

- every element of the innermost repeating level is a row;
- values of outer levels are repeated in each row of their element;
- an outer element without inner elements still gives one row, e.g. an order without lines.

Repeating levels without mapped columns count too. Such maps can be imported, but not exported.

**Size of the table.** On overwrite, the data rows are resized to fit the data; an empty table keeps one empty row.
On append, the rows are added after the existing ones. When an import empties a table that had data, the table
gets an *insert row* (`insertRow`); ClosedXML keeps track of it, because it changes later appends and exports.

**Calculated columns** (columns that aren't mapped and have the same formula in all data rows) get the formula in
added rows. The totals row moves with the table.

### Growing tables

A table grows downward only, never sideways:

1. The table first uses empty rows below it. Formatting alone doesn't make a row non-empty, merged cells do.
2. When something is in the way, the cells are shifted down — only in the columns of the table, from the first
   non-empty row to the end of the sheet. Cells left and right of the table stay where they are.
3. Merged cells, array formulas and other tables are shifted along when they lie completely within the columns of the
   table.

The import is refused (see [Results and errors](#results-and-errors)) when the shift would:

- split another table, merged cells or an array formula that reaches beyond the columns of the table;
- shift non-empty cells off the sheet, or the table would grow past row 1,048,576.

### Filters and sort order

When a table has an autofilter with conditions, the filter is applied again to the imported rows: rows of the old
data range are shown first, then rows not matching are hidden. This happens whatever the value of
`PreserveColumnFilter`. The sort order is **not** applied again: the rows stay in the order of the XML. The sort state
of the table is kept.

### Column widths

With `AdjustColumnWidth`, the mapped columns are adjusted to their content after the import. The widths are estimated
from the font and the content, so they are approximate.

### Formulas

Formulas in mapped cells are replaced by the imported values. Formulas elsewhere that depend on mapped cells (e.g.
`=SUM(Table1[qty])`) stay as they are. ClosedXML saves formulas without calculated values; they are calculated by the
application that opens the workbook. Programs that read the file without calculating (e.g. openpyxl, previews) see no
values for these formulas.

## Export

The export produces the XML from the current content of the mapped cells, so it also contains values that were
entered or changed after an import. The exported XML has a fixed format:

- XML declaration with `standalone="yes"`, indentation with tabs, lines separated by CR LF, no line break at the end;
- namespaces are declared on the root element with the prefixes of the map (`ns1`, …); `xsi` is declared when a
  mapped element is `nillable`;
- elements in the order of the schema, attributes in the order of their declaration.

**Empty values.**

- An empty cell of an optional element or attribute leaves it out.
- An empty cell of a required element gives an empty element (`<ns1:name/>`), or `xsi:nil="true"` for a nillable
  element.
- A required attribute is written empty.
- An error value (e.g. `#DIV/0!`) gives an empty element.
- Elements without content are left out, unless they are required or contain rows of a table.

**Rows.**

- Every data row of a table is written, even an empty row (`<ns1:row/>`).
- A table emptied by an import (with only the insert row) has no rows.

**Values.**

| Cell value | XSD type | Exported text |
|---|---|---|
| Number | `decimal`, `integer`, `int`, `long`, … (decimal types) | Rounded to 15 significant digits, always without exponent: `12345678901234600000`, `0.00000000000000000001` |
| Number | `double`, `float`, string types and all others | Rounded to 15 significant digits; without exponent if that fits into 21 characters (without sign), otherwise with exponent: `100000000000000000000`, `1E+21`, `0.0000000000000000001`, `1E-20`, `1.23456789012345E-07` |
| Number / date | `date` | `yyyy-mm-dd` |
| Number / date | `dateTime` | `yyyy-mm-ddThh:mm:ss.fff` |
| Number / time | `time` | `hh:mm:ss.fff` of the fraction of the day |
| Number / date | `gYearMonth`, `gMonthDay` | `yyyy-mm`, `--mm-dd` of the date |
| Number | `gYear` | Whole numbers with at least four digits: `2020`, `0999`, `-0005` |
| Number | `gDay`, `gMonth` | `---01` … `---31`, `--01--` … `--12--` (the notation of XML Schema 1.0 drafts) |
| Logical | any | `true` / `false` |
| Text | any | As is, line breaks as CR LF; `&`, `<` and `>` escaped |

**The number format doesn't matter.** The value of a cell is exported, not the displayed text: `51.50%` is
exported as `0.515`, `0042` (format `0000`) as `42`, `€1 234,50` as `1234.5`. This also applies to cells with the Text
format (`@`) that contain a number. Only text stays text: `00042` entered as text in a Text cell is exported as
`00042`. The displayed text is never exported, also not for the Text format.

**Numbers in date and time types.**

- Values are rounded to milliseconds before date and time are split, so a time just before midnight belongs to the
  next day: 45000.9999999999 is `2023-03-16T00:00:00.000` and, as a `time`, `00:00:00.000`.
- Dates follow the date system of the workbook, including 29 February 1900 (serial 60), which exists because the
  1900 date system treats 1900 as a leap year (ECMA-376); it is exported as `1900-02-29`.
- Negative serial numbers count back from the start of the date system. In the 1900 system, 0 is `1899-12-31` and -1 is
  `1899-12-30`. In the 1904 system, negative numbers are a day later: -1 and -0.5 are both `1904-01-01`.
- Dates are exported from `1499-12-31` to `9999-12-31` (1900 system; serial -146087 to below 2958466) and from
  `1504-01-01` to `9999-12-31` (1904 system; serial -146098 to below 2957004). Numbers outside are exported as numbers.
  The range also applies to `time`, `gYearMonth` and `gMonthDay`. In the 1900 system, ten days are skipped before
  `1500-03-02`.
- A number that isn't a valid `gYear`, `gMonth` or `gDay` (0, 13, 31.5, 45000) is exported as a number with 17
  significant digits: `45000`, `31.000000000000099`, `1E+20`.
- Numbers smaller than about 2.2251E-308 (the smallest normalized double) are exported as `0`.

### Maps that can't be exported

`IsExportable` is `false` and the export is refused when the XML structure can't be restored from the cells:

- a table contains a list of lists: its row element is inside another repeating element;
- a table contains denormalized data: a column of an element outside the row element;
- elements of the same repeating element are mapped to different tables;
- a mapped element is part of an `xs:choice`;
- a mapped element is part of a repeating group with more than one particle (e.g. `<xs:sequence maxOccurs="unbounded">`).

A map without single cells needs at least one table with data rows, otherwise the export is
refused. A map with single cells is exported even when all cells are empty (as an empty root element).

## Delete

`map.Delete()` removes the map:

- the map, its mapped single cells and the XML properties of its table columns are removed;
- the values stay in the cells;
- XML tables without remaining mappings become normal tables;
- the `xmlMaps.xml` part is removed when the last map is deleted.

Schemas still used by other maps are kept.

## Loading and saving

The XML map parts are read when a workbook is loaded and always written again from the loaded data when it is
saved. Parts left over from the source file are never trusted. This works for saving into the loaded file, for
`SaveAs` to a new file and for workbooks created with `XLWorkbook.OpenFromTemplate`.

| What | Where in the package | Handling |
|---|---|---|
| Maps and schemas | `xl/xmlMaps.xml` | Read and written as plain XML (LINQ to XML), because the OpenXML SDK loses namespace declarations of embedded schemas. |
| Mapped single cells | `xl/tables/tableSingleCells*.xml` per sheet | Written for every sheet with mapped cells, removed otherwise. Ids are taken from the same counter as table ids, so they are unique in the workbook. |
| Table columns | `tableType="xml"`, `tableColumn/@uniqueName`, `xmlColumnPr` | Added to the tables ClosedXML writes. Columns are matched by name first, then by position. Calculated columns keep their `uniqueName`. |
| Sort state of XML tables | `table/sortState` | Kept (ClosedXML otherwise doesn't keep the sort state of tables). |
| Data bindings | `Map/DataBinding`, `table/@connectionId` | Kept while the package contains the connection (`connections.xml`), left out otherwise, e.g. after `OpenFromTemplate`. |

Mapped cells move with inserted and deleted rows and columns. Mappings of deleted tables and sheets are dropped;
the map itself stays.

## Limitations

**Out of scope**

- **No user interface.** Results are returned and errors are thrown as exceptions; there are no dialogs.
- **No data bindings are created.** An import doesn't bind the map to the imported file through a data connection.
  ClosedXML doesn't create or change connections; existing bindings are kept.
- **No calculation.** Formulas are saved without values, they are calculated when the workbook is opened (see
  [Formulas](#formulas)).
- **No map creation.** Maps must exist in the loaded workbook.

**Validator restrictions** (marked in the tests as known deviations, the tests fail when the behavior changes)

- The .NET validator rejects `24:00:00` as `xs:time` (result `ValidationFailed`). The value (the time `1`) and the
  number format are imported anyway.
- The .NET validator rejects negative years (e.g. `-0001`) as `xs:gYear`.
- The .NET validator rejects an `xs:token` consisting only of whitespace.

**Other limitations**

- Filters of two XML tables of the same map in the same worksheet rows may hide unexpected rows, because the result
  depends on the order in which the tables are filtered.
- ClosedXML turns array formulas into normal formulas when it shifts their cells (an upstream issue). The import
  restores the array formulas it shifts itself.
- When a date or time is imported into a cell with the Text format (`@`), ClosedXML stores the number and keeps the
  Text format. It doesn't convert the date into locale-dependent text.
- ClosedXML saves the largest double (1.7976931348623157E+308) rounded to 15 digits, which is out of range when the
  file is read again (an upstream issue).
- XML maps aren't supported by all spreadsheet applications.
- Schema constructs that can't be mapped to cells (`xs:any`, `xs:anyAttribute`, recursion deeper than one level,
  abstract elements, substitution groups, mixed content) are supported only as far as the map in the workbook refers
  to them.

## Architecture

All code is in `ClosedXML/Excel/XmlMaps`:

| File | Content |
|---|---|
| `IXLXmlMap.cs`, `XLXmlMap.cs` | Public map interface and implementation. |
| `IXLXmlMaps.cs`, `XLXmlMaps.cs` | The maps of a workbook. Map creation would be an `Add` method here. |
| `XLWorkbookXmlMapExtensions.cs` | `IXLWorkbook.XmlMaps()`. |
| `XLWorkbook_XmlMaps.cs` | Partial `XLWorkbook`: the store field and the load hooks. |
| `XLXmlMapStore.cs` | Per-workbook state: maps, schemas, cell bindings and table bindings (`ConditionalWeakTable`, so `XLTable` needs no change). |
| `XLXmlCellBinding.cs`, `XLXmlTableBinding.cs` | A mapped cell (tracked range) and the XML properties of a table. |
| `XLXmlSchema.cs`, `XmlMapSchema.cs` | Embedded schema and its compiled form (validation, repeating elements, declarations). |
| `XmlMapImporter.cs` | Import: validation, flattening of rows, planning and checking table growth, writing cells. |
| `XmlMapExporter.cs` | Export: exportability, building the XML in schema order, serialization in the export format. |
| `XmlDataTypeConverter.cs` | XSD text ↔ cell value. |
| `XLXmlImportResult.cs`, `XLXmlExportResult.cs` | Results. |
| `IO/XmlMapPartXml.cs` | Reading and writing `xmlMaps.xml`. |
| `IO/XmlMapReader.cs`, `IO/XmlMapWriter.cs` | Reading and writing mapped cells and XML properties of tables. |

Only `System.Xml` is used, there are no additional dependencies.

**Prepared for map creation.** Bindings refer to map objects, not to ids, and are created through internal
constructors. Creating maps would need `IXLXmlMaps.Add(xsd, rootElement)` and methods to bind a cell or a table
column to an XPath; loading, saving, import and export work with them unchanged.

### Merging upstream

Upstream files contain only these hook calls:

| File | Hook |
|---|---|
| `Excel/XLWorkbook_Load.cs` | `LoadXmlMaps(workbookPart)`, `LoadXmlMapTable(xlTable, dTable)`, `LoadXmlMapCells(worksheetPart, ws)` |
| `Excel/XLWorkbook_Save.cs` | `XmlMapWriter.WriteCellBindings(worksheetPart, worksheet, context)` after the table parts, `XmlMapWriter.WriteMaps(workbookPart, this, context)` after the worksheets |
| `Excel/IO/TablePartWriter.cs` | `XmlMapWriter.ApplyTableBinding(tableDefinitionPart, table, xlTable)` before the table is assigned to the part |

When a merge conflicts in these places, keep the hook call at the equivalent position. Afterwards run both test
projects (see below). The XML map code uses internal members of ClosedXML (e.g. `XLCell.SetOnlyValue`,
`XLTableField.Column`, `XLAutoFilter.Columns`), so compile errors after a merge point to changed internals.

## Tests

`ClosedXML.XmlMaps.Tests` (`net8.0-windows`, NUnit) contains two kinds of tests:

- **Unit tests** (`ImportTests`, `ExportTests`, `RoundTripTests`): API, culture independence, round trips.
- **Comparison tests** (`ExcelComparisonTests`, category `Excel`): the same operation is done by a reference
  application through COM and by ClosedXML, and the results are compared. These tests are explicit, they only run
  when selected, and are skipped when no reference application is installed.

```powershell
# unit tests
dotnet test ClosedXML.XmlMaps.Tests

# comparison tests (about one minute)
dotnet test ClosedXML.XmlMaps.Tests --filter "FullyQualifiedName~ExcelComparisonTests"
```

A comparison test:

1. lets the reference application build the template from a scenario (`Scenarios/Scenarios.cs`: schema, mapped cells, tables, formats,
   filters, …), so there are no hand-made test files;
2. imports the XML of the case (`Scenarios/ImportCases.cs`) in the reference application and in ClosedXML and saves both workbooks;
3. compares import results, maps, mappings, cell values, types, number formats, tables, filters, sort state and hidden
   rows (`Infrastructure/WorkbookSnapshot.cs`, `SnapshotComparer.cs`);
4. compares `IsExportable`, the export result and the exported XML of every map;
5. opens the ClosedXML workbook in the reference application: it must contain the same maps and mappings, export the same XML and
   calculate the same formula values.

The generated files of each test are in `bin/Debug/net8.0-windows/Generated/<test name>`, which helps to analyze
differences.

**Adding a case:** add an `ImportCase` to one of the lists in `ImportCases.cs`. A new layout needs a scenario in
`Scenarios.cs`, built with `ExcelWorkbookBuilder` (maps, mapped cells, tables, formulas, number formats; anything else
through COM). A difference that is accepted is marked with `KnownDeviation`.

The upstream test project `ClosedXML.Tests` must stay green for `net8.0` and `net462`.
