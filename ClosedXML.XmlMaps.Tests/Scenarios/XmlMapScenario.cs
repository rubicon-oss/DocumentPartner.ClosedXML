using System;
using ClosedXML.XmlMaps.Tests.Infrastructure;

namespace ClosedXML.XmlMaps.Tests.Scenarios
{
    /// <summary>
    /// A workbook with XML maps used as template for tests. It is either an embedded resource or built by Excel.
    /// </summary>
    public sealed class XmlMapScenario
    {
        private XmlMapScenario(string name, string? resource, Action<ExcelWorkbookBuilder>? build, bool compareAllCells)
        {
            Name = name;
            Resource = resource;
            Build = build;
            CompareAllCells = compareAllCells;
        }

        public string Name { get; }

        /// <summary>
        /// Compare all cells with a value, not only the cells mapped by XML maps. Off for workbooks with
        /// content ClosedXML doesn't keep (e.g. output of pivot tables).
        /// </summary>
        internal bool CompareAllCells { get; }

        /// <summary>
        /// Embedded xlsx resource used as template.
        /// </summary>
        internal string? Resource { get; }

        /// <summary>
        /// Builds the template in Excel.
        /// </summary>
        internal Action<ExcelWorkbookBuilder>? Build { get; }

        internal static XmlMapScenario FromResource(string resource, bool compareAllCells = true) => new(resource, resource, null, compareAllCells);

        internal static XmlMapScenario FromExcel(string name, Action<ExcelWorkbookBuilder> build) => new(name, null, build, true);

        public override string ToString() => Name;
    }

    /// <summary>
    /// One import of XML data into a map.
    /// </summary>
    public sealed record XmlImport(string MapName, string Xml, bool Overwrite = true)
    {
        public override string ToString() => $"{MapName}{(Overwrite ? string.Empty : " (append)")}";
    }

    /// <summary>
    /// A test case: imports applied in order to a template, once by Excel and once by ClosedXML.
    /// </summary>
    public sealed record ImportCase(string Name, XmlMapScenario Scenario, params XmlImport[] Imports)
    {
        /// <summary>
        /// A difference to Excel that is accepted. The test reports a warning instead of failing,
        /// and fails if the difference disappeared, so the marker gets removed.
        /// </summary>
        public string? KnownDeviation { get; init; }

        public override string ToString() => Name;
    }
}
