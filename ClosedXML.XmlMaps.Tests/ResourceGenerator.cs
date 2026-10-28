using System.IO;
using System.Linq;
using ClosedXML.XmlMaps.Tests.Infrastructure;
using NUnit.Framework;

namespace ClosedXML.XmlMaps.Tests
{
    /// <summary>
    /// Builds the xlsx resources used by tests that run without Excel. Run explicitly after changing a scenario,
    /// the files are written into the Resources folder of the source tree.
    /// </summary>
    [TestFixture]
    [Explicit("Regenerates test resources with Excel.")]
    [Category("Excel")]
    public class ResourceGenerator
    {
        [Test]
        public void GenerateResources()
        {
            var resources = Path.Combine(FindProjectDirectory(), "Resources");
            using var excel = ExcelSession.Start();
            ExcelWorkbookBuilder.Build(excel, Path.Combine(resources, "TwoMaps.xlsx"), Scenarios.Scenarios.TwoMaps.Build!);
            ExcelWorkbookBuilder.Build(excel, Path.Combine(resources, "Types.xlsx"), Scenarios.Scenarios.Types.Build!);
            ExcelWorkbookBuilder.Build(excel, Path.Combine(resources, "Tables.xlsx"), Scenarios.Scenarios.Tables.Build!);
            foreach (var schema in Directory.GetFiles(resources, "schema*.xsd"))
                File.Delete(schema);
        }

        private static string FindProjectDirectory()
        {
            var directory = new DirectoryInfo(TestContext.CurrentContext.TestDirectory);
            while (directory is not null && !directory.GetFiles("ClosedXML.XmlMaps.Tests.csproj").Any())
                directory = directory.Parent;

            return directory?.FullName ?? throw new DirectoryNotFoundException("Project directory not found.");
        }
    }
}
