using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.XmlMaps.Tests
{
    /// <summary>
    /// XML maps are opt-in, enable them for all tests of the assembly.
    /// </summary>
    [SetUpFixture]
    public class XmlMapsSetUp
    {
        [OneTimeSetUp]
        public void EnableXmlMaps() => XLWorkbook.EnableXmlMaps = true;

        [OneTimeTearDown]
        public void DisableXmlMaps() => XLWorkbook.EnableXmlMaps = false;
    }
}
