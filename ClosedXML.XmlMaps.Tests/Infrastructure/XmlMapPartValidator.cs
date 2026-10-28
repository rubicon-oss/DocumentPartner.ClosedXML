using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using NUnit.Framework;

namespace ClosedXML.XmlMaps.Tests.Infrastructure
{
    /// <summary>
    /// Schema validation of the parts that contain XML map data, using the validator of the OpenXML SDK.
    /// The SDK also reports errors for files written by Excel (e.g. <c>mapId</c> references or the
    /// <c>uniqueName</c> of single cells), so only errors that the file written by Excel doesn't have are reported.
    /// </summary>
    internal static class XmlMapPartValidator
    {
        /// <param name="path">The file written by ClosedXML.</param>
        /// <param name="excelPath">The corresponding file written by Excel.</param>
        internal static void AssertValid(string path, string excelPath)
        {
            var excelErrors = Validate(excelPath).Select(e => e.Description).ToHashSet();
            var errors = Validate(path)
                .Where(e => !excelErrors.Contains(e.Description))
                .Select(e => $"{e.Part}: {e.Path}: {e.Description}")
                .ToList();
            Assert.That(errors, Is.Empty, "XML map parts must be valid.");
        }

        private static List<(string Part, string? Path, string Description)> Validate(string path)
        {
            using var document = SpreadsheetDocument.Open(path, false);
            var validator = new OpenXmlValidator(FileFormatVersions.Office2016);
            return validator.Validate(document)
                .Where(e => e.Part?.Uri.OriginalString is { } uri && (uri.Contains("xmlMaps") || uri.Contains("tableSingleCells") || uri.Contains("/tables/")))
                .Select(e => (e.Part!.Uri.OriginalString, e.Path?.XPath, e.Description))
                .ToList();
        }
    }
}
