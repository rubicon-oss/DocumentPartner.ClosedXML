using System;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.XmlMaps.Tests.Infrastructure
{
    /// <summary>
    /// Access to embedded test resources and to the output directory of a test.
    /// </summary>
    internal static class TestFiles
    {
        /// <summary>
        /// Save options for tests. Package validation is off, because the OpenXML SDK rejects some parts written
        /// by Excel (e.g. newer chart elements) that ClosedXML just keeps.
        /// </summary>
        internal static SaveOptions SaveOptions => new() { ValidatePackage = false };

        internal static Stream Open(string resourceName)
        {
            var assembly = typeof(TestFiles).Assembly;
            var fullName = assembly.GetManifestResourceNames()
                               .SingleOrDefault(n => n.EndsWith(".Resources." + resourceName, StringComparison.Ordinal))
                           ?? throw new FileNotFoundException($"Resource '{resourceName}' not found.");
            return assembly.GetManifestResourceStream(fullName)!;
        }

        internal static string ReadText(string resourceName)
        {
            using var reader = new StreamReader(Open(resourceName));
            return reader.ReadToEnd();
        }

        /// <summary>
        /// Directory for files of the current test. Files are kept after the test for inspection.
        /// </summary>
        internal static string OutputDirectory()
        {
            var test = TestContext.CurrentContext.Test;
            var name = string.Concat((test.ClassName?.Split('.').Last() + "." + test.Name).Select(c => Path.GetInvalidFileNameChars().Contains(c) ? '_' : c));
            var directory = Path.Combine(TestContext.CurrentContext.WorkDirectory, "Generated", name);
            if (Directory.Exists(directory))
                Directory.Delete(directory, true);

            Directory.CreateDirectory(directory);
            return directory;
        }

        /// <summary>
        /// Copy an embedded resource into the <paramref name="directory"/>.
        /// </summary>
        internal static string Extract(string resourceName, string directory, string? fileName = null)
        {
            var path = Path.Combine(directory, fileName ?? resourceName);
            using var source = Open(resourceName);
            using var target = File.Create(path);
            source.CopyTo(target);
            return path;
        }
    }
}
