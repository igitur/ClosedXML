// Keep this file CodeMaid organised and cleaned
using ClosedXML.Excel;
using ClosedXML.Examples;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using Path = System.IO.Path;

namespace ClosedXML.Tests
{
    internal static class TestHelper
    {
        public const string ActualTestResultPostFix = "";

        public static readonly string ExampleTestsOutputDirectory = Path.Combine(TestsOutputDirectory, "Examples");

        private const bool CompareWithResources = true;

        private static readonly ResourceFileExtractor _extractor = new ResourceFileExtractor(".Resource.");

        public static string CurrencySymbol
        {
            get { return Thread.CurrentThread.CurrentCulture.NumberFormat.CurrencySymbol; }
        }

        public static bool IsRunningOnUnix
        {
            get
            {
                int p = (int)Environment.OSVersion.Platform;
                return ((p == 4) || (p == 6) || (p == 128));
            }
        }

        // Because different fonts are installed on Unix,
        // the columns widths after AdjustToContents() will
        // cause the tests to fail.
        // Therefore we ignore the width attribute when running on Unix
        public static bool StripColumnWidths { get { return IsRunningOnUnix; } }

        //Note: Run example tests parameters
        public static string TestsOutputDirectory
        {
            get
            {
                return Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "Generated");
            }
        }

        public static void CreateAndCompare(Func<IXLWorkbook> workbookGenerator, string referenceResource, bool evaluateFormulae = false)
        {
            Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");

            string[] pathParts = referenceResource.Split(new char[] { '\\' });
            string filePath1 = Path.Combine(new List<string>() { TestsOutputDirectory }.Concat(pathParts).ToArray());

            var extension = Path.GetExtension(filePath1);
            var directory = Path.GetDirectoryName(filePath1);

            var fileName = Path.GetFileNameWithoutExtension(filePath1);
            fileName += ActualTestResultPostFix;
            fileName = Path.ChangeExtension(fileName, extension);

            var filePath2 = Path.Combine(directory, fileName);

            using (var wb = workbookGenerator.Invoke())
                wb.SaveAs(filePath2, true, evaluateFormulae);

            if (CompareWithResources)
            {
                CompareFiles(filePath2, referenceResource.Replace('\\', '.').TrimStart('.'));
            }
        }

        public static string GetResourcePath(string filePartName)
        {
            return filePartName.Replace('\\', '.').TrimStart('.');
        }

        public static Stream GetStreamFromResource(string resourcePath)
        {
            return _extractor.ReadFileFromResourceToStream(resourcePath);
        }

        public static IEnumerable<String> ListResourceFiles(Func<String, Boolean> predicate = null)
        {
            return _extractor.GetFileNames(predicate);
        }

        public static void LoadFile(string filePartName, LoadOptions loadOptions = null)
        {
            loadOptions = loadOptions ?? new LoadOptions();
            using var stream = GetStreamFromResource(GetResourcePath(filePartName));
            Assert.DoesNotThrow(() => new XLWorkbook(stream, loadOptions), "Unable to load resource {0}", filePartName);
        }

        public static void RunTestExample<T>(string filePartName, bool evaluateFormulae = false)
                where T : IXLExample, new()
        {
            // Make sure tests run on a deterministic culture
            Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");

            var example = new T();
            string[] pathParts = filePartName.Split(new char[] { '\\' });
            string filePath1 = Path.Combine(new List<string>() { ExampleTestsOutputDirectory }.Concat(pathParts).ToArray());

            var extension = Path.GetExtension(filePath1);
            var directory = Path.GetDirectoryName(filePath1);

            var fileName = Path.GetFileNameWithoutExtension(filePath1);
            fileName += ActualTestResultPostFix;
            fileName = Path.ChangeExtension(fileName, extension);

            filePath1 = Path.Combine(directory, "z" + fileName);
            var filePath2 = Path.Combine(directory, fileName);

            //Run test
            example.Create(filePath1);
            using (var wb = new XLWorkbook(filePath1))
                wb.SaveAs(filePath2, validate: true, evaluateFormulae);

            // Also load from template and save it again - but not necessary to test against reference file
            // We're just testing that it can save.
            using (var ms = new MemoryStream())
            using (var wb = XLWorkbook.OpenFromTemplate(filePath1))
                wb.SaveAs(ms, validate: true, evaluateFormulae);

            if (CompareWithResources)
            {
                CompareFiles(filePath2, "Examples." + filePartName.Replace('\\', '.').TrimStart('.'));
            }
        }

        public static void SaveWorkbook(XLWorkbook workbook, params string[] fileNameParts)
        {
            workbook.SaveAs(Path.Combine(new string[] { TestsOutputDirectory }.Concat(fileNameParts).ToArray()), true);
        }

        private static void CompareFiles(string filePath2, string resourcePath)
        {
            using (var streamExpected = _extractor.ReadFileFromResourceToStream(resourcePath))
            using (var streamActual = File.OpenRead(filePath2))
            {
                Assert.IsTrue(ExcelDocsComparer.Compare(streamActual, streamExpected, out string message),
                    $"Actual file `.\\{GetRelativePath(filePath2, Environment.CurrentDirectory)}` is different to the expected file `{resourcePath}`.\r\n{message}");
            }
        }

        private static string GetRelativePath(string filespec, string folder)
        {
            Uri pathUri = new Uri(filespec);
            // Folders must end in a slash
            if (!folder.EndsWith(Path.DirectorySeparatorChar.ToString()))
            {
                folder += Path.DirectorySeparatorChar;
            }
            Uri folderUri = new Uri(folder);
            return Uri.UnescapeDataString(folderUri.MakeRelativeUri(pathUri).ToString().Replace('/', Path.DirectorySeparatorChar));
        }
    }
}
