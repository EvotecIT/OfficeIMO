#if NET
using OfficeIMO.TestSupport.OpenXml;
using Xunit;
using Xunit.Abstractions;

namespace OfficeIMO.Tests {
    /// <summary>
    /// Tracks how much of OfficeIMO.Excel still uses the Open XML SDK and <c>System.IO.Packaging</c>, the package API the
    /// SDK brings in. Every use of either belongs in the bridge folder (<c>OfficeIMO.Excel/OpenXmlSdk</c>); these count the
    /// source files outside it whose code uses their types. The counts may only go down: lower the baseline in the change
    /// that removes files from a list. The engine's own Open XML layer names neither.
    /// </summary>
    public class ExcelOpenXmlSdkBoundary {
        private const int Baseline = 468;
        private const int PackagingBaseline = 14;
        private readonly ITestOutputHelper _output;

        public ExcelOpenXmlSdkBoundary(ITestOutputHelper output) => _output = output;

        private static readonly string[] OwnedOpenXmlFiles = {
            "ExcelDeferredPackageParts.cs",
            "ExcelDeferredWorksheetWriter.cs",
            "ExcelDocument.CreateShell.cs",
            "ExcelDocument.DeferredLoad.cs",
            "ExcelDocument.DeferredSave.cs",
            "ExcelDocument.DeferredSheets.cs",
            "ExcelOwnedStyles.cs",
            "ExcelSheet.Deferred.cs",
            "ExcelSheet.DeferredEdits.cs",
            "ExcelStyleOperation.cs",
        };

        [Fact]
        public void FilesOutsideTheBridgeThatUseTheSdkOnlyGoDown() {
            IReadOnlyList<string> files = FindFiles(OpenXmlSdkReferenceCounter.Dependency.OpenXmlSdk);

            _output.WriteLine($"OfficeIMO.Excel files outside OpenXmlSdk/ that use the Open XML SDK: {files.Count} (baseline {Baseline})");
            foreach (string file in files) _output.WriteLine(file);

            Assert.True(files.Count <= Baseline,
                $"{files.Count} files outside OfficeIMO.Excel/OpenXmlSdk use the Open XML SDK, more than the baseline of {Baseline}. Keep new SDK use in the bridge.");
            Assert.True(files.Count == Baseline,
                $"Only {files.Count} files outside OfficeIMO.Excel/OpenXmlSdk use the Open XML SDK now. Lower the baseline in {nameof(ExcelOpenXmlSdkBoundary)} to {files.Count}.");
        }

        [Fact]
        public void FilesOutsideTheBridgeThatUseSystemIOPackagingOnlyGoDown() {
            IReadOnlyList<string> files = FindFiles(OpenXmlSdkReferenceCounter.Dependency.SystemIOPackaging);

            _output.WriteLine($"OfficeIMO.Excel files outside OpenXmlSdk/ that use System.IO.Packaging: {files.Count} (baseline {PackagingBaseline})");
            foreach (string file in files) _output.WriteLine(file);

            Assert.True(files.Count <= PackagingBaseline,
                $"{files.Count} files outside OfficeIMO.Excel/OpenXmlSdk use System.IO.Packaging, more than the baseline of {PackagingBaseline}. Use OfficeIMO's own package layer, or keep it in the bridge.");
            Assert.True(files.Count == PackagingBaseline,
                $"Only {files.Count} files outside OfficeIMO.Excel/OpenXmlSdk use System.IO.Packaging now. Lower {nameof(PackagingBaseline)} in {nameof(ExcelOpenXmlSdkBoundary)} to {files.Count}.");
        }

        // The owned files still call engine members whose signatures name SDK types (the replay into the model), so the
        // counts above include them; their own text names neither namespace.
        [Fact]
        public void TheOwnedOpenXmlLayerNamesNeitherTheSdkNorSystemIOPackaging() {
            string root = Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.Excel");
            foreach (string file in OwnedOpenXmlFiles) {
                string text = File.ReadAllText(Path.Combine(root, file));
                Assert.False(OpenXmlSdkReferenceCounter.Dependency.OpenXmlSdk.NamespaceText.IsMatch(text), file + " names the Open XML SDK.");
                Assert.False(OpenXmlSdkReferenceCounter.Dependency.SystemIOPackaging.NamespaceText.IsMatch(text), file + " names System.IO.Packaging.");
            }
        }

        private static IReadOnlyList<string> FindFiles(OpenXmlSdkReferenceCounter.Dependency dependency) =>
            OpenXmlSdkReferenceCounter.FindFiles(typeof(OfficeIMO.Excel.ExcelDocument).Assembly.Location,
                Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.Excel"), "OpenXmlSdk", dependency);
    }
}
#endif
