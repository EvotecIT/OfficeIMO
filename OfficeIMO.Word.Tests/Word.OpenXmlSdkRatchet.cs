#if NET
using OfficeIMO.TestSupport.OpenXml;
using Xunit;
using Xunit.Abstractions;

namespace OfficeIMO.Tests {
    /// <summary>
    /// Tracks how much of OfficeIMO.Word still uses the Open XML SDK and <c>System.IO.Packaging</c>, the package API the
    /// SDK brings in. Every use of either belongs in the bridge folder (<c>OfficeIMO.Word/OpenXmlSdk</c>); these count the
    /// source files outside it whose code uses their types. The counts may only go down: lower the baseline in the change
    /// that removes files from a list.
    /// </summary>
    public class WordOpenXmlSdkRatchet {
        private const int Baseline = 443;
        private const int PackagingBaseline = 9;
        private readonly ITestOutputHelper _output;

        public WordOpenXmlSdkRatchet(ITestOutputHelper output) => _output = output;

        [Fact]
        public void FilesOutsideTheBridgeThatUseTheSdkOnlyGoDown() {
            string projectDirectory = Path.Combine(FindRepositoryRoot(), "OfficeIMO.Word");
            IReadOnlyList<string> files = OpenXmlSdkReferenceCounter.FindFiles(
                typeof(OfficeIMO.Word.WordDocument).Assembly.Location, projectDirectory, "OpenXmlSdk");

            _output.WriteLine($"OfficeIMO.Word files outside OpenXmlSdk/ that use the Open XML SDK: {files.Count} (baseline {Baseline})");
            foreach (string file in files) _output.WriteLine(file);

            Assert.True(files.Count <= Baseline,
                $"{files.Count} files outside OfficeIMO.Word/OpenXmlSdk use the Open XML SDK, more than the baseline of {Baseline}. Keep new SDK use in the bridge.");
            Assert.True(files.Count == Baseline,
                $"Only {files.Count} files outside OfficeIMO.Word/OpenXmlSdk use the Open XML SDK now. Lower the baseline in {nameof(WordOpenXmlSdkRatchet)} to {files.Count}.");
        }

        [Fact]
        public void FilesOutsideTheBridgeThatUseSystemIOPackagingOnlyGoDown() {
            string projectDirectory = Path.Combine(FindRepositoryRoot(), "OfficeIMO.Word");
            IReadOnlyList<string> files = OpenXmlSdkReferenceCounter.FindFiles(
                typeof(OfficeIMO.Word.WordDocument).Assembly.Location, projectDirectory, "OpenXmlSdk",
                OpenXmlSdkReferenceCounter.Dependency.SystemIOPackaging);

            _output.WriteLine($"OfficeIMO.Word files outside OpenXmlSdk/ that use System.IO.Packaging: {files.Count} (baseline {PackagingBaseline})");
            foreach (string file in files) _output.WriteLine(file);

            Assert.True(files.Count <= PackagingBaseline,
                $"{files.Count} files outside OfficeIMO.Word/OpenXmlSdk use System.IO.Packaging, more than the baseline of {PackagingBaseline}. Use OfficeIMO's own package layer, or keep it in the bridge.");
            Assert.True(files.Count == PackagingBaseline,
                $"Only {files.Count} files outside OfficeIMO.Word/OpenXmlSdk use System.IO.Packaging now. Lower {nameof(PackagingBaseline)} in {nameof(WordOpenXmlSdkRatchet)} to {files.Count}.");
        }

        private static string FindRepositoryRoot() {
            string? directory = AppContext.BaseDirectory;
            while (directory != null) {
                if (File.Exists(Path.Combine(directory, "OfficeIMO.Word", "OfficeIMO.Word.csproj"))) return directory;
                directory = Path.GetDirectoryName(directory);
            }
            throw new InvalidOperationException("The repository root was not found above the test output.");
        }
    }
}
#endif
