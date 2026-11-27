#if NET
using OfficeIMO.TestSupport.OpenXml;
using Xunit;
using Xunit.Abstractions;

namespace OfficeIMO.Tests {
    /// <summary>
    /// Tracks how much of OfficeIMO.Word.Pdf still uses the Open XML SDK. The converter reads documents through the
    /// Word model and its <c>OfficeXmlNode</c> handles; the count may only go down.
    /// </summary>
    public class WordPdfOpenXmlSdkRatchet {
        private const int Baseline = 0;
        private readonly ITestOutputHelper _output;

        public WordPdfOpenXmlSdkRatchet(ITestOutputHelper output) => _output = output;

        [Fact]
        public void FilesThatUseTheSdkOnlyGoDown() {
            string projectDirectory = Path.Combine(FindRepositoryRoot(), "OfficeIMO.Word.Pdf");
            IReadOnlyList<string> files = OpenXmlSdkReferenceCounter.FindFiles(
                typeof(OfficeIMO.Word.Pdf.WordPdfConverterExtensions).Assembly.Location, projectDirectory, "OpenXmlSdk");

            _output.WriteLine($"OfficeIMO.Word.Pdf files that use the Open XML SDK: {files.Count} (baseline {Baseline})");
            foreach (string file in files) _output.WriteLine(file);

            Assert.True(files.Count <= Baseline,
                $"{files.Count} files in OfficeIMO.Word.Pdf use the Open XML SDK, more than the baseline of {Baseline}. Read through the Word model and OfficeXmlNode instead.");
            Assert.True(files.Count == Baseline,
                $"Only {files.Count} files in OfficeIMO.Word.Pdf use the Open XML SDK now. Lower the baseline in {nameof(WordPdfOpenXmlSdkRatchet)} to {files.Count}.");
        }

        private static string FindRepositoryRoot() {
            string? directory = AppContext.BaseDirectory;
            while (directory != null) {
                if (File.Exists(Path.Combine(directory, "OfficeIMO.Word.Pdf", "OfficeIMO.Word.Pdf.csproj"))) return directory;
                directory = Path.GetDirectoryName(directory);
            }
            throw new InvalidOperationException("The repository root was not found above the test output.");
        }
    }
}
#endif
