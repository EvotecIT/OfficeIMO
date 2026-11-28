#if NET
using OfficeIMO.TestSupport.OpenXml;
using Xunit;
using Xunit.Abstractions;

namespace OfficeIMO.Tests {
    /// <summary>
    /// Tracks how much of OfficeIMO.Word.Html still uses the Open XML SDK. The converter reads documents through the
    /// Word model and its <c>OfficeXmlNode</c> handles; the count may only go down.
    /// </summary>
    public class WordHtmlOpenXmlSdkRatchet {
        private const int Baseline = 0;
        private readonly ITestOutputHelper _output;

        public WordHtmlOpenXmlSdkRatchet(ITestOutputHelper output) => _output = output;

        [Fact]
        public void FilesThatUseTheSdkOnlyGoDown() {
            string projectDirectory = Path.Combine(FindRepositoryRoot(), "OfficeIMO.Word.Html");
            IReadOnlyList<string> files = OpenXmlSdkReferenceCounter.FindFiles(
                typeof(OfficeIMO.Word.Html.WordHtmlConverterExtensions).Assembly.Location, projectDirectory, "OpenXmlSdk");

            _output.WriteLine($"OfficeIMO.Word.Html files that use the Open XML SDK: {files.Count} (baseline {Baseline})");
            foreach (string file in files) _output.WriteLine(file);

            Assert.True(files.Count <= Baseline,
                $"{files.Count} files in OfficeIMO.Word.Html use the Open XML SDK, more than the baseline of {Baseline}. Read through the Word model and OfficeXmlNode instead.");
            Assert.True(files.Count == Baseline,
                $"Only {files.Count} files in OfficeIMO.Word.Html use the Open XML SDK now. Lower the baseline in {nameof(WordHtmlOpenXmlSdkRatchet)} to {files.Count}.");
        }

        private static string FindRepositoryRoot() {
            string? directory = AppContext.BaseDirectory;
            while (directory != null) {
                if (File.Exists(Path.Combine(directory, "OfficeIMO.Word.Html", "OfficeIMO.Word.Html.csproj"))) return directory;
                directory = Path.GetDirectoryName(directory);
            }
            throw new InvalidOperationException("The repository root was not found above the test output.");
        }
    }
}
#endif
