#if NET
using OfficeIMO.TestSupport.OpenXml;
using Xunit;
using Xunit.Abstractions;

namespace OfficeIMO.Tests {
    /// <summary>
    /// Tracks how much of OfficeIMO.Word.Markdown still uses the Open XML SDK. The converter reads documents through the
    /// Word model and its <c>OfficeXmlNode</c> handles; the count may only go down.
    /// </summary>
    public class WordMarkdownOpenXmlSdkRatchet {
        private const int Baseline = 16;
        private readonly ITestOutputHelper _output;

        public WordMarkdownOpenXmlSdkRatchet(ITestOutputHelper output) => _output = output;

        [Fact]
        public void FilesThatUseTheSdkOnlyGoDown() {
            string projectDirectory = Path.Combine(FindRepositoryRoot(), "OfficeIMO.Word.Markdown");
            IReadOnlyList<string> files = OpenXmlSdkReferenceCounter.FindFiles(
                typeof(OfficeIMO.Word.Markdown.WordMarkdownConverterExtensions).Assembly.Location, projectDirectory, "OpenXmlSdk");

            _output.WriteLine($"OfficeIMO.Word.Markdown files that use the Open XML SDK: {files.Count} (baseline {Baseline})");
            foreach (string file in files) _output.WriteLine(file);

            Assert.True(files.Count <= Baseline,
                $"{files.Count} files in OfficeIMO.Word.Markdown use the Open XML SDK, more than the baseline of {Baseline}. Read through the Word model and OfficeXmlNode instead.");
            Assert.True(files.Count == Baseline,
                $"Only {files.Count} files in OfficeIMO.Word.Markdown use the Open XML SDK now. Lower the baseline in {nameof(WordMarkdownOpenXmlSdkRatchet)} to {files.Count}.");
        }

        private static string FindRepositoryRoot() {
            string? directory = AppContext.BaseDirectory;
            while (directory != null) {
                if (File.Exists(Path.Combine(directory, "OfficeIMO.Word.Markdown", "OfficeIMO.Word.Markdown.csproj"))) return directory;
                directory = Path.GetDirectoryName(directory);
            }
            throw new InvalidOperationException("The repository root was not found above the test output.");
        }
    }
}
#endif
