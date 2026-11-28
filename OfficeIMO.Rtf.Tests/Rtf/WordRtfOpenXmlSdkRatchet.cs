#if NET
using OfficeIMO.TestSupport.OpenXml;
using Xunit;
using Xunit.Abstractions;

namespace OfficeIMO.Rtf.Tests;

/// <summary>
/// Tracks how much of OfficeIMO.Word.Rtf still uses the Open XML SDK. The converter reads and builds documents through
/// the Word model and its <c>OfficeXmlNode</c> handles; the count may only go down.
/// </summary>
public class WordRtfOpenXmlSdkRatchet {
    private const int Baseline = 0;
    private readonly ITestOutputHelper _output;

    public WordRtfOpenXmlSdkRatchet(ITestOutputHelper output) => _output = output;

    [Fact]
    public void FilesThatUseTheSdkOnlyGoDown() {
        string projectDirectory = Path.Combine(FindRepositoryRoot(), "OfficeIMO.Word.Rtf");
        IReadOnlyList<string> files = OpenXmlSdkReferenceCounter.FindFiles(
            typeof(OfficeIMO.Word.Rtf.WordRtfConverterExtensions).Assembly.Location, projectDirectory, "OpenXmlSdk");

        _output.WriteLine($"OfficeIMO.Word.Rtf files that use the Open XML SDK: {files.Count} (baseline {Baseline})");
        foreach (string file in files) _output.WriteLine(file);

        Assert.True(files.Count <= Baseline,
            $"{files.Count} files in OfficeIMO.Word.Rtf use the Open XML SDK, more than the baseline of {Baseline}. Read through the Word model and OfficeXmlNode instead.");
        Assert.True(files.Count == Baseline,
            $"Only {files.Count} files in OfficeIMO.Word.Rtf use the Open XML SDK now. Lower the baseline in {nameof(WordRtfOpenXmlSdkRatchet)} to {files.Count}.");
    }

    private static string FindRepositoryRoot() {
        string? directory = AppContext.BaseDirectory;
        while (directory != null) {
            if (File.Exists(Path.Combine(directory, "OfficeIMO.Word.Rtf", "OfficeIMO.Word.Rtf.csproj"))) return directory;
            directory = Path.GetDirectoryName(directory);
        }
        throw new InvalidOperationException("The repository root was not found above the test output.");
    }
}
#endif
