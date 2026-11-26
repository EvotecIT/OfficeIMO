#if NET
using System.Xml.Linq;
using OfficeIMO.TestSupport.OpenXml;
using Xunit;
using Xunit.Abstractions;

namespace OfficeIMO.Tests;

/// <summary>
/// Tracks how much of OfficeIMO.PowerPoint still uses the Open XML SDK. Every SDK use belongs in the bridge folder
/// (<c>OfficeIMO.PowerPoint/OpenXmlSdk</c>), which a later version moves into its own package; this counts the source
/// files outside it whose compiled code uses SDK types. The count may only go down: lower <see cref="Baseline"/> in the
/// change that removes files from the list. The engine is off the SDK when it reaches zero.
/// </summary>
public class PowerPointSdkReferenceRatchetTests {
    private const int Baseline = 294;
    private const int PackagingBaseline = 1;
    private readonly ITestOutputHelper _output;

    public PowerPointSdkReferenceRatchetTests(ITestOutputHelper output) => _output = output;

    [Fact]
    public void FilesOutsideTheBridgeThatUseTheSdkOnlyGoDown() {
        string projectDirectory = Path.Combine(FindRepositoryRoot(), "OfficeIMO.PowerPoint");
        IReadOnlyList<string> files = OpenXmlSdkReferenceCounter.FindFiles(
            typeof(OfficeIMO.PowerPoint.PowerPointDocument).Assembly.Location, projectDirectory, "OpenXmlSdk");

        _output.WriteLine($"OfficeIMO.PowerPoint files outside OpenXmlSdk/ that use the Open XML SDK: {files.Count} (baseline {Baseline})");
        foreach (string file in files) _output.WriteLine(file);

        Assert.True(files.Count <= Baseline,
            $"{files.Count} files outside OfficeIMO.PowerPoint/OpenXmlSdk use the Open XML SDK, more than the baseline of {Baseline}. Keep new SDK use in the bridge.");
        Assert.True(files.Count == Baseline,
            $"Only {files.Count} files outside OfficeIMO.PowerPoint/OpenXmlSdk use the Open XML SDK now. Lower the baseline in {nameof(PowerPointSdkReferenceRatchetTests)} to {files.Count}.");
    }

    [Fact]
    public void FilesOutsideTheBridgeThatUseSystemIOPackagingOnlyGoDown() {
        string projectDirectory = Path.Combine(FindRepositoryRoot(), "OfficeIMO.PowerPoint");
        IReadOnlyList<string> files = OpenXmlSdkReferenceCounter.FindFiles(
            typeof(OfficeIMO.PowerPoint.PowerPointDocument).Assembly.Location, projectDirectory, "OpenXmlSdk",
            OpenXmlSdkReferenceCounter.Dependency.SystemIOPackaging);

        _output.WriteLine($"OfficeIMO.PowerPoint files outside OpenXmlSdk/ that use System.IO.Packaging: {files.Count} (baseline {PackagingBaseline})");
        foreach (string file in files) _output.WriteLine(file);

        Assert.True(files.Count <= PackagingBaseline,
            $"{files.Count} files outside OfficeIMO.PowerPoint/OpenXmlSdk use System.IO.Packaging, more than the baseline of {PackagingBaseline}. Use OfficeIMO's own package layer, or keep it in the bridge.");
        Assert.True(files.Count == PackagingBaseline,
            $"Only {files.Count} files outside OfficeIMO.PowerPoint/OpenXmlSdk use System.IO.Packaging now. Lower {nameof(PackagingBaseline)} in {nameof(PowerPointSdkReferenceRatchetTests)} to {files.Count}.");
    }

    [Fact]
    public void TheProjectDoesNotImportTheSdkNamespaceGlobally() {
        // A global using would let any file use SDK types without naming them.
        XDocument project = XDocument.Load(Path.Combine(FindRepositoryRoot(), "OfficeIMO.PowerPoint", "OfficeIMO.PowerPoint.csproj"));
        Assert.DoesNotContain(project.Descendants().Where(element => element.Name.LocalName == "Using"),
            element => ((string?)element.Attribute("Include") ?? string.Empty).StartsWith("DocumentFormat.OpenXml", StringComparison.Ordinal));
    }

    private static string FindRepositoryRoot() {
        string? directory = AppContext.BaseDirectory;
        while (directory != null) {
            if (File.Exists(Path.Combine(directory, "OfficeIMO.PowerPoint", "OfficeIMO.PowerPoint.csproj"))) return directory;
            directory = Path.GetDirectoryName(directory);
        }
        throw new InvalidOperationException("The repository root was not found above the test output.");
    }
}
#endif
