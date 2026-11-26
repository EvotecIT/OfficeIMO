using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>
/// The PowerPoint engine is moving off the Open XML SDK. Its remaining SDK use belongs in the bridge folder
/// <c>OfficeIMO.PowerPoint/OpenXmlSdk</c>, which a later version moves into its own package; every other source file
/// compiled into the engine, linked shared sources included, should stop naming SDK types. The count of files that still
/// do may only go down. When a change lowers it, lower <see cref="MaximumFilesReferencingSdk"/> to the new count.
/// </summary>
public class PowerPointSdkReferenceRatchetTests {
    private const int MaximumFilesReferencingSdk = 360;
    private const string BridgeFolder = "OpenXmlSdk";

    [Fact]
    public void FilesOutsideTheBridgeReferencingTheSdkDoNotGrow() {
        string? project = FindProjectFolder();
        // Test runs from a published build have no sources; the check runs wherever the repository is checked out.
        if (project == null) return;
        List<string> files = EngineSources(project)
            .Where(file => !IsInBridge(project, file) && ReferencesSdk(file))
            .Select(file => Relative(Path.GetDirectoryName(project)!, file).Replace('\\', '/'))
            .OrderBy(file => file, StringComparer.Ordinal)
            .ToList();
        Assert.True(files.Count <= MaximumFilesReferencingSdk,
            files.Count + " files outside the bridge reference the Open XML SDK, more than the " + MaximumFilesReferencingSdk
            + " allowed; move the new SDK use into " + BridgeFolder + " or the engine's own model.");
    }

    [Fact]
    public void TheProjectNamesTheSdkNamespaceOnlyInSourceFiles() {
        string? project = FindProjectFolder();
        if (project == null) return;
        // A global using would let any file use SDK types without naming them, which the count above can't see.
        XDocument csproj = XDocument.Load(Path.Combine(project, "OfficeIMO.PowerPoint.csproj"));
        Assert.DoesNotContain(csproj.Descendants().Where(element => element.Name.LocalName == "Using"),
            element => ((string?)element.Attribute("Include") ?? string.Empty).StartsWith("DocumentFormat.OpenXml", StringComparison.Ordinal));
    }

    private static bool ReferencesSdk(string file) => File.ReadAllText(file).IndexOf("DocumentFormat.OpenXml", StringComparison.Ordinal) >= 0;

    private static string Relative(string folder, string file) =>
        file.StartsWith(folder + Path.DirectorySeparatorChar, StringComparison.Ordinal) ? file.Substring(folder.Length + 1) : file;

    private static bool IsInBridge(string project, string file) =>
        file.StartsWith(Path.Combine(project, BridgeFolder) + Path.DirectorySeparatorChar, StringComparison.Ordinal);

    // The engine's own sources and the shared sources its project links in.
    private static IEnumerable<string> EngineSources(string project) {
        foreach (string file in Directory.EnumerateFiles(project, "*.cs", SearchOption.AllDirectories)) {
            string relative = Relative(project, file);
            if (relative.StartsWith("obj" + Path.DirectorySeparatorChar, StringComparison.Ordinal)
                || relative.StartsWith("bin" + Path.DirectorySeparatorChar, StringComparison.Ordinal)) continue;
            yield return file;
        }
        XDocument csproj = XDocument.Load(Path.Combine(project, "OfficeIMO.PowerPoint.csproj"));
        foreach (string include in csproj.Descendants().Where(element => element.Name.LocalName == "Compile")
                     .Select(element => (string?)element.Attribute("Include")).OfType<string>()
                     .Where(include => include.StartsWith("..", StringComparison.Ordinal))) {
            string pattern = include.Replace('\\', Path.DirectorySeparatorChar);
            string folder = Path.GetFullPath(Path.Combine(project, Path.GetDirectoryName(pattern)!));
            if (!Directory.Exists(folder)) continue;
            foreach (string file in Directory.EnumerateFiles(folder, Path.GetFileName(pattern))) yield return file;
        }
    }

    private static string? FindProjectFolder() {
        for (DirectoryInfo? directory = new(AppContext.BaseDirectory); directory != null; directory = directory.Parent) {
            string candidate = Path.Combine(directory.FullName, "OfficeIMO.PowerPoint");
            if (File.Exists(Path.Combine(candidate, "OfficeIMO.PowerPoint.csproj"))) return candidate;
        }
        return null;
    }
}
