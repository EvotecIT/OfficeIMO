using System.Xml.Linq;

namespace OfficeIMO.OneNote.Tests;

public sealed class PackagingContractTests {
    [Fact]
    public void CorePackageCarriesRequiredThirdPartyNotice() {
        string projectPath = Path.Combine(GetRepositoryRoot(), "OfficeIMO.OneNote", "OfficeIMO.OneNote.csproj");
        XDocument project = XDocument.Load(projectPath);
        XNamespace ns = project.Root?.Name.Namespace ?? XNamespace.None;

        XElement notice = Assert.Single(
            project.Descendants(ns + "None"),
            element => string.Equals(
                (string?)element.Attribute("Include"),
                "THIRD-PARTY-NOTICES.md",
                StringComparison.OrdinalIgnoreCase));

        Assert.Equal("True", notice.Element(ns + "Pack")?.Value, ignoreCase: true);
        Assert.Equal("THIRD-PARTY-NOTICES.md", notice.Element(ns + "PackagePath")?.Value);
    }

    private static string GetRepositoryRoot() {
        DirectoryInfo? directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory != null) {
            if (File.Exists(Path.Combine(directory.FullName, "OfficeIMO.sln"))) return directory.FullName;
            directory = directory.Parent;
        }

        throw new DirectoryNotFoundException("Unable to locate the OfficeIMO repository root.");
    }
}
