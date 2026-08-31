using System.Xml.Linq;

namespace OfficeIMO.Bibliography.Tests.Packaging;

public sealed class BibliographyDependencyGuardrailTests {
    [Fact]
    public void Production_package_has_only_the_intended_dependencies() {
        XDocument project = XDocument.Load(Path.Combine(GetRepositoryRoot(), "OfficeIMO.Bibliography", "OfficeIMO.Bibliography.csproj"));

        string[] projects = project.Descendants().Where(element => element.Name.LocalName == "ProjectReference").Select(element => Path.GetFileNameWithoutExtension(((string?)element.Attribute("Include") ?? string.Empty).Replace('\\', '/'))).ToArray();
        string[] packages = project.Descendants().Where(element => element.Name.LocalName == "PackageReference").Select(element => (string?)element.Attribute("Include") ?? string.Empty).ToArray();

        Assert.Empty(projects);
        Assert.Equal(new[] { "System.Text.Encoding.CodePages", "System.Text.Json" }, packages);
        Assert.DoesNotContain(project.Descendants(), element => ((string?)element.Attribute("Include"))?.IndexOf("OfficeIMO.Word", StringComparison.OrdinalIgnoreCase) >= 0);
        Assert.DoesNotContain(project.Descendants(), element => ((string?)element.Attribute("Include"))?.IndexOf("DocumentFormat.OpenXml", StringComparison.OrdinalIgnoreCase) >= 0);

        XElement jsonReference = Assert.Single(project.Descendants(), element => element.Name.LocalName == "PackageReference" && (string?)element.Attribute("Include") == "System.Text.Json");
        string? jsonCondition = (string?)jsonReference.Parent?.Attribute("Condition");
        Assert.Contains("netstandard2.0", jsonCondition, StringComparison.Ordinal);
        Assert.Contains("net472", jsonCondition, StringComparison.Ordinal);
        Assert.DoesNotContain("net8.0", jsonCondition, StringComparison.Ordinal);
        Assert.DoesNotContain("net10.0", jsonCondition, StringComparison.Ordinal);
    }

    private static string GetRepositoryRoot() {
        DirectoryInfo? directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory != null) {
            if (File.Exists(Path.Combine(directory.FullName, "OfficeIMO.sln"))) return directory.FullName;
            directory = directory.Parent;
        }
        throw new DirectoryNotFoundException("Unable to locate repository root.");
    }
}
