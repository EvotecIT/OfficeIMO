using System.Xml.Linq;

namespace OfficeIMO.Bibliography.Tests.Packaging;

public sealed class BibliographyDependencyGuardrailTests {
    [Fact]
    public void Production_package_has_only_the_intended_dependencies() {
        XDocument project = XDocument.Load(Path.Combine(GetRepositoryRoot(), "OfficeIMO.Bibliography", "OfficeIMO.Bibliography.csproj"));

        string[] projects = project.Descendants().Where(element => element.Name.LocalName == "ProjectReference").Select(element => Path.GetFileNameWithoutExtension(((string?)element.Attribute("Include") ?? string.Empty).Replace('\\', '/'))).ToArray();
        string[] packages = project.Descendants().Where(element => element.Name.LocalName == "PackageReference").Select(element => (string?)element.Attribute("Include") ?? string.Empty).ToArray();

        Assert.Empty(projects);
        Assert.Equal(new[] { "System.Text.Json" }, packages);
        Assert.DoesNotContain(project.Descendants(), element => ((string?)element.Attribute("Include"))?.IndexOf("OfficeIMO.Word", StringComparison.OrdinalIgnoreCase) >= 0);
        Assert.DoesNotContain(project.Descendants(), element => ((string?)element.Attribute("Include"))?.IndexOf("DocumentFormat.OpenXml", StringComparison.OrdinalIgnoreCase) >= 0);
    }

    [Fact]
    public void Production_source_does_not_execute_processes_or_use_network_clients() {
        string folder = Path.Combine(GetRepositoryRoot(), "OfficeIMO.Bibliography");
        string[] files = Directory.EnumerateFiles(folder, "*.cs", SearchOption.AllDirectories)
            .Where(static file => file.IndexOf(Path.DirectorySeparatorChar + "obj" + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase) < 0)
            .Where(static file => file.IndexOf(Path.DirectorySeparatorChar + "bin" + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase) < 0)
            .ToArray();

        Assert.NotEmpty(files);
        foreach (string file in files) {
            string source = File.ReadAllText(file);
            Assert.DoesNotContain("System.Diagnostics.Process", source, StringComparison.Ordinal);
            Assert.DoesNotContain("HttpClient", source, StringComparison.Ordinal);
            Assert.DoesNotContain("WebClient", source, StringComparison.Ordinal);
        }
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
