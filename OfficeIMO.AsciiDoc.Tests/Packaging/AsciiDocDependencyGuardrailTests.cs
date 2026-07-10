namespace OfficeIMO.AsciiDoc.Tests;

public sealed class AsciiDocDependencyGuardrailTests {
    [Theory]
    [InlineData("OfficeIMO.AsciiDoc/OfficeIMO.AsciiDoc.csproj")]
    [InlineData("OfficeIMO.AsciiDoc.Markdown/OfficeIMO.AsciiDoc.Markdown.csproj")]
    [InlineData("OfficeIMO.Reader.AsciiDoc/OfficeIMO.Reader.AsciiDoc.csproj")]
    public void ProductionProjects_DoNotAddNuGetDependencies(string relativeProjectPath) {
        string path = Path.Combine(GetRepositoryRoot(), relativeProjectPath.Replace('/', Path.DirectorySeparatorChar));
        XDocument project = XDocument.Load(path);

        Assert.DoesNotContain(project.Descendants(), element => element.Name.LocalName == "PackageReference");
    }

    [Theory]
    [InlineData("OfficeIMO.AsciiDoc")]
    [InlineData("OfficeIMO.AsciiDoc.Markdown")]
    [InlineData("OfficeIMO.Reader.AsciiDoc")]
    public void ProductionProjects_DoNotUseRegexOrExternalProcesses(string projectFolderName) {
        string folder = Path.Combine(GetRepositoryRoot(), projectFolderName);
        string[] files = Directory.EnumerateFiles(folder, "*.cs", SearchOption.AllDirectories)
            .Where(static file => file.IndexOf(Path.DirectorySeparatorChar + "bin" + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase) < 0)
            .Where(static file => file.IndexOf(Path.DirectorySeparatorChar + "obj" + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase) < 0)
            .ToArray();

        Assert.NotEmpty(files);
        foreach (string file in files) {
            string source = File.ReadAllText(file);
            Assert.DoesNotContain("System.Text.RegularExpressions", source, StringComparison.Ordinal);
            Assert.DoesNotContain("System.Diagnostics.Process", source, StringComparison.Ordinal);
            Assert.DoesNotContain("new Process", source, StringComparison.Ordinal);
        }
    }

    private static string GetRepositoryRoot() {
        DirectoryInfo? directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory != null) {
            if (File.Exists(Path.Combine(directory.FullName, "OfficeIMO.sln"))) return directory.FullName;
            directory = directory.Parent;
        }
        throw new DirectoryNotFoundException("Unable to locate OfficeIMO repository root.");
    }
}
