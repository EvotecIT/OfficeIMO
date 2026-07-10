namespace OfficeIMO.Latex.Tests;

public sealed class LatexDependencyGuardrailTests {
    [Theory]
    [InlineData("OfficeIMO.Latex/OfficeIMO.Latex.csproj")]
    [InlineData("OfficeIMO.Latex.Markdown/OfficeIMO.Latex.Markdown.csproj")]
    [InlineData("OfficeIMO.Reader.Latex/OfficeIMO.Reader.Latex.csproj")]
    public void ProductionProjects_HaveNoNuGetDependencies(string relativeProjectPath) {
        string root = GetRepositoryRoot();
        XDocument project = XDocument.Load(Path.Combine(root, relativeProjectPath.Replace('/', Path.DirectorySeparatorChar)));

        Assert.DoesNotContain(project.Descendants(), element => element.Name.LocalName == "PackageReference");
    }

    [Theory]
    [InlineData("OfficeIMO.Latex")]
    [InlineData("OfficeIMO.Latex.Markdown")]
    [InlineData("OfficeIMO.Reader.Latex")]
    public void ProductionProjects_DoNotUseRegexOrExternalProcesses(string projectFolderName) {
        string folder = Path.Combine(GetRepositoryRoot(), projectFolderName);
        string[] files = Directory.EnumerateFiles(folder, "*.cs", SearchOption.AllDirectories)
            .Where(static file => !file.Contains(Path.DirectorySeparatorChar + "obj" + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))
            .Where(static file => !file.Contains(Path.DirectorySeparatorChar + "bin" + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))
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
        throw new DirectoryNotFoundException("Unable to locate repository root.");
    }
}
