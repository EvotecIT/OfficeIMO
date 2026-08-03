using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class ReleasePackagingGuardrails {
    [Fact]
    public void CodexMarketplace_LocalPluginSourcesResolveFromRepositoryRoot() {
        string repositoryRoot = GetRepositoryRoot();
        string marketplacePath = Path.Combine(repositoryRoot, ".agents", "plugins", "marketplace.json");
        using JsonDocument marketplace = JsonDocument.Parse(File.ReadAllText(marketplacePath));

        Assert.All(marketplace.RootElement.GetProperty("plugins").EnumerateArray(), plugin => {
            JsonElement source = plugin.GetProperty("source");
            if (!string.Equals(source.GetProperty("source").GetString(), "local", StringComparison.Ordinal)) {
                return;
            }

            string relativePath = source.GetProperty("path").GetString()
                ?? throw new InvalidDataException("Local plugin sources must declare a path.");
            string resolvedPath = Path.GetFullPath(Path.Combine(repositoryRoot, relativePath));

            Assert.True(
                Directory.Exists(resolvedPath),
                $"Local plugin source '{relativePath}' does not resolve from the repository root.");
        });
    }

    [Fact]
    public void CodexPlugin_McpConfigurationUsesSupportedServerSchema() {
        string repositoryRoot = GetRepositoryRoot();
        string coordinatedVersion = ReadCoordinatedReleaseVersion(repositoryRoot);
        string mcpPath = Path.Combine(
            repositoryRoot,
            ".agents",
            "plugins",
            "officeimo-document-tools",
            ".mcp.json");
        using JsonDocument mcp = JsonDocument.Parse(File.ReadAllText(mcpPath));
        JsonElement officeImo = mcp.RootElement
            .GetProperty("mcpServers")
            .GetProperty("officeimo");

        Assert.Equal("stdio", officeImo.GetProperty("type").GetString());
        Assert.Equal("dotnet", officeImo.GetProperty("command").GetString());
        Assert.Equal(
            ["dnx", $"OfficeIMO.Tool@{coordinatedVersion}", "mcp", "serve", "--stdio"],
            officeImo.GetProperty("args").EnumerateArray()
                .Select(static value => value.GetString() ?? string.Empty)
                .ToArray());
        Assert.False(
            officeImo.TryGetProperty("tools", out JsonElement tools) &&
            tools.ValueKind == JsonValueKind.Array,
            "Codex MCP tool configuration is a map when present, not an array allowlist.");
    }

    [Fact]
    public void PackageLocks_DoNotRetainOlderOfficeIMOReleaseLines() {
        string repositoryRoot = GetRepositoryRoot();
        string coordinatedVersion = ReadCoordinatedReleaseVersion(repositoryRoot);
        string[] lockFiles = Directory
            .EnumerateFiles(repositoryRoot, "packages.lock.json", SearchOption.AllDirectories)
            .Where(static path => !ContainsBuildOutput(path))
            .ToArray();
        Assert.NotEmpty(lockFiles);

        var staleDependencies = new List<string>();
        foreach (string lockFile in lockFiles) {
            string content = File.ReadAllText(lockFile);
            foreach (Match match in Regex.Matches(
                content,
                "\"OfficeIMO\\.[^\"]+\"\\s*:\\s*\"\\[(?<version>\\d+\\.\\d+\\.\\d+),",
                RegexOptions.CultureInvariant | RegexOptions.IgnoreCase)) {
                if (!string.Equals(
                    match.Groups["version"].Value,
                    coordinatedVersion,
                    StringComparison.Ordinal)) {
                    staleDependencies.Add(
                        GetRepositoryRelativePath(repositoryRoot, lockFile)
                        + " -> "
                        + match.Value);
                }
            }
        }

        Assert.Empty(staleDependencies);
    }

    [Fact]
    public void ProjectBuild_IncludesEveryPublishablePackageExactlyOnceAndUsesOneVersion() {
        string repositoryRoot = GetRepositoryRoot();
        string projectBuildPath = Path.Combine(repositoryRoot, "Build", "project.build.json");
        using JsonDocument buildDocument = JsonDocument.Parse(File.ReadAllText(projectBuildPath));

        JsonElement buildRoot = buildDocument.RootElement;
        Dictionary<string, string> expectedVersions = buildRoot
            .GetProperty("ExpectedVersionMap")
            .EnumerateObject()
            .ToDictionary(
                static property => property.Name,
                static property => property.Value.GetString() ?? string.Empty,
                StringComparer.OrdinalIgnoreCase);
        string releaseBand = buildRoot.GetProperty("ExpectedVersion").GetString()
            ?? throw new InvalidDataException("Build/project.build.json must declare ExpectedVersion.");
        Assert.Matches(@"^\d+\.\d+\.X$", releaseBand);
        Assert.NotEmpty(expectedVersions);
        Assert.All(expectedVersions, entry => Assert.Equal(releaseBand, entry.Value));
        HashSet<string> excludedProjects = buildRoot
            .GetProperty("ExcludeProjects")
            .EnumerateArray()
            .Select(static element => element.GetString())
            .Where(static value => !string.IsNullOrWhiteSpace(value))
            .Select(static value => value!)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        Assert.Empty(excludedProjects);

        PackageProject[] packageProjects = Directory
            .EnumerateFiles(repositoryRoot, "*.csproj", SearchOption.AllDirectories)
            .Where(static path => !ContainsBuildOutput(path))
            .Select(ReadPackageProject)
            .Where(static project => project is not null)
            .Select(static project => project!)
            .ToArray();

        string[] duplicatePackageIds = packageProjects
            .GroupBy(static project => project.PackageId, StringComparer.OrdinalIgnoreCase)
            .Where(static group => group.Count() > 1)
            .Select(static group => group.Key)
            .OrderBy(static packageId => packageId, StringComparer.OrdinalIgnoreCase)
            .ToArray();
        Assert.Empty(duplicatePackageIds);

        string[] missingFromBuild = packageProjects
            .Where(project => !expectedVersions.ContainsKey(project.PackageId))
            .Select(static project => project.PackageId)
            .OrderBy(static packageId => packageId, StringComparer.OrdinalIgnoreCase)
            .ToArray();
        Assert.Empty(missingFromBuild);

        string[] staleBuildEntries = expectedVersions.Keys
            .Where(packageId => !packageProjects.Any(project =>
                string.Equals(project.PackageId, packageId, StringComparison.OrdinalIgnoreCase)))
            .OrderBy(static packageId => packageId, StringComparer.OrdinalIgnoreCase)
            .ToArray();
        Assert.Empty(staleBuildEntries);

        PackageProject[] includedProjects = packageProjects
            .Where(project => expectedVersions.ContainsKey(project.PackageId))
            .ToArray();
        foreach (PackageProject project in includedProjects) {
            AssertVersionMatchesReleaseBand(project, expectedVersions[project.PackageId]);
        }

        string[] releaseVersions = includedProjects
            .Select(static project => project.Version)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
        Assert.Single(releaseVersions);
    }

    [Fact]
    public void SolutionReleaseConfiguration_IncludesEveryPublishablePackage() {
        string repositoryRoot = GetRepositoryRoot();
        string solution = File.ReadAllText(Path.Combine(repositoryRoot, "OfficeIMO.sln"));
        PackageProject[] packageProjects = Directory
            .EnumerateFiles(repositoryRoot, "*.csproj", SearchOption.AllDirectories)
            .Where(static path => !ContainsBuildOutput(path))
            .Select(ReadPackageProject)
            .Where(static project => project is not null)
            .Select(static project => project!)
            .ToArray();

        Assert.All(packageProjects, project => {
            Match projectDeclaration = Assert.Single(Regex.Matches(
                solution,
                $@"^Project\(""[^""]+""\) = ""{Regex.Escape(project.ProjectName)}"", ""[^""]+"", ""\{{(?<guid>[A-F0-9-]+)\}}""\r?$",
                RegexOptions.Multiline | RegexOptions.CultureInvariant | RegexOptions.IgnoreCase)
                .Cast<Match>());
            string projectGuid = projectDeclaration.Groups["guid"].Value;

            Assert.Contains(
                $"{{{projectGuid}}}.Release|Any CPU.Build.0 = Release|Any CPU",
                solution,
                StringComparison.OrdinalIgnoreCase);
        });
    }

    private static PackageProject? ReadPackageProject(string projectPath) {
        XDocument document = XDocument.Load(projectPath);
        XNamespace ns = document.Root?.Name.Namespace ?? XNamespace.None;
        string? packageId = document.Descendants(ns + "PackageId").Select(static element => element.Value).FirstOrDefault();
        string? version = document.Descendants(ns + "VersionPrefix").Select(static element => element.Value).FirstOrDefault();
        if (string.IsNullOrWhiteSpace(packageId) || string.IsNullOrWhiteSpace(version)) {
            return null;
        }

        bool isPackable = !document
            .Descendants(ns + "IsPackable")
            .Any(static element => string.Equals(element.Value, "false", StringComparison.OrdinalIgnoreCase));
        bool isPublishable = !document
            .Descendants(ns + "IsPublishable")
            .Any(static element => string.Equals(element.Value, "false", StringComparison.OrdinalIgnoreCase));
        if (!isPackable || !isPublishable) {
            return null;
        }

        string projectName = Path.GetFileNameWithoutExtension(projectPath);
        return new PackageProject(projectName, packageId, version);
    }

    private static string ReadCoordinatedReleaseVersion(string repositoryRoot) {
        string[] versions = Directory
            .EnumerateFiles(repositoryRoot, "*.csproj", SearchOption.AllDirectories)
            .Where(static path => !ContainsBuildOutput(path))
            .Select(ReadPackageProject)
            .Where(static project => project is not null)
            .Select(static project => project!.Version)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
        return Assert.Single(versions);
    }

    private static void AssertVersionMatchesReleaseBand(PackageProject project, string expectedBand) {
        string[] expectedParts = expectedBand.Split('.');
        string[] versionParts = project.Version.Split('.');

        Assert.Equal(3, expectedParts.Length);
        Assert.Equal(3, versionParts.Length);
        Assert.Equal("X", expectedParts[2], ignoreCase: true);
        Assert.Equal(expectedParts[0], versionParts[0]);
        Assert.Equal(expectedParts[1], versionParts[1]);
        Assert.True(
            int.TryParse(versionParts[2], out _),
            $"Package '{project.PackageId}' has invalid patch version '{project.Version}'.");
    }

    private static bool ContainsBuildOutput(string path) =>
        path.IndexOf($"{Path.DirectorySeparatorChar}bin{Path.DirectorySeparatorChar}", StringComparison.OrdinalIgnoreCase) >= 0 ||
        path.IndexOf($"{Path.DirectorySeparatorChar}obj{Path.DirectorySeparatorChar}", StringComparison.OrdinalIgnoreCase) >= 0;

    private static string GetRepositoryRelativePath(string repositoryRoot, string path) {
        string normalizedRoot = Path.GetFullPath(repositoryRoot);
        if (!normalizedRoot.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal)) {
            normalizedRoot += Path.DirectorySeparatorChar;
        }

        var rootUri = new Uri(normalizedRoot, UriKind.Absolute);
        var pathUri = new Uri(Path.GetFullPath(path), UriKind.Absolute);
        string relativePath = Uri.UnescapeDataString(rootUri.MakeRelativeUri(pathUri).ToString());
        Assert.False(
            relativePath == ".." || relativePath.StartsWith("../", StringComparison.Ordinal),
            "Path must stay under repository root: " + path);
        return relativePath.Replace('\\', '/');
    }

    private static string GetRepositoryRoot() {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null) {
            if (File.Exists(Path.Combine(directory.FullName, "OfficeIMO.sln"))) {
                return directory.FullName;
            }
            directory = directory.Parent;
        }

        throw new DirectoryNotFoundException("Could not locate the OfficeIMO repository root.");
    }

    private sealed class PackageProject {
        internal PackageProject(string projectName, string packageId, string version) {
            ProjectName = projectName;
            PackageId = packageId;
            Version = version;
        }

        internal string ProjectName { get; }

        internal string PackageId { get; }

        internal string Version { get; }
    }
}
