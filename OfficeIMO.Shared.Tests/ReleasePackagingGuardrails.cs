using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class ReleasePackagingGuardrails {
    private const string CurrentPublishedPackageVersion = "3.0.3";

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
    public void ReadmeInventory_MatchesReleaseMapAndLinkedProjectCatalog() {
        string repositoryRoot = GetRepositoryRoot();
        string readme = File.ReadAllText(Path.Combine(repositoryRoot, "README.md"));
        using JsonDocument buildDocument = JsonDocument.Parse(
            File.ReadAllText(Path.Combine(repositoryRoot, "Build", "project.build.json")));
        HashSet<string> releasePackageIds = buildDocument.RootElement
            .GetProperty("ExpectedVersionMap")
            .EnumerateObject()
            .Select(static property => property.Name)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        int releasePackageCount = releasePackageIds.Count;

        MatchCollection projectHeadings = Regex.Matches(
            readme,
            @"^#### \[(?<name>OfficeIMO\.[^\]]+)\]\((?<path>[^)]+)\)\r?$",
            RegexOptions.Multiline | RegexOptions.CultureInvariant);
        string[] duplicateNames = projectHeadings
            .Cast<Match>()
            .Select(static match => match.Groups["name"].Value)
            .GroupBy(static name => name, StringComparer.OrdinalIgnoreCase)
            .Where(static group => group.Count() > 1)
            .Select(static group => group.Key)
            .ToArray();

        Assert.Empty(duplicateNames);
        Assert.All(projectHeadings.Cast<Match>(), match =>
            Assert.True(
                File.Exists(Path.Combine(repositoryRoot, match.Groups["path"].Value)),
                "README project link is missing: " + match.Value));
        Assert.Equal(26, CountProjectHeadings(readme, "Native formats and shared foundations"));
        Assert.Equal(31, CountProjectHeadings(readme, "Conversion and cloud bridges"));
        Assert.Equal(27, CountProjectHeadings(readme, "Unified Reader family"));
        Assert.Equal(10, CountProjectHeadings(readme, "Markdown rendering and OfficeIMO Markup"));
        Assert.Equal(94, projectHeadings.Count);

        Assert.Contains($"| Coordinated `3.1.x` source packages | {releasePackageCount} |", readme, StringComparison.Ordinal);
        Assert.Contains($"| Documented package, tool, and example projects below | {projectHeadings.Count} |", readme, StringComparison.Ordinal);
        Assert.Contains("| Native format, foundation, and shared-service packages | 26 |", readme, StringComparison.Ordinal);
        Assert.Contains("| Conversion and cloud bridge packages | 31 |", readme, StringComparison.Ordinal);
        Assert.Contains("| Unified Reader packages | 27 |", readme, StringComparison.Ordinal);
        Assert.Contains("| Markdown renderer and OfficeIMO Markup surfaces | 10 |", readme, StringComparison.Ordinal);
        Assert.Contains("The current source line is `3.1.x`; the latest NuGet release is `3.0.3`", readme, StringComparison.Ordinal);
        Assert.Contains("Keep OfficeIMO package references in one application on the same published version", readme, StringComparison.Ordinal);
        AssertDotNetInstallCommands(
            readme,
            releasePackageIds,
            expectedCount: 16,
            expectedVersion: CurrentPublishedPackageVersion);
        Assert.DoesNotContain(
            "dotnet tool install --global OfficeIMO.Tool --version 3.0.3",
            readme,
            StringComparison.Ordinal);
        Assert.Contains(
            "dotnet run --project OfficeIMO.Tool/OfficeIMO.Tool.csproj --framework net8.0 -- <command>",
            readme,
            StringComparison.Ordinal);
        Assert.DoesNotContain("OfficeIMO.Tool/OfficeIMO.Tool.csproj -- ", readme, StringComparison.Ordinal);

        string toolReadme = File.ReadAllText(Path.Combine(repositoryRoot, "OfficeIMO.Tool", "README.md"));
        Assert.Contains("This README describes the `3.1.x` source-tree surface", toolReadme, StringComparison.Ordinal);
        Assert.Contains(
            "dotnet run --project OfficeIMO.Tool/OfficeIMO.Tool.csproj --framework net8.0 -- mcp serve --stdio",
            toolReadme,
            StringComparison.Ordinal);
        Assert.DoesNotContain("OfficeIMO.Tool@3.0.3", toolReadme, StringComparison.Ordinal);
        Assert.DoesNotContain("OfficeIMO.Tool/OfficeIMO.Tool.csproj -- ", toolReadme, StringComparison.Ordinal);
        Assert.DoesNotContain("\nofficeimo ", toolReadme, StringComparison.Ordinal);
        Assert.DoesNotContain("= officeimo ", toolReadme, StringComparison.Ordinal);
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

    [Fact]
    public void PublicReleaseDocs_UseCurrentPackageVersionsAndDocumentationOwners() {
        string repositoryRoot = GetRepositoryRoot();
        string installation = File.ReadAllText(Path.Combine(
            repositoryRoot,
            "Website",
            "content",
            "docs",
            "getting-started",
            "installation",
            "index.md"));
        Dictionary<string, PackageProject> packageProjects = Directory
            .EnumerateFiles(repositoryRoot, "*.csproj", SearchOption.AllDirectories)
            .Where(static path => !ContainsBuildOutput(path))
            .Select(ReadPackageProject)
            .Where(static project => project is not null)
            .Select(static project => project!)
            .ToDictionary(static project => project.PackageId, StringComparer.OrdinalIgnoreCase);
        MatchCollection documentedPackages = Regex.Matches(
            installation,
            @"<PackageReference Include=""(?<id>OfficeIMO\.[^""]+)"" Version=""(?<version>[^""]+)"" />",
            RegexOptions.CultureInvariant);

        Assert.Equal(9, documentedPackages.Count);
        Assert.All(documentedPackages.Cast<Match>(), match => {
            string packageId = match.Groups["id"].Value;
            Assert.True(
                packageProjects.TryGetValue(packageId, out PackageProject? project),
                "Installation guide references an unknown package: " + packageId);
            Assert.Equal(CurrentPublishedPackageVersion, match.Groups["version"].Value);
        });
        Assert.Contains("The current NuGet package line is `3.0.3`", installation, StringComparison.Ordinal);
        Assert.Contains("keep coordinated OfficeIMO package references on the same version", installation, StringComparison.Ordinal);
        HashSet<string> releasePackageIds = packageProjects.Keys.ToHashSet(StringComparer.OrdinalIgnoreCase);
        string[] installationCommandIds = AssertDotNetInstallCommands(
            installation,
            releasePackageIds,
            expectedCount: 9,
            expectedVersion: CurrentPublishedPackageVersion);
        string[] packageReferenceIds = documentedPackages
            .Cast<Match>()
            .Select(static match => match.Groups["id"].Value)
            .OrderBy(static packageId => packageId, StringComparer.OrdinalIgnoreCase)
            .ToArray();
        Assert.Equal(
            packageReferenceIds,
            installationCommandIds.OrderBy(static packageId => packageId, StringComparer.OrdinalIgnoreCase).ToArray());
        AssertPackageManagerInstallCommands(
            installation,
            releasePackageIds,
            expectedCount: 4,
            expectedVersion: CurrentPublishedPackageVersion);

        string openXmlVersion = ReadPackageReferenceVersion(
            repositoryRoot,
            "OfficeIMO.Word/OfficeIMO.Word.csproj",
            "DocumentFormat.OpenXml");
        string angleSharpVersion = ReadPackageReferenceVersion(
            repositoryRoot,
            "OfficeIMO.Html/OfficeIMO.Html.csproj",
            "AngleSharp");
        string angleSharpCssVersion = ReadPackageReferenceVersion(
            repositoryRoot,
            "OfficeIMO.Html/OfficeIMO.Html.csproj",
            "AngleSharp.Css");
        string bouncyCastleVersion = ReadPackageReferenceVersion(
            repositoryRoot,
            "OfficeIMO.Security/OfficeIMO.Security.csproj",
            "BouncyCastle.Cryptography");
        string systemBuffersVersion = ReadPackageReferenceVersion(
            repositoryRoot,
            "OfficeIMO.CSV/OfficeIMO.CSV.csproj",
            "System.Buffers");
        string wordAsyncInterfacesVersion = ReadPackageReferenceVersion(
            repositoryRoot,
            "OfficeIMO.Word/OfficeIMO.Word.csproj",
            "Microsoft.Bcl.AsyncInterfaces");
        string excelAsyncInterfacesVersion = ReadPackageReferenceVersion(
            repositoryRoot,
            "OfficeIMO.Excel/OfficeIMO.Excel.csproj",
            "Microsoft.Bcl.AsyncInterfaces");
        string systemTextJsonVersion = ReadPackageReferenceVersion(
            repositoryRoot,
            "OfficeIMO.Excel/OfficeIMO.Excel.csproj",
            "System.Text.Json");

        Assert.Contains($"DocumentFormat.OpenXml** (`{openXmlVersion}`)", installation, StringComparison.Ordinal);
        Assert.Contains($"AngleSharp** (`{angleSharpVersion}`)", installation, StringComparison.Ordinal);
        Assert.Contains($"AngleSharp.Css** (`{angleSharpCssVersion}`)", installation, StringComparison.Ordinal);
        Assert.Contains($"BouncyCastle.Cryptography** (`{bouncyCastleVersion}`)", installation, StringComparison.Ordinal);
        Assert.Contains($"System.Buffers** (`{systemBuffersVersion}`)", installation, StringComparison.Ordinal);
        Assert.Equal(wordAsyncInterfacesVersion, excelAsyncInterfacesVersion);
        Assert.Contains(
            $"Microsoft.Bcl.AsyncInterfaces** (`{wordAsyncInterfacesVersion}`)",
            installation,
            StringComparison.Ordinal);
        Assert.Contains($"System.Text.Json** (`{systemTextJsonVersion}`)", installation, StringComparison.Ordinal);

        string aotGuide = File.ReadAllText(Path.Combine(
            repositoryRoot,
            "Website",
            "content",
            "docs",
            "advanced",
            "aot-trimming",
            "index.md"));
        Assert.Contains($"`{openXmlVersion}`", aotGuide, StringComparison.Ordinal);

        string migration = File.ReadAllText(Path.Combine(repositoryRoot, "MIGRATION.md"));
        Assert.Contains("# Upgrading OfficeIMO", migration, StringComparison.Ordinal);
        Assert.Contains("GitHub Releases", migration, StringComparison.Ordinal);
        Assert.Contains("## OfficeIMO 3.0 to 3.1", migration, StringComparison.Ordinal);
        Assert.Contains("## OfficeIMO 2.x to 3.0", migration, StringComparison.Ordinal);
        Assert.Contains("## OfficeIMO 1.x to 2.0", migration, StringComparison.Ordinal);
        Assert.Contains("SaveTablesAsExcel", migration, StringComparison.Ordinal);
        Assert.Contains("SaveAsPowerPoint", migration, StringComparison.Ordinal);
        Assert.Contains("ImportTablesToExcelDocument", migration, StringComparison.Ordinal);
        Assert.Contains("ToPowerPointPresentation", migration, StringComparison.Ordinal);
        Assert.Contains("`PdfWordReadOptions`", migration, StringComparison.Ordinal);
        Assert.Contains("`PdfWordImportOptions`", migration, StringComparison.Ordinal);
        Assert.Contains("`PdfPowerPointTableImportOptions`", migration, StringComparison.Ordinal);
        Assert.Contains("`PdfPowerPointImportOptions`", migration, StringComparison.Ordinal);
        Assert.Contains("`ImportTablesToPowerPointPresentation`", migration, StringComparison.Ordinal);
        Assert.Contains("### PowerPoint lifecycle, composition, and inspection", migration, StringComparison.Ordinal);
        Assert.Contains("`OpenRead(path)`", migration, StringComparison.Ordinal);
        Assert.Contains("`PowerPointTemplate.Inspect(path)`", migration, StringComparison.Ordinal);
        Assert.Contains("PowerPointCompositionOptions.FromBrief(brief)", migration, StringComparison.Ordinal);
        Assert.Contains("`InspectPreflight()`", migration, StringComparison.Ordinal);
        Assert.Contains("`OfficeIMO.Drawing.OfficeChartData`", migration, StringComparison.Ordinal);
        Assert.Contains("`WithDpi(...)`", migration, StringComparison.Ordinal);
        Assert.Contains("`ForHighResolution(...)`", migration, StringComparison.Ordinal);
        Assert.Contains("`OfficeImageExportFileConflictPolicy.FailIfExists`", migration, StringComparison.Ordinal);
        Assert.Contains("`Replace` or `CreateUnique`", migration, StringComparison.Ordinal);
        Assert.Contains("`AllowSystemFontEmbedding`", migration, StringComparison.Ordinal);
        Assert.Contains("`ResourcePolicy.AllowSystemFontEmbedding`", migration, StringComparison.Ordinal);
        Assert.Contains("Markdown `IncludeLocalImages`", migration, StringComparison.Ordinal);
        Assert.Contains("`ResourcePolicy.AllowLocalFileAccess`", migration, StringComparison.Ordinal);
        Assert.Contains("Markdown `IncludeDataUriImages`", migration, StringComparison.Ordinal);
        Assert.Contains("`ResourcePolicy.AllowDataUris`", migration, StringComparison.Ordinal);
        Assert.Contains("`OfficeDocumentReadResultSchema.Version`", migration, StringComparison.Ordinal);
        Assert.Contains("`OfficeDocumentReadResultSchema.CurrentVersion`", migration, StringComparison.Ordinal);
        Assert.Contains("`ApplyWordLikeTheme()`", migration, StringComparison.Ordinal);
        Assert.Contains("`ApplyDefaultTheme()`", migration, StringComparison.Ordinal);
        Assert.Contains("`UseFrontMatterVisualTheme`", migration, StringComparison.Ordinal);
        Assert.Contains("`UseFrontMatterTheme`", migration, StringComparison.Ordinal);
        Assert.Contains("`SaveResult` / `SaveResultAsync`", migration, StringComparison.Ordinal);
        Assert.Contains("`Save` / `SaveAsync` returning `OdfSaveResult`", migration, StringComparison.Ordinal);
        Assert.Contains("`ToBytesResult`", migration, StringComparison.Ordinal);
        Assert.Contains("`Serialize` returning `OdfSaveResult`", migration, StringComparison.Ordinal);
        Assert.Contains("Word `IncludePageNumbers`", migration, StringComparison.Ordinal);
        Assert.Contains("Excel `IncludeSheetHeadings`", migration, StringComparison.Ordinal);
        Assert.Contains("`SourceConversionReports`", migration, StringComparison.Ordinal);
        Assert.Contains("`OdfConversionReport.Mappings`", migration, StringComparison.Ordinal);
        Assert.Contains("`ODF_*`", migration, StringComparison.Ordinal);
        Assert.Contains("`CsvDataReaderOptions`", migration, StringComparison.Ordinal);
        Assert.Contains("`MaxXlsbCells`", migration, StringComparison.Ordinal);
        Assert.Contains("`MaxDataReaderBufferedCells`", migration, StringComparison.Ordinal);
        Assert.Contains("`CsvFieldSpanAction`", migration, StringComparison.Ordinal);
        Assert.Contains("`ICsvProjectedFieldSpanVisitor`", migration, StringComparison.Ordinal);
        Assert.Contains("`PdfWordImportOptions.CreateTablesOnly()`", migration, StringComparison.Ordinal);
        Assert.Contains("`PdfPowerPointConversionReport.TableEntries`", migration, StringComparison.Ordinal);
        Assert.Contains("`RtfDocument.ReadAsync(string)`", migration, StringComparison.Ordinal);
        Assert.Contains("`RtfDocument.LoadAsync(byte[])`", migration, StringComparison.Ordinal);
        Assert.Contains("Save and associate a path", migration, StringComparison.Ordinal);
        Assert.Contains("Write once to a caller-owned stream", migration, StringComparison.Ordinal);
        Assert.Contains("does not replace the document's associated path or source stream", migration, StringComparison.Ordinal);
        Assert.Contains("`CommentsByObjectType`", migration, StringComparison.Ordinal);
        Assert.Contains("`DataValidationsByType`", migration, StringComparison.Ordinal);
        Assert.Contains("`HasImportErrors`", migration, StringComparison.Ordinal);
        Assert.Contains("`HasUnsupportedFeatures`", migration, StringComparison.Ordinal);
        Assert.Contains("public `Diagnostics`, `UnsupportedFeatures`, `PreservedFeatures`, `UnsupportedSheets`, and `CompoundFeatures` collections", migration, StringComparison.Ordinal);
        string[] retainedLegacyMigrationNames = {
            "`SheetComposer.DefinitionList(...)`",
            "`PowerPointUnits.Cm/Mm/Inches/Points(...)`",
            "`VisioDocument.UseMastersFromTemplate(...)`",
            "`OrderedListBlock.ListItems` / `UnorderedListBlock.ListItems`",
            "`ListItem.Children`",
            "`QuoteBlock.Children` / `DetailsBlock.Children`",
            "`TableCell.Blocks` / `DefinitionListDefinition.Blocks`",
            "`FootnoteDefinitionBlock.Blocks`",
            "tuple-based `DefinitionListBlock.Items`",
            "`OutlookContact.Email1Address`",
            "phone compatibility properties",
            "`TrackComments`",
            "`ImageShapeStyleHelper`",
            "`HorizontalAlignmentHelper`",
            "`WasLoadedFromLegacyDoc`",
            "`MaxWordDocumentStreamBytes`",
            "`ReportUnsupportedFeatures`",
            "`WasLoadedFromLegacyXls`",
            "`MaxWorkbookStreamBytes`",
            "`ReportUnsupportedRecords`",
            "`AsciiDocPdfSaveOptions.PdfOptions`",
            "`LatexPdfSaveOptions.PdfOptions`",
            "`OneNotePdfSaveOptions.PdfOptions`",
            "`SavePdfAsWord()` / `SavePdfAsRtf()`",
            "`SavePdfTablesAsExcel/Word/PowerPoint()`"
        };
        foreach (string legacyName in retainedLegacyMigrationNames) {
            Assert.Contains(legacyName, migration, StringComparison.Ordinal);
        }
        Assert.DoesNotContain("SaveAs{Format}FromPdfTables", migration, StringComparison.Ordinal);
        Assert.DoesNotContain("To{Format}BytesFromPdfTables", migration, StringComparison.Ordinal);

        string rootReadme = File.ReadAllText(Path.Combine(repositoryRoot, "README.md"));
        Assert.Contains("Save and associate a path", rootReadme, StringComparison.Ordinal);
        Assert.Contains("Write once to a caller-owned stream", rootReadme, StringComparison.Ordinal);
        Assert.Contains("does not replace the document's associated path or source stream", rootReadme, StringComparison.Ordinal);
        Assert.DoesNotContain("Save and associate a path or stream", rootReadme, StringComparison.Ordinal);
        Assert.Contains("PDF -->|\"visual pages or editable tables\"| PowerPoint", rootReadme, StringComparison.Ordinal);
        Assert.DoesNotContain("PDF -->|\"logical tables only\"| PowerPoint", rootReadme, StringComparison.Ordinal);
        Assert.Contains("ignored-error metadata preservation", rootReadme, StringComparison.Ordinal);

        string visioReadme = File.ReadAllText(Path.Combine(repositoryRoot, "OfficeIMO.Visio", "README.md"));
        Assert.Contains("### Loaded-diagram compatibility boundary", visioReadme, StringComparison.Ordinal);
        Assert.Contains("advanced nested and container behavior", visioReadme, StringComparison.Ordinal);
        Assert.Contains("richer threaded comment and author workflows", visioReadme, StringComparison.Ordinal);
        Assert.Contains("broader whole-diagram relayout and polish", visioReadme, StringComparison.Ordinal);
        Assert.Contains("rather than a complete typed object model", visioReadme, StringComparison.Ordinal);

        string pdfCurrentState = File.ReadAllText(Path.Combine(repositoryRoot, "Docs", "officeimo.pdf.current-state.md"));
        Assert.Contains("| PDF to Excel |", pdfCurrentState, StringComparison.Ordinal);
        Assert.Contains("| PDF to PowerPoint |", pdfCurrentState, StringComparison.Ordinal);
        Assert.Contains("`PdfPowerPointImportOptions` defaults to `VisualPages`", pdfCurrentState, StringComparison.Ordinal);
        Assert.Contains("`CreateEditableTables()` / `EditableTables`", pdfCurrentState, StringComparison.Ordinal);
        Assert.Contains("omitted non-table page content", pdfCurrentState, StringComparison.Ordinal);

        string excelCompatibility = File.ReadAllText(Path.Combine(
            repositoryRoot,
            "OfficeIMO.Excel",
            "COMPATIBILITY.md"));
        Assert.Contains("## Formula evaluator", excelCompatibility, StringComparison.Ordinal);
        Assert.Contains("`FORECAST.LINEAR`", excelCompatibility, StringComparison.Ordinal);
        Assert.Contains("`WORKDAY.INTL`", excelCompatibility, StringComparison.Ordinal);
        Assert.Contains("Custom application functions", excelCompatibility, StringComparison.Ordinal);
        int formulaEvaluatorIndex = excelCompatibility.IndexOf("## Formula evaluator", StringComparison.Ordinal);
        Assert.True(formulaEvaluatorIndex >= 0);
        int documentedFormulaFunctions = Regex.Matches(
                excelCompatibility.Substring(formulaEvaluatorIndex),
                @"`([A-Z][A-Z0-9]*(?:\.[A-Z]+)?)`")
            .Cast<Match>()
            .Select(static match => match.Groups[1].Value)
            .Distinct(StringComparer.Ordinal)
            .Count();
        Assert.Equal(151, documentedFormulaFunctions);

        string releasesPage = File.ReadAllText(Path.Combine(
            repositoryRoot,
            "Website",
            "content",
            "pages",
            "releases.md"));
        Assert.Contains("  - /changelog/", releasesPage, StringComparison.Ordinal);
        Assert.Contains("  - /docs/workflows/release-previews/", releasesPage, StringComparison.Ordinal);
        Assert.Contains("title: \"Releases and Downloads\"", releasesPage, StringComparison.Ordinal);

        string documentationIndex = File.ReadAllText(Path.Combine(repositoryRoot, "Docs", "README.md"));
        Assert.Contains("# OfficeIMO documentation", documentationIndex, StringComparison.Ordinal);
        Assert.Contains("Use this index to find package guides", documentationIndex, StringComparison.Ordinal);
        Assert.Contains("Package READMEs are the best starting point", documentationIndex, StringComparison.Ordinal);
        Assert.Contains("[Migration guide](../MIGRATION.md)", documentationIndex, StringComparison.Ordinal);
        Assert.Contains("[GitHub Releases](https://github.com/EvotecIT/OfficeIMO/releases)", documentationIndex, StringComparison.Ordinal);
        Assert.Contains("These reports are generated from repository catalogs", documentationIndex, StringComparison.Ordinal);
        Assert.DoesNotContain("Documentation ownership rules", documentationIndex, StringComparison.Ordinal);

        string agentInstructions = File.ReadAllText(Path.Combine(repositoryRoot, "AGENTS.md"));
        Assert.Contains("## Documentation audiences", agentInstructions, StringComparison.Ordinal);
        Assert.Contains("root README and package READMEs are user-facing", agentInstructions, StringComparison.Ordinal);
        Assert.Contains("`MIGRATION.md` is the user-facing upgrade contract", agentInstructions, StringComparison.Ordinal);
        Assert.Contains("Docs/README.md` is a navigation page", agentInstructions, StringComparison.Ordinal);
        Assert.Contains("Docs/ROADMAP.md` is the single product backlog", agentInstructions, StringComparison.Ordinal);
        Assert.Contains("Do not put agent workflow", agentInstructions, StringComparison.Ordinal);
        Assert.Contains("## Documentation maintenance", agentInstructions, StringComparison.Ordinal);
        Assert.Contains("Do not add release-wait", agentInstructions, StringComparison.Ordinal);

        string roadmap = File.ReadAllText(Path.Combine(repositoryRoot, "Docs", "ROADMAP.md"));
        Assert.Contains("# OfficeIMO roadmap", roadmap, StringComparison.Ordinal);
        Assert.Contains("It contains open work only", roadmap, StringComparison.Ordinal);
        Assert.Contains("## Word", roadmap, StringComparison.Ordinal);
        Assert.Contains("## Excel", roadmap, StringComparison.Ordinal);
        Assert.Contains("## PDF, HTML, and image rendering", roadmap, StringComparison.Ordinal);
        Assert.Contains("## Markdown and text formats", roadmap, StringComparison.Ordinal);
        Assert.Contains("## Reader and document intelligence", roadmap, StringComparison.Ordinal);
        Assert.Contains("## Email, stores, and cloud adapters", roadmap, StringComparison.Ordinal);
        Assert.Contains("## Browser and agent surfaces", roadmap, StringComparison.Ordinal);
        Assert.Contains("Complete XML-signature validation", roadmap, StringComparison.Ordinal);
        Assert.Contains("cross-platform package signing", roadmap, StringComparison.Ordinal);
        Assert.Contains("macro-project signing", roadmap, StringComparison.Ordinal);
        Assert.Contains("allowed-edit ranges and ignored-error regions", roadmap, StringComparison.Ordinal);
        Assert.Contains("relationship-backed drawings, workbook-level structures, charts, and template bindings", roadmap, StringComparison.Ordinal);
        Assert.Contains("### Image-export evidence", roadmap, StringComparison.Ordinal);
        Assert.DoesNotContain("Publish the browser conversion playground", roadmap, StringComparison.Ordinal);
        Assert.DoesNotContain("Add Visio and portable-document adapters", roadmap, StringComparison.Ordinal);

        string markdownRoundtripDesign = File.ReadAllText(Path.Combine(
            repositoryRoot,
            "Docs",
            "officeimo.markdown.lossless-roundtrip-design.md"));
        Assert.Contains("MarkdownParseResult parsed = MarkdownReader.ParseWithSyntaxTree(", markdownRoundtripDesign, StringComparison.Ordinal);
        Assert.DoesNotContain("MarkdownParseResult parsed = MarkdownReader.Parse(", markdownRoundtripDesign, StringComparison.Ordinal);

        string legacyXlsCompatibility = File.ReadAllText(Path.Combine(repositoryRoot, "Docs", "officeimo.excel.legacy-xls-compatibility.md"));
        string legacyDocCompatibility = File.ReadAllText(Path.Combine(repositoryRoot, "Docs", "officeimo.word.legacy-doc-compatibility.md"));
        Assert.DoesNotContain("## Breaking API cleanup", legacyXlsCompatibility, StringComparison.Ordinal);
        Assert.DoesNotContain("## Breaking API cleanup", legacyDocCompatibility, StringComparison.Ordinal);
        Assert.DoesNotContain("`WasLoadedFromLegacyXls`", legacyXlsCompatibility, StringComparison.Ordinal);
        Assert.DoesNotContain("`WasLoadedFromLegacyDoc`", legacyDocCompatibility, StringComparison.Ordinal);

        string wordCompatibility = File.ReadAllText(Path.Combine(
            repositoryRoot,
            "OfficeIMO.Word",
            "COMPATIBILITY.md"));
        Assert.Contains("## Field evaluator", wordCompatibility, StringComparison.Ordinal);
        Assert.Contains("`CREATEDATE`, `SAVEDATE`, `PRINTDATE`", wordCompatibility, StringComparison.Ordinal);
        Assert.Contains("`SECTION` / `SECTIONPAGES`", wordCompatibility, StringComparison.Ordinal);
        Assert.Contains("`SUM`, `AVERAGE`, `MIN`, `MAX`, `PRODUCT`, `COUNT`", wordCompatibility, StringComparison.Ordinal);
        Assert.Contains("`DEFINED`", wordCompatibility, StringComparison.Ordinal);
        Assert.Contains("`ABOVE`, `BELOW`, `LEFT`, and `RIGHT`", wordCompatibility, StringComparison.Ordinal);

        string excelRemoteLoading = File.ReadAllText(Path.Combine(
            repositoryRoot,
            "Docs",
            "officeimo.excel.remote-loading.md"));
        Assert.Contains("workbook.CreateDataReader()", excelRemoteLoading, StringComparison.Ordinal);
        Assert.DoesNotContain("ExcelDocumentReader", excelRemoteLoading, StringComparison.Ordinal);

        Assert.True(File.Exists(Path.Combine(repositoryRoot, "MIGRATION.md")));
        Assert.False(File.Exists(Path.Combine(repositoryRoot, "CHANGELOG.MD")));
        Assert.False(File.Exists(Path.Combine(repositoryRoot, "powerpoint.md")));
        Assert.False(File.Exists(Path.Combine(repositoryRoot, "OfficeIMO.Markup", "IMPLEMENTATION_PLAN.md")));
        Assert.False(File.Exists(Path.Combine(repositoryRoot, "Docs", "officeimo-3.0-public-api-review.md")));
    }

    private static string[] AssertDotNetInstallCommands(
        string content,
        ISet<string> releasePackageIds,
        int expectedCount,
        string expectedVersion,
        IReadOnlyCollection<string>? toolPackageIds = null) {
        MatchCollection commands = Regex.Matches(
            content,
            @"^dotnet (?<verb>add package|tool install --global) (?<id>OfficeIMO\.[^\s]+)(?<arguments>[^\r\n]*)\r?$",
            RegexOptions.Multiline | RegexOptions.CultureInvariant);

        Assert.Equal(expectedCount, commands.Count);
        string[] packageIds = commands
            .Cast<Match>()
            .Select(match => {
                string packageId = match.Groups["id"].Value;
                Assert.True(
                    releasePackageIds.Contains(packageId),
                    "Install command references an unknown package: " + match.Value);
                Assert.Equal($" --version {expectedVersion}", match.Groups["arguments"].Value);

                bool isTool = toolPackageIds?.Contains(packageId, StringComparer.OrdinalIgnoreCase) == true;
                Assert.Equal(
                    isTool ? "tool install --global" : "add package",
                    match.Groups["verb"].Value);
                return packageId;
            })
            .ToArray();

        Assert.Equal(
            expectedCount,
            packageIds.Distinct(StringComparer.OrdinalIgnoreCase).Count());
        return packageIds;
    }

    private static void AssertPackageManagerInstallCommands(
        string content,
        ISet<string> releasePackageIds,
        int expectedCount,
        string expectedVersion) {
        MatchCollection commands = Regex.Matches(
            content,
            @"^Install-Package (?<id>OfficeIMO\.[^\s]+)(?<arguments>[^\r\n]*)\r?$",
            RegexOptions.Multiline | RegexOptions.CultureInvariant);

        Assert.Equal(expectedCount, commands.Count);
        string[] packageIds = commands
            .Cast<Match>()
            .Select(match => {
                string packageId = match.Groups["id"].Value;
                Assert.True(
                    releasePackageIds.Contains(packageId),
                    "Package Manager command references an unknown package: " + match.Value);
                Assert.Equal($" -Version {expectedVersion}", match.Groups["arguments"].Value);
                return packageId;
            })
            .ToArray();

        Assert.Equal(
            expectedCount,
            packageIds.Distinct(StringComparer.OrdinalIgnoreCase).Count());
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

    private static string ReadPackageReferenceVersion(
        string repositoryRoot,
        string relativeProjectPath,
        string packageId) {
        XDocument document = XDocument.Load(Path.Combine(repositoryRoot, relativeProjectPath));
        XElement packageReference = Assert.Single(
            document.Descendants(),
            element =>
                string.Equals(element.Name.LocalName, "PackageReference", StringComparison.Ordinal) &&
                string.Equals((string?)element.Attribute("Include"), packageId, StringComparison.OrdinalIgnoreCase));
        return (string?)packageReference.Attribute("Version")
            ?? throw new InvalidDataException(
                $"Package reference '{packageId}' in '{relativeProjectPath}' does not declare a Version attribute.");
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

    private static int CountProjectHeadings(string readme, string sectionName) {
        string marker = "### " + sectionName;
        int sectionStart = readme.IndexOf(marker, StringComparison.Ordinal);
        Assert.True(sectionStart >= 0, "README section is missing: " + sectionName);
        int nextSection = readme.IndexOf("\n### ", sectionStart + marker.Length, StringComparison.Ordinal);
        string section = nextSection >= 0
            ? readme.Substring(sectionStart, nextSection - sectionStart)
            : readme.Substring(sectionStart);
        return Regex.Matches(
            section,
            @"^#### \[OfficeIMO\.",
            RegexOptions.Multiline | RegexOptions.CultureInvariant).Count;
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
