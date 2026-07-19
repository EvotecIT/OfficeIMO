using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class PackageDependencyGuardrailTests {
    private const string CurrentMarkdigVersion = "1.3.2";

    private static readonly string[] ForbiddenRenderingPackageIds = [
        "Aspose.Cells",
        "Aspose.PDF",
        "Aspose.Slides",
        "Aspose.Words",
        "Docnet.Core",
        "GemBox.Document",
        "GemBox.Pdf",
        "GemBox.Presentation",
        "GemBox.Spreadsheet",
        "Magick.NET-Q16-AnyCPU",
        "Magick.NET-Q8-AnyCPU",
        "Microsoft.Office.Interop.Excel",
        "Microsoft.Office.Interop.PowerPoint",
        "Microsoft.Office.Interop.Visio",
        "Microsoft.Office.Interop.Word",
        "Microsoft.Playwright",
        "PdfiumViewer",
        "PDFiumSharp",
        "PDFtoImage",
        "PuppeteerSharp",
        "Selenium.WebDriver",
        "SixLabors.ImageSharp",
        "SixLabors.Fonts",
        "SkiaSharp.Extended",
        "SkiaSharp",
        "SkiaSharp.Views",
        "Spire.Doc",
        "Spire.PDF",
        "Spire.Presentation",
        "Spire.XLS",
        "Syncfusion.DocIO.Net.Core",
        "Syncfusion.EJ2.PdfViewer",
        "Syncfusion.Pdf.Net.Core",
        "Syncfusion.PdfToImageConverter.Net",
        "Syncfusion.Presentation.Net.Core",
        "Syncfusion.XlsIO.Net.Core",
        "System.Drawing.Common"
    ];

    private static readonly string[] DocumentImageRenderingRoots = [
        "OfficeIMO.Drawing",
        "OfficeIMO.Excel",
        "OfficeIMO.Excel.Pdf",
        "OfficeIMO.Markdown.Pdf",
        "OfficeIMO.Pdf",
        "OfficeIMO.PowerPoint",
        "OfficeIMO.PowerPoint.Pdf",
        "OfficeIMO.Visio",
        "OfficeIMO.Word",
        "OfficeIMO.Word.Pdf"
    ];

    private static readonly string[] ForbiddenDocumentImageRenderingPackageIds = [
        ..ForbiddenRenderingPackageIds,
        "Microsoft.Office.Interop.Visio",
        "Microsoft.Web.WebView2"
    ];

    private static readonly string[] ForbiddenDocumentImageRenderingSourceTerms = [
        "Aspose.",
        "Excel.Application",
        "GemBox.",
        "ImageMagick.",
        "LibreOffice",
        "Microsoft.Office.Interop",
        "Microsoft.Playwright",
        "Microsoft.Web.WebView2",
        "Pdfium",
        "PDFium",
        "pdftocairo",
        "pdftoppm",
        "Poppler",
        "PowerPoint.Application",
        "PuppeteerSharp",
        "Selenium.",
        "SixLabors.",
        "SkiaSharp.",
        "soffice",
        "Spire.",
        "Syncfusion.",
        "System.Drawing.",
        "using System.Drawing;",
        "Visio.Application",
        "Word.Application"
    ];

    private static readonly Dictionary<string, HashSet<string>> ApprovedExternalCompatibilitySourceTerms =
        new(StringComparer.OrdinalIgnoreCase) {
            ["OfficeIMO.PowerPoint/PowerPointDesktopReferenceRenderer.cs"] =
                new HashSet<string>(["PowerPoint.Application"], StringComparer.OrdinalIgnoreCase),
            ["OfficeIMO.PowerPoint/PowerPointCompatibilityReport.cs"] =
                new HashSet<string>(["LibreOffice", "soffice"], StringComparer.OrdinalIgnoreCase)
        };

    [Fact]
    public void MarkdownParityProjects_UseTheSameCurrentMarkdigBaseline() {
        string[] projectPaths = [
            "OfficeIMO.Shared.Tests/OfficeIMO.Shared.Tests.csproj",
            "OfficeIMO.Markdown.Tests/OfficeIMO.Markdown.Tests.csproj",
            "OfficeIMO.Markdown.Benchmarks/OfficeIMO.Markdown.Benchmarks.csproj"
        ];

        foreach (var relativeProjectPath in projectPaths) {
            var projectPath = GetRepositoryPath(relativeProjectPath);
            Assert.True(File.Exists(projectPath), "Project file is missing: " + projectPath);

            Assert.Equal(CurrentMarkdigVersion, GetPackageReferenceVersion(projectPath, "Markdig"));
        }
    }

    [Fact]
    public void MarkdownCompatibilityDocs_TrackCurrentMarkdigBaselineVersion() {
        string compatibilityMatrix = File.ReadAllText(GetRepositoryPath("Docs/officeimo.markdown.compatibility-matrix.md"));
        Assert.Contains($"| External comparison package | Markdig `{CurrentMarkdigVersion}`", compatibilityMatrix, StringComparison.Ordinal);

        string competitorRoadmap = File.ReadAllText(GetRepositoryPath("Docs/officeimo.markdown.markdig-competitor-roadmap.md"));
        Assert.Contains($"external parity baseline: Markdig `{CurrentMarkdigVersion}`", competitorRoadmap, StringComparison.Ordinal);

        string correctnessBacklog = File.ReadAllText(GetRepositoryPath("Docs/officeimo.markdown.correctness-backlog.md"));
        Assert.Contains($"`OfficeIMO.Shared.Tests`, `OfficeIMO.Markdown.Tests`, and `OfficeIMO.Markdown.Benchmarks` all reference Markdig `{CurrentMarkdigVersion}`", correctnessBacklog, StringComparison.Ordinal);

        string packageCompatibility = File.ReadAllText(GetRepositoryPath("OfficeIMO.Markdown/COMPATIBILITY.md"));
        Assert.Contains($"curated Markdig {CurrentMarkdigVersion} parity cases", packageCompatibility, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownCompatibilityDocs_TrackCurrentFixtureBaselineCounts() {
        int commonMarkFixtureCount = CountJsonArrayEntries("OfficeIMO.Markdown.Tests/Markdown/Fixtures/CommonMark/commonmark-0.31.2-smoke.json");
        int gfmFixtureCount = CountJsonArrayEntries("OfficeIMO.Markdown.Tests/Markdown/Fixtures/GitHubFlavoredMarkdown/cmark-gfm-extensions-smoke.json");

        string compatibilityMatrix = File.ReadAllText(GetRepositoryPath("Docs/officeimo.markdown.compatibility-matrix.md"));
        Assert.Contains($"| CommonMark reference | {commonMarkFixtureCount} of 652 official CommonMark `0.31.2` examples pinned as smoke fixtures |", compatibilityMatrix, StringComparison.Ordinal);
        Assert.Contains($"| GFM reference | {gfmFixtureCount} cmark-gfm extension smoke fixtures plus focused OfficeIMO supplements for upstream ignored-autolink crash and query/fragment autolink regressions |", compatibilityMatrix, StringComparison.Ordinal);

        string competitorRoadmap = File.ReadAllText(GetRepositoryPath("Docs/officeimo.markdown.markdig-competitor-roadmap.md"));
        Assert.Contains($"standards smoke baseline: {commonMarkFixtureCount} CommonMark `0.31.2` fixtures, {gfmFixtureCount} cmark-gfm extension fixtures", competitorRoadmap, StringComparison.Ordinal);

        string parityGapPlan = File.ReadAllText(GetRepositoryPath("Docs/officeimo.markdown.markdig-parity-gap-plan.md"));
        Assert.Contains($"| CommonMark corpus | {commonMarkFixtureCount} of 652 official CommonMark `0.31.2` examples pinned as smoke fixtures |", parityGapPlan, StringComparison.Ordinal);
        Assert.Contains($"| GFM corpus | {gfmFixtureCount} cmark-gfm extension smoke fixtures plus focused crash/regression coverage |", parityGapPlan, StringComparison.Ordinal);

        string packageCompatibility = File.ReadAllText(GetRepositoryPath("OfficeIMO.Markdown/COMPATIBILITY.md"));
        Assert.Contains($"includes {commonMarkFixtureCount} pinned CommonMark 0.31.2 fixtures, {gfmFixtureCount} cmark-gfm smoke fixtures", packageCompatibility, StringComparison.Ordinal);
    }

    [Fact]
    public void Projects_DoNotReferenceImageSharpPackage() {
        var projectFiles = Directory.EnumerateFiles(GetRepositoryRoot(), "*.csproj", SearchOption.AllDirectories)
            .Where(static path => !path.Contains($"{Path.DirectorySeparatorChar}bin{Path.DirectorySeparatorChar}", StringComparison.OrdinalIgnoreCase))
            .Where(static path => !path.Contains($"{Path.DirectorySeparatorChar}obj{Path.DirectorySeparatorChar}", StringComparison.OrdinalIgnoreCase))
            .Where(static path => !path.Contains($"{Path.DirectorySeparatorChar}Ignore{Path.DirectorySeparatorChar}", StringComparison.OrdinalIgnoreCase))
            .Where(static path => new FileInfo(path).Length > 0)
            .ToArray();

        var offenders = projectFiles
            .Where(ProjectReferencesImageSharp)
            .ToArray();

        Assert.Empty(offenders);
    }

    [Fact]
    public void Projects_DoNotReferenceExternalGraphicsPackages() {
        var offenders = EnumerateProjectFiles()
            .Where(static projectPath => !IsNonProductionProject(projectPath))
            .SelectMany(projectPath => ProjectReferencesPackages(projectPath, ForbiddenRenderingPackageIds)
                .Select(packageId => GetRepositoryRelativePath(projectPath) + " -> " + packageId))
            .ToArray();

        Assert.Empty(offenders);
    }

    [Fact]
    public void DocumentImageRenderingProjects_DoNotReferenceExternalRenderersOrAutomation() {
        var offenders = EnumerateDocumentImageRenderingProjectFiles()
            .SelectMany(projectPath => ProjectReferencesPackages(projectPath, ForbiddenDocumentImageRenderingPackageIds)
                .Select(packageId => GetRepositoryRelativePath(projectPath) + " -> " + packageId))
            .ToArray();

        Assert.Empty(offenders);
    }

    [Fact]
    public void DocumentImageRenderingPaths_StayFirstPartyAndDependencyFree() {
        var sourceOffenders = new List<string>();
        foreach (string sourceFile in EnumerateDocumentImageRenderingSourceFiles()) {
            string source = File.ReadAllText(sourceFile);
            string relativePath = GetRepositoryRelativePath(sourceFile).Replace('\\', '/');
            foreach (string forbiddenTerm in ForbiddenDocumentImageRenderingSourceTerms) {
                if (ContainsForbiddenDocumentImageRenderingSourceTerm(source, forbiddenTerm) &&
                    !IsApprovedExternalCompatibilitySourceTerm(relativePath, forbiddenTerm)) {
                    sourceOffenders.Add(relativePath + " -> " + forbiddenTerm);
                }
            }
        }

        Assert.Empty(sourceOffenders);
    }

    [Fact]
    public void ImageExportGapManifest_StaysFirstPartyAndActionable() {
        string manifestPath = GetRepositoryPath("Docs/officeimo.image-export-gap-manifest.json");
        Assert.True(File.Exists(manifestPath), "Image export gap manifest is missing: " + manifestPath);

        using JsonDocument document = JsonDocument.Parse(File.ReadAllText(manifestPath));
        JsonElement root = document.RootElement;

        Assert.Equal("first-party-only", root.GetProperty("runtimeDependencyPolicy").GetString());
        JsonElement workstreams = root.GetProperty("workstreams");
        Assert.True(workstreams.GetArrayLength() >= 5, "Image export goal should track the main document and QA workstreams.");

        foreach (JsonElement workstream in workstreams.EnumerateArray()) {
            string id = workstream.GetProperty("id").GetString() ?? string.Empty;
            string owner = workstream.GetProperty("owner").GetString() ?? string.Empty;
            string policy = workstream.GetProperty("runtimeDependencyPolicy").GetString() ?? string.Empty;
            string status = workstream.GetProperty("status").GetString() ?? string.Empty;
            JsonElement nextSlices = workstream.GetProperty("nextSlices");

            Assert.False(string.IsNullOrWhiteSpace(id), "Every image export workstream needs an id.");
            Assert.StartsWith("OfficeIMO.", owner, StringComparison.Ordinal);
            Assert.Equal("first-party-only", policy);
            Assert.Contains(status, new[] { "active", "planned" });
            Assert.True(nextSlices.GetArrayLength() > 0, "Workstream '" + id + "' needs at least one next slice.");
        }
    }

    [Fact]
    public void Projects_DoNotReferenceSixLaborsFontsPackage() {
        var projectFiles = Directory.EnumerateFiles(GetRepositoryRoot(), "*.csproj", SearchOption.AllDirectories)
            .Where(static path => !path.Contains($"{Path.DirectorySeparatorChar}bin{Path.DirectorySeparatorChar}", StringComparison.OrdinalIgnoreCase))
            .Where(static path => !path.Contains($"{Path.DirectorySeparatorChar}obj{Path.DirectorySeparatorChar}", StringComparison.OrdinalIgnoreCase))
            .Where(static path => !path.Contains($"{Path.DirectorySeparatorChar}Ignore{Path.DirectorySeparatorChar}", StringComparison.OrdinalIgnoreCase))
            .Where(static path => new FileInfo(path).Length > 0)
            .ToArray();

        var offenders = projectFiles
            .Where(ProjectReferencesSixLaborsFonts)
            .ToArray();

        Assert.Empty(offenders);
    }

    [Theory]
    [InlineData("OfficeIMO.Word.Rtf/OfficeIMO.Word.Rtf.csproj")]
    [InlineData("OfficeIMO.Rtf.Pdf/OfficeIMO.Rtf.Pdf.csproj")]
    [InlineData("OfficeIMO.Drawing/OfficeIMO.Drawing.csproj")]
    [InlineData("OfficeIMO.Pdf/OfficeIMO.Pdf.csproj")]
    [InlineData("OfficeIMO.Word.Pdf/OfficeIMO.Word.Pdf.csproj")]
    [InlineData("OfficeIMO.Excel.Pdf/OfficeIMO.Excel.Pdf.csproj")]
    [InlineData("OfficeIMO.Markdown.Pdf/OfficeIMO.Markdown.Pdf.csproj")]
    [InlineData("OfficeIMO.PowerPoint.Pdf/OfficeIMO.PowerPoint.Pdf.csproj")]
    [InlineData("OfficeIMO.Html.Pdf/OfficeIMO.Html.Pdf.csproj")]
    [InlineData("OfficeIMO.Reader.Pdf/OfficeIMO.Reader.Pdf.csproj")]
    [InlineData("OfficeIMO.Reader.Rtf/OfficeIMO.Reader.Rtf.csproj")]
    public void DependencyLightProjects_HaveNoPackageReferences(string relativeProjectPath) {
        var projectPath = GetRepositoryPath(relativeProjectPath);
        Assert.True(File.Exists(projectPath), "Project file is missing: " + projectPath);

        var document = XDocument.Load(projectPath);
        var ns = document.Root?.Name.Namespace ?? XNamespace.None;

        var references = document
            .Descendants(ns + "PackageReference")
            .Select(static e => (string?)e.Attribute("Include") ?? string.Empty)
            .Where(static include => !string.IsNullOrWhiteSpace(include))
            .ToArray();

        Assert.Empty(references);
    }

    [Fact]
    public void RtfCore_OnlyReferencesTheCodePageCompatibilityPackage() {
        var projectPath = GetRepositoryPath("OfficeIMO.Rtf/OfficeIMO.Rtf.csproj");
        Assert.True(File.Exists(projectPath), "Project file is missing: " + projectPath);

        var document = XDocument.Load(projectPath);
        var ns = document.Root?.Name.Namespace ?? XNamespace.None;
        var references = document
            .Descendants(ns + "PackageReference")
            .Select(static element => new {
                Id = (string?)element.Attribute("Include") ?? string.Empty,
                Version = (string?)element.Attribute("Version") ?? string.Empty
            })
            .ToArray();

        var reference = Assert.Single(references);
        Assert.Equal("System.Text.Encoding.CodePages", reference.Id);
        Assert.Equal("8.0.0", reference.Version);
    }

    [Fact]
    public void Email_DeclaresTheCodePageCompatibilityPackageItUsesDirectly() {
        var projectPath = GetRepositoryPath("OfficeIMO.Email/OfficeIMO.Email.csproj");
        Assert.True(File.Exists(projectPath), "Project file is missing: " + projectPath);

        var document = XDocument.Load(projectPath);
        var ns = document.Root?.Name.Namespace ?? XNamespace.None;
        var references = document
            .Descendants(ns + "PackageReference")
            .Select(static element => new {
                Id = (string?)element.Attribute("Include") ?? string.Empty,
                Version = (string?)element.Attribute("Version") ?? string.Empty
            })
            .ToArray();

        var reference = Assert.Single(references);
        Assert.Equal("System.Text.Encoding.CodePages", reference.Id);
        Assert.Equal("8.0.0", reference.Version);
    }

    [Theory]
    [InlineData("OfficeIMO.Examples/OfficeIMO.Examples.csproj")]
    [InlineData("OfficeIMO.Markup.Cli/OfficeIMO.Markup.Cli.csproj")]
    [InlineData("OfficeIMO.Excel.Benchmarks.LegacyEpPlus/OfficeIMO.Excel.Benchmarks.LegacyEpPlus.csproj")]
    public void RepositoryExecutablesAndBenchmarks_AreNotPublishedAsLibraryPackages(string relativeProjectPath) {
        var projectPath = GetRepositoryPath(relativeProjectPath);
        Assert.True(File.Exists(projectPath), "Project file is missing: " + projectPath);

        var document = XDocument.Load(projectPath);
        var ns = document.Root?.Name.Namespace ?? XNamespace.None;

        Assert.Equal("Exe", (string?)document.Descendants(ns + "OutputType").Single());
        Assert.Equal("false", (string?)document.Descendants(ns + "IsPackable").Single());
    }

    [Fact]
    public void RtfHtmlBridge_IsUnifiedIntoOfficeIMOHtml() {
        var projectPath = GetRepositoryPath("OfficeIMO.Html/OfficeIMO.Html.csproj");
        Assert.True(File.Exists(projectPath), "Project file is missing: " + projectPath);
        Assert.False(Directory.Exists(GetRepositoryPath("OfficeIMO.Rtf.Html")), "Retired RTF HTML project folder should not be restored.");
        Assert.False(Directory.Exists(GetRepositoryPath("OfficeIMO.Html.Rtf")), "Retired HTML RTF project folder should not be restored.");

        var document = XDocument.Load(projectPath);
        var ns = document.Root?.Name.Namespace ?? XNamespace.None;

        Assert.Equal("OfficeIMO.Html", (string?)document.Descendants(ns + "PackageId").Single());
        Assert.Equal("OfficeIMO.Html", (string?)document.Descendants(ns + "AssemblyName").Single());

        var exportedTypeNames = typeof(OfficeIMO.Html.HtmlToRtfOptions)
            .Assembly
            .GetExportedTypes()
            .Select(static type => type.FullName ?? type.Name)
            .ToArray();

        Assert.Contains("OfficeIMO.Html.HtmlToRtfOptions", exportedTypeNames);
        Assert.Contains("OfficeIMO.Html.RtfToHtmlOptions", exportedTypeNames);
        Assert.DoesNotContain(exportedTypeNames, static typeName => typeName.Contains(".RtfHtml", StringComparison.Ordinal));

        var projectReferences = document
            .Descendants(ns + "ProjectReference")
            .Select(static e => NormalizeProjectPath((string?)e.Attribute("Include")))
            .ToArray();

        Assert.Contains(projectReferences, static include => include.EndsWith("OfficeIMO.Rtf/OfficeIMO.Rtf.csproj", StringComparison.OrdinalIgnoreCase));
        Assert.DoesNotContain(projectReferences, static include => include.Contains("OfficeIMO.Rtf.Html", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void RetiredPackages_AreNotReferencedBySolutionOrProjects() {
        string[] retiredPackageIds = ["OfficeIMO.Rtf.Html", "OfficeIMO.Html.Rtf", "OfficeIMO.Reader.Text"];

        var solutionPath = GetRepositoryPath("OfficeIMO.sln");
        Assert.True(File.Exists(solutionPath), "Solution file is missing: " + solutionPath);

        var solutionText = File.ReadAllText(solutionPath);
        foreach (var retiredPackageId in retiredPackageIds) {
            Assert.DoesNotContain(retiredPackageId, solutionText, StringComparison.OrdinalIgnoreCase);
        }

        var projectFiles = EnumerateProjectFiles();
        foreach (var projectFile in projectFiles) {
            var document = XDocument.Load(projectFile);
            var ns = document.Root?.Name.Namespace ?? XNamespace.None;

            var packageIds = document
                .Descendants(ns + "PackageId")
                .Select(static element => (string?)element)
                .Where(static value => !string.IsNullOrWhiteSpace(value))
                .ToArray();

            var packageReferences = document
                .Descendants(ns + "PackageReference")
                .Select(static element => (string?)element.Attribute("Include"))
                .Where(static value => !string.IsNullOrWhiteSpace(value))
                .ToArray();

            var projectReferences = document
                .Descendants(ns + "ProjectReference")
                .Select(static element => NormalizeProjectPath((string?)element.Attribute("Include")))
                .Where(static value => !string.IsNullOrWhiteSpace(value))
                .ToArray();

            foreach (var retiredPackageId in retiredPackageIds) {
                Assert.DoesNotContain(packageIds, value => string.Equals(value, retiredPackageId, StringComparison.OrdinalIgnoreCase));
                Assert.DoesNotContain(packageReferences, value => string.Equals(value, retiredPackageId, StringComparison.OrdinalIgnoreCase));
                Assert.DoesNotContain(projectReferences, value => value.Contains(retiredPackageId, StringComparison.OrdinalIgnoreCase));
            }
        }

        var projectBuildPath = GetRepositoryPath("Build/project.build.json");
        Assert.True(File.Exists(projectBuildPath), "Project build file is missing: " + projectBuildPath);

        var projectBuildText = File.ReadAllText(projectBuildPath);
        foreach (var retiredPackageId in retiredPackageIds) {
            Assert.DoesNotContain(retiredPackageId, projectBuildText, StringComparison.OrdinalIgnoreCase);
        }
    }

    [Fact]
    public void RetiredRtfHtmlNamespaces_AreNotUsedBySourceFiles() {
        string[] retiredNamespaces = ["OfficeIMO.Rtf.Html", "OfficeIMO.Html.Rtf"];

        var sourceFiles = Directory.EnumerateFiles(GetRepositoryRoot(), "*.cs", SearchOption.AllDirectories)
            .Where(static path => !path.Contains($"{Path.DirectorySeparatorChar}bin{Path.DirectorySeparatorChar}", StringComparison.OrdinalIgnoreCase))
            .Where(static path => !path.Contains($"{Path.DirectorySeparatorChar}obj{Path.DirectorySeparatorChar}", StringComparison.OrdinalIgnoreCase))
            .Where(static path => !path.Contains($"{Path.DirectorySeparatorChar}Ignore{Path.DirectorySeparatorChar}", StringComparison.OrdinalIgnoreCase))
            .Where(static path => new FileInfo(path).Length > 0)
            .ToArray();

        foreach (var sourceFile in sourceFiles) {
            string source = File.ReadAllText(sourceFile);
            foreach (var retiredNamespace in retiredNamespaces) {
                Assert.DoesNotContain($"namespace {retiredNamespace}", source, StringComparison.Ordinal);
                Assert.DoesNotContain($"using {retiredNamespace}", source, StringComparison.Ordinal);
            }
        }
    }

    [Fact]
    public void RetiredAggregateTestAssembly_IsNotGrantedFriendAccess() {
        var projectOffenders = EnumerateProjectFiles()
            .SelectMany(projectPath => XDocument.Load(projectPath)
                .Descendants()
                .Where(static element => element.Name.LocalName == "InternalsVisibleTo")
                .Where(static element => string.Equals((string?)element.Attribute("Include"), "OfficeIMO.Tests", StringComparison.OrdinalIgnoreCase))
                .Select(_ => GetRepositoryRelativePath(projectPath)))
            .ToArray();

        var sourceOffenders = Directory.EnumerateFiles(GetRepositoryRoot(), "*.cs", SearchOption.AllDirectories)
            .Where(static path => !path.Contains($"{Path.DirectorySeparatorChar}bin{Path.DirectorySeparatorChar}", StringComparison.OrdinalIgnoreCase))
            .Where(static path => !path.Contains($"{Path.DirectorySeparatorChar}obj{Path.DirectorySeparatorChar}", StringComparison.OrdinalIgnoreCase))
            .Where(static path => !path.Contains($"{Path.DirectorySeparatorChar}Ignore{Path.DirectorySeparatorChar}", StringComparison.OrdinalIgnoreCase))
            .Where(static path => new FileInfo(path).Length > 0)
            .Where(SourceGrantsRetiredAggregateTestAccess)
            .Select(GetRepositoryRelativePath)
            .ToArray();

        var offenders = projectOffenders
            .Concat(sourceOffenders)
            .OrderBy(static offender => offender, StringComparer.OrdinalIgnoreCase)
            .ToArray();

        Assert.Empty(offenders);
    }

    [Fact]
    public void RtfPackages_AreIncludedInProjectBuildVersionMap() {
        var projectBuildPath = GetRepositoryPath("Build/project.build.json");
        Assert.True(File.Exists(projectBuildPath), "Project build file is missing: " + projectBuildPath);

        using JsonDocument document = JsonDocument.Parse(File.ReadAllText(projectBuildPath));
        string? expectedVersion = document.RootElement.GetProperty("ExpectedVersion").GetString();
        JsonElement expectedVersionMap = document.RootElement.GetProperty("ExpectedVersionMap");

        Assert.Equal(expectedVersion, expectedVersionMap.GetProperty("OfficeIMO.Rtf").GetString());
        Assert.Equal(expectedVersion, expectedVersionMap.GetProperty("OfficeIMO.Word.Rtf").GetString());
        Assert.Equal(expectedVersion, expectedVersionMap.GetProperty("OfficeIMO.Rtf.Pdf").GetString());
        Assert.Equal(expectedVersion, expectedVersionMap.GetProperty("OfficeIMO.Reader.Rtf").GetString());
    }

    [Theory]
    [InlineData("OfficeIMO.Reader.Ocr.Process")]
    [InlineData("OfficeIMO.Reader.Ocr.Tesseract")]
    public void ReaderOcrPackages_AreIncludedInProjectBuildVersionMap(string packageId) {
        var projectBuildPath = GetRepositoryPath("Build/project.build.json");
        Assert.True(File.Exists(projectBuildPath), "Project build file is missing: " + projectBuildPath);

        using JsonDocument document = JsonDocument.Parse(File.ReadAllText(projectBuildPath));
        string? expectedVersion = document.RootElement.GetProperty("ExpectedVersion").GetString();
        JsonElement expectedVersionMap = document.RootElement.GetProperty("ExpectedVersionMap");

        Assert.Equal(expectedVersion, expectedVersionMap.GetProperty(packageId).GetString());
    }

    [Fact]
    public void DrawingFoundation_UsesTheConsolidatedPackageIdentity() {
        string projectPath = GetRepositoryPath("OfficeIMO.Drawing/OfficeIMO.Drawing.csproj");
        var document = XDocument.Load(projectPath);
        XNamespace ns = document.Root?.Name.Namespace ?? XNamespace.None;
        string? packageId = document.Descendants(ns + "PackageId").Select(static element => element.Value).SingleOrDefault();

        Assert.Equal("OfficeIMO.Drawing", packageId);
    }

    [Fact]
    public void DrawingOwnedFoundationKernels_RemainInDrawing() {
        Assert.False(Directory.Exists(GetRepositoryPath("OfficeIMO.Shared")));

        string[] expectedDrawingOwnedFiles = [
            "ObjectDataHelpers.cs",
            "OfficeEncryption.cs",
            "OfficeFileConversion.cs",
            "Ole/OfficeOlePropertySetReader.cs",
            "Packaging/OfficeArchiveSafety.cs",
            "Streams/NonDisposingMemoryStream.cs"
        ];
        Assert.All(expectedDrawingOwnedFiles, relativePath =>
            Assert.True(
                File.Exists(GetRepositoryPath("OfficeIMO.Drawing/Internal/" + relativePath)),
                "Drawing-owned foundation file is missing: " + relativePath));

        string[] linkedFoundationSources = EnumerateProjectFiles()
            .SelectMany(projectPath => XDocument.Load(projectPath)
                .Descendants()
                .Where(static element => element.Name.LocalName == "Compile")
                .Select(element => new {
                    Project = GetRepositoryRelativePath(projectPath),
                    Include = NormalizeProjectPath((string?)element.Attribute("Include"))
                }))
            .Where(static item => item.Include.Contains("OfficeIMO.Drawing/Internal", StringComparison.OrdinalIgnoreCase)
                || item.Include.Contains("OfficeIMO.Shared/", StringComparison.OrdinalIgnoreCase))
            .Select(static item => item.Project + " -> " + item.Include)
            .ToArray();

        Assert.Empty(linkedFoundationSources);
    }

    [Fact]
    public void SharedSource_IsLimitedToDependencySpecificAdaptersAndPolyfills() {
        string sharedSourceRoot = GetRepositoryPath("OfficeIMO.SharedSource");
        string[] files = Directory.EnumerateFiles(sharedSourceRoot, "*.cs", SearchOption.AllDirectories)
            .Select(GetRepositoryRelativePath)
            .OrderBy(static path => path, StringComparer.OrdinalIgnoreCase)
            .ToArray();

        Assert.Equal(
            new[] {
                "OfficeIMO.SharedSource/Compatibility/TrimmingAttributes.cs",
                "OfficeIMO.SharedSource/Compound/OfficeCompoundDocumentDetector.cs",
                "OfficeIMO.SharedSource/Compound/OfficeCompoundFile.cs",
                "OfficeIMO.SharedSource/Compound/OfficeCompoundFileEntry.cs",
                "OfficeIMO.SharedSource/Compound/OfficeCompoundFileReader.cs",
                "OfficeIMO.SharedSource/Compound/OfficeCompoundFileReader.Inspection.cs",
                "OfficeIMO.SharedSource/Compound/OfficeCompoundFileReader.Selective.cs",
                "OfficeIMO.SharedSource/Compound/OfficeCompoundFileWriter.cs",
                "OfficeIMO.SharedSource/Compound/OfficeCompoundWriterLayout.cs",
                "OfficeIMO.SharedSource/OpenXml/OfficeOpenXmlPackagePayload.cs",
                "OfficeIMO.SharedSource/OpenXml/OfficeOpenXmlThemeColorResolver.cs"
            },
            files);
    }

    [Theory]
    [InlineData("OfficeIMO.CSV/OfficeIMO.CSV.csproj")]
    [InlineData("OfficeIMO.Email/OfficeIMO.Email.csproj")]
    [InlineData("OfficeIMO.OpenDocument/OfficeIMO.OpenDocument.csproj")]
    [InlineData("OfficeIMO.Rtf/OfficeIMO.Rtf.csproj")]
    [InlineData("OfficeIMO.Word/OfficeIMO.Word.csproj")]
    [InlineData("OfficeIMO.Excel/OfficeIMO.Excel.csproj")]
    [InlineData("OfficeIMO.Visio/OfficeIMO.Visio.csproj")]
    [InlineData("OfficeIMO.Pdf/OfficeIMO.Pdf.csproj")]
    [InlineData("OfficeIMO.PowerPoint.Pdf/OfficeIMO.PowerPoint.Pdf.csproj")]
    [InlineData("OfficeIMO.Word.Html/OfficeIMO.Word.Html.csproj")]
    [InlineData("OfficeIMO.Word.Markdown/OfficeIMO.Word.Markdown.csproj")]
    [InlineData("OfficeIMO.Zip/OfficeIMO.Zip.csproj")]
    public void SharedFoundationConsumers_ReferenceOfficeImoDrawing(string relativeProjectPath) {
        var projectPath = GetRepositoryPath(relativeProjectPath);
        Assert.True(File.Exists(projectPath), "Project file is missing: " + projectPath);

        var document = XDocument.Load(projectPath);
        var ns = document.Root?.Name.Namespace ?? XNamespace.None;

        var references = document
            .Descendants(ns + "ProjectReference")
            .Select(static e => NormalizeProjectPath((string?)e.Attribute("Include")))
            .Where(static include => include.EndsWith("OfficeIMO.Drawing/OfficeIMO.Drawing.csproj", StringComparison.OrdinalIgnoreCase))
            .ToArray();

        Assert.Single(references);
    }

    [Fact]
    public void ExcelCore_DoesNotReferenceCsvPackages() {
        string projectPath = GetRepositoryPath("OfficeIMO.Excel/OfficeIMO.Excel.csproj");
        string[] references = GetProjectReferences(projectPath);

        Assert.DoesNotContain(references, reference => reference.EndsWith("OfficeIMO.CSV/OfficeIMO.CSV.csproj", StringComparison.OrdinalIgnoreCase));
        Assert.DoesNotContain(references, reference => reference.EndsWith("OfficeIMO.Reader.Csv/OfficeIMO.Reader.Csv.csproj", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void LegacyDocRuntime_StaysFirstPartyAndDependencyFree() {
        string[] forbiddenPackageIds = [
            "NPOI",
            "OpenMcdf",
            "ExcelDataReader",
            "Microsoft.Office.Interop.Word",
            "Microsoft.Office.Interop.Excel",
            "System.Data.OleDb"
        ];
        string[] forbiddenSourceTerms = [
            "NPOI",
            "OpenMcdf",
            "ExcelDataReader",
            "Microsoft.Office.Interop",
            "Word.Application",
            "LibreOffice",
            "soffice",
            "OleDb",
            "Jet.OLEDB",
            "ACE.OLEDB"
        ];

        string projectPath = GetRepositoryPath("OfficeIMO.Word/OfficeIMO.Word.csproj");
        Assert.True(File.Exists(projectPath), "Project file is missing: " + projectPath);
        var projectDocument = XDocument.Load(projectPath);
        var ns = projectDocument.Root?.Name.Namespace ?? XNamespace.None;

        string[] packageOffenders = projectDocument
            .Descendants(ns + "PackageReference")
            .Select(static element => (string?)element.Attribute("Include") ?? string.Empty)
            .Where(include => forbiddenPackageIds.Contains(include, StringComparer.OrdinalIgnoreCase))
            .ToArray();
        Assert.Empty(packageOffenders);

        string[] sourceRoots = [
            "OfficeIMO.Word/LegacyDoc",
            "OfficeIMO.Drawing/Internal/Compound"
        ];
        string[] sourceFiles = sourceRoots
            .Select(GetRepositoryPath)
            .Where(Directory.Exists)
            .SelectMany(root => Directory.EnumerateFiles(root, "*.cs", SearchOption.AllDirectories))
            .Concat(new[] {
                GetRepositoryPath("OfficeIMO.Word/WordDocument.LoadRouting.cs")
            })
            .Where(File.Exists)
            .ToArray();

        var sourceOffenders = new List<string>();
        foreach (string sourceFile in sourceFiles) {
            string source = File.ReadAllText(sourceFile);
            foreach (string forbiddenTerm in forbiddenSourceTerms) {
                if (source.Contains(forbiddenTerm, StringComparison.OrdinalIgnoreCase)) {
                    sourceOffenders.Add(GetRepositoryRelativePath(sourceFile) + " -> " + forbiddenTerm);
                }
            }
        }

        Assert.Empty(sourceOffenders);
    }

    public static IEnumerable<object[]> PdfConversionAdapters() {
        yield return new object[] {
            "OfficeIMO.Excel.Pdf/OfficeIMO.Excel.Pdf.csproj",
            new[] {
                "OfficeIMO.Excel/OfficeIMO.Excel.csproj",
                "OfficeIMO.Pdf/OfficeIMO.Pdf.csproj"
            }
        };
        yield return new object[] {
            "OfficeIMO.Word.Pdf/OfficeIMO.Word.Pdf.csproj",
            new[] {
                "OfficeIMO.Word/OfficeIMO.Word.csproj",
                "OfficeIMO.Pdf/OfficeIMO.Pdf.csproj"
            }
        };
        yield return new object[] {
            "OfficeIMO.PowerPoint.Pdf/OfficeIMO.PowerPoint.Pdf.csproj",
            new[] {
                "OfficeIMO.PowerPoint/OfficeIMO.PowerPoint.csproj",
                "OfficeIMO.Pdf/OfficeIMO.Pdf.csproj",
                "OfficeIMO.Drawing/OfficeIMO.Drawing.csproj"
            }
        };
    }

    [Theory]
    [MemberData(nameof(PdfConversionAdapters))]
    public void PdfConversionAdapters_StayThinOverDocumentAndPdfEngines(string relativeProjectPath, string[] expectedProjectReferences) {
        var projectPath = GetRepositoryPath(relativeProjectPath);
        Assert.True(File.Exists(projectPath), "Project file is missing: " + projectPath);

        string[] projectReferences = GetProjectReferences(projectPath);
        foreach (var expectedReference in expectedProjectReferences) {
            Assert.Contains(
                projectReferences,
                reference => reference.EndsWith(NormalizeProjectPath(expectedReference), StringComparison.OrdinalIgnoreCase));
        }

        Assert.DoesNotContain(projectReferences, reference => reference.Contains("iText", StringComparison.OrdinalIgnoreCase));
        Assert.DoesNotContain(projectReferences, reference => reference.Contains("PdfSharp", StringComparison.OrdinalIgnoreCase));
    }

    [Theory]
    [InlineData("OfficeIMO.Visio/VisioPngRenderer.PngRaster.cs")]
    [InlineData("OfficeIMO.Visio/VisioPngRenderer.Encoding.cs")]
    public void RetiredPrivateRenderingBrains_AreNotRestored(string relativePath) {
        Assert.False(File.Exists(GetRepositoryPath(relativePath)), "Retired private renderer file should stay in OfficeIMO.Drawing instead: " + relativePath);
    }

    [Fact]
    public void RenderingAdapters_DoNotDeclarePrivateRasterInfrastructure() {
        string[] renderingAdapterRoots = [
            "OfficeIMO.Excel",
            "OfficeIMO.Visio",
            "OfficeIMO.PowerPoint",
            "OfficeIMO.Excel.Pdf",
            "OfficeIMO.Word.Pdf",
            "OfficeIMO.PowerPoint.Pdf"
        ];
        string[] forbiddenTypeNames = [
            "PngRaster",
            "PngEncoder",
            "PngWriter",
            "PngDecoder",
            "RgbaImage",
            "RgbaCanvas",
            "RasterImage",
            "RasterRenderTarget"
        ];
        Regex forbiddenDeclaration = new(
            @"\b(class|struct)\s+(" + string.Join("|", forbiddenTypeNames.Select(Regex.Escape)) + @")\b",
            RegexOptions.CultureInvariant);

        var offenders = new List<string>();
        foreach (string root in renderingAdapterRoots) {
            string rootPath = GetRepositoryPath(root);
            if (!Directory.Exists(rootPath)) {
                continue;
            }

            foreach (string sourceFile in Directory.EnumerateFiles(rootPath, "*.cs", SearchOption.AllDirectories)) {
                if (sourceFile.Contains($"{Path.DirectorySeparatorChar}bin{Path.DirectorySeparatorChar}", StringComparison.OrdinalIgnoreCase) ||
                    sourceFile.Contains($"{Path.DirectorySeparatorChar}obj{Path.DirectorySeparatorChar}", StringComparison.OrdinalIgnoreCase)) {
                    continue;
                }

                string source = File.ReadAllText(sourceFile);
                Match match = forbiddenDeclaration.Match(source);
                if (match.Success) {
                    offenders.Add(GetRepositoryRelativePath(sourceFile) + " declares " + match.Groups[2].Value);
                }
            }
        }

        Assert.Empty(offenders);
    }

    [Theory]
    [InlineData("OfficeIMO.Reader.Core/OfficeIMO.Reader.Core.csproj")]
    [InlineData("OfficeIMO.Reader.Json/OfficeIMO.Reader.Json.csproj")]
    [InlineData("OfficeIMO.MarkdownRenderer/OfficeIMO.MarkdownRenderer.csproj")]
    [InlineData("OfficeIMO.MarkdownRenderer.SamplePlugin/OfficeIMO.MarkdownRenderer.SamplePlugin.csproj")]
    [InlineData("OfficeIMO.GoogleWorkspace/OfficeIMO.GoogleWorkspace.csproj")]
    [InlineData("OfficeIMO.Excel.GoogleSheets/OfficeIMO.Excel.GoogleSheets.csproj")]
    [InlineData("OfficeIMO.Word.GoogleDocs/OfficeIMO.Word.GoogleDocs.csproj")]
    public void SystemTextJsonPackageReference_IsLimitedToNonInboxTargets(string relativeProjectPath) {
        var projectPath = GetRepositoryPath(relativeProjectPath);
        Assert.True(File.Exists(projectPath), "Project file is missing: " + projectPath);

        var document = XDocument.Load(projectPath);
        var ns = document.Root?.Name.Namespace ?? XNamespace.None;

        var references = document
            .Descendants(ns + "PackageReference")
            .Where(static e => string.Equals((string?)e.Attribute("Include"), "System.Text.Json", StringComparison.Ordinal))
            .ToArray();

        Assert.Single(references);

        var parentItemGroup = references[0].Parent;
        Assert.NotNull(parentItemGroup);

        var condition = (string?)parentItemGroup!.Attribute("Condition");
        Assert.False(string.IsNullOrWhiteSpace(condition));
        Assert.Contains("netstandard2.0", condition!, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("net472", condition!, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("net8.0", condition!, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("net10.0", condition!, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ReaderCore_HasNoFormatEngineProjectDependencies() {
        string projectPath = GetRepositoryPath("OfficeIMO.Reader.Core/OfficeIMO.Reader.Core.csproj");

        Assert.Empty(GetProjectReferences(projectPath));
        Assert.Equal(
            ["System.Text.Json"],
            GetPackageReferences(projectPath).Distinct(StringComparer.OrdinalIgnoreCase).ToArray());
    }

    [Fact]
    public void ReaderEmail_IsTheSingleEmailAdapterPackage() {
        string projectPath = GetRepositoryPath("OfficeIMO.Reader.Email/OfficeIMO.Reader.Email.csproj");
        string[] references = GetProjectReferences(projectPath);

        Assert.Equal(
            [
                "../OfficeIMO.Reader.Core/OfficeIMO.Reader.Core.csproj",
                "../OfficeIMO.Email/OfficeIMO.Email.csproj"
            ],
            references);
        Assert.False(File.Exists(GetRepositoryPath("OfficeIMO.Reader.EmailStore/OfficeIMO.Reader.EmailStore.csproj")));
        Assert.False(File.Exists(GetRepositoryPath("OfficeIMO.Reader.EmailAddressBook/OfficeIMO.Reader.EmailAddressBook.csproj")));
    }

    [Fact]
    public void ReaderImage_ExposesDrawingAsARuntimeDependency() {
        string projectPath = GetRepositoryPath("OfficeIMO.Reader.Image/OfficeIMO.Reader.Image.csproj");
        Assert.Equal(
            [
                "../OfficeIMO.Reader.Core/OfficeIMO.Reader.Core.csproj",
                "../OfficeIMO.Drawing/OfficeIMO.Drawing.csproj"
            ],
            GetProjectReferences(projectPath));

        XDocument document = XDocument.Load(projectPath);
        XNamespace ns = document.Root?.Name.Namespace ?? XNamespace.None;
        XElement drawingReference = Assert.Single(
            document.Descendants(ns + "ProjectReference"),
            static reference => string.Equals(
                ((string?)reference.Attribute("Include"))?.Replace('\\', '/'),
                "../OfficeIMO.Drawing/OfficeIMO.Drawing.csproj",
                StringComparison.Ordinal));
        string? privateAssets = (string?)drawingReference.Attribute("PrivateAssets") ??
            (string?)drawingReference.Element(ns + "PrivateAssets");
        Assert.True(string.IsNullOrWhiteSpace(privateAssets),
            "OfficeIMO.Drawing supplies runtime code used by Reader.Image and must remain visible in its NuGet dependency graph.");
    }

    [Fact]
    public void OfficeFormatReaderPackages_ExposeDrawingAsARuntimeDependency() {
        string[] projectNames = [
            "OfficeIMO.Reader.Word",
            "OfficeIMO.Reader.Excel",
            "OfficeIMO.Reader.PowerPoint"
        ];

        foreach (string projectName in projectNames) {
            string projectPath = GetRepositoryPath($"{projectName}/{projectName}.csproj");
            XDocument document = XDocument.Load(projectPath);
            XNamespace ns = document.Root?.Name.Namespace ?? XNamespace.None;
            XElement drawingReference = Assert.Single(
                document.Descendants(ns + "ProjectReference"),
                static reference => string.Equals(
                    ((string?)reference.Attribute("Include"))?.Replace('\\', '/'),
                    "../OfficeIMO.Drawing/OfficeIMO.Drawing.csproj",
                    StringComparison.Ordinal));
            string? privateAssets = (string?)drawingReference.Attribute("PrivateAssets") ??
                (string?)drawingReference.Element(ns + "PrivateAssets");
            Assert.True(string.IsNullOrWhiteSpace(privateAssets),
                $"OfficeIMO.Drawing supplies runtime code used by {projectName} and must remain visible in its NuGet dependency graph.");
        }
    }

    [Fact]
    public void ReaderAll_ComposesOnlyReaderPackages() {
        string projectPath = GetRepositoryPath("OfficeIMO.Reader.All/OfficeIMO.Reader.All.csproj");
        string[] references = GetProjectReferences(projectPath);

        Assert.NotEmpty(references);
        Assert.All(references, reference =>
            Assert.StartsWith("../OfficeIMO.Reader.", reference, StringComparison.Ordinal));
        Assert.Contains("../OfficeIMO.Reader.Core/OfficeIMO.Reader.Core.csproj", references, StringComparer.Ordinal);
        Assert.Contains("../OfficeIMO.Reader.Email/OfficeIMO.Reader.Email.csproj", references, StringComparer.Ordinal);
        Assert.Contains("../OfficeIMO.Reader.Word/OfficeIMO.Reader.Word.csproj", references, StringComparer.Ordinal);
        Assert.Contains("../OfficeIMO.Reader.Excel/OfficeIMO.Reader.Excel.csproj", references, StringComparer.Ordinal);
        Assert.Contains("../OfficeIMO.Reader.PowerPoint/OfficeIMO.Reader.PowerPoint.csproj", references, StringComparer.Ordinal);
        Assert.Contains("../OfficeIMO.Reader.Markdown/OfficeIMO.Reader.Markdown.csproj", references, StringComparer.Ordinal);
        Assert.Empty(GetPackageReferences(projectPath));
    }

    [Fact]
    public void Security_IsTheNeutralSingleCmsOwner() {
        string projectPath = GetRepositoryPath("OfficeIMO.Security/OfficeIMO.Security.csproj");

        Assert.Empty(GetProjectReferences(projectPath));
        Assert.Equal(["BouncyCastle.Cryptography"], GetPackageReferences(projectPath));
        Assert.False(File.Exists(GetRepositoryPath(
            "OfficeIMO.Pdf.Cryptography.Pkcs/OfficeIMO.Pdf.Cryptography.Pkcs.csproj")));
    }

    [Fact]
    public void PdfAndEmail_ConsumeSecurityDirectlyWithoutACompatibilityLayer() {
        Assert.Contains(
            "../OfficeIMO.Security/OfficeIMO.Security.csproj",
            GetProjectReferences(GetRepositoryPath("OfficeIMO.Pdf/OfficeIMO.Pdf.csproj")),
            StringComparer.Ordinal);
        Assert.Contains(
            "../OfficeIMO.Security/OfficeIMO.Security.csproj",
            GetProjectReferences(GetRepositoryPath("OfficeIMO.Email/OfficeIMO.Email.csproj")),
            StringComparer.Ordinal);
    }

    [Theory]
    [InlineData("OfficeIMO.CSV/OfficeIMO.CSV.csproj")]
    [InlineData("OfficeIMO.CSV.Tests/OfficeIMO.CSV.Tests.csproj")]
    public void NetFrameworkReferenceAssemblies_AreLimitedToNet472(string relativeProjectPath) {
        var projectPath = GetRepositoryPath(relativeProjectPath);
        Assert.True(File.Exists(projectPath), "Project file is missing: " + projectPath);

        var document = XDocument.Load(projectPath);
        var ns = document.Root?.Name.Namespace ?? XNamespace.None;

        var references = document
            .Descendants(ns + "PackageReference")
            .Where(static e => string.Equals((string?)e.Attribute("Include"), "Microsoft.NETFramework.ReferenceAssemblies", StringComparison.Ordinal))
            .ToArray();

        Assert.Single(references);

        var parentItemGroup = references[0].Parent;
        Assert.NotNull(parentItemGroup);

        var condition = (string?)parentItemGroup!.Attribute("Condition");
        Assert.False(string.IsNullOrWhiteSpace(condition));
        Assert.Contains("net472", condition!, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("net8.0", condition!, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("net10.0", condition!, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("netstandard2.0", condition!, StringComparison.OrdinalIgnoreCase);
    }

    private static string GetRepositoryRoot() {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory != null) {
            if (
                File.Exists(Path.Combine(directory.FullName, "OfficeIMO.sln")) ||
                File.Exists(Path.Combine(directory.FullName, "OfficeImo.sln"))
            ) {
                return directory.FullName;
            }

            directory = directory.Parent;
        }

        throw new DirectoryNotFoundException("Unable to locate OfficeIMO repository root from test runtime base directory.");
    }

    private static string GetRepositoryPath(string relativePath) {
        Assert.False(Path.IsPathRooted(relativePath), "Repository-relative path must not be rooted: " + relativePath);

        var repositoryRoot = Path.GetFullPath(GetRepositoryRoot());
        if (!repositoryRoot.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal)) {
            repositoryRoot += Path.DirectorySeparatorChar;
        }

        var parts = NormalizeProjectPath(relativePath)
            .Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
        var combinedPath = repositoryRoot;
        foreach (var part in parts) {
            Assert.False(Path.IsPathRooted(part), "Repository-relative path segment must not be rooted: " + relativePath);
            combinedPath = AppendRepositoryPathSegment(combinedPath, part);
        }

        combinedPath = Path.GetFullPath(combinedPath);

        Assert.True(
            combinedPath.StartsWith(repositoryRoot, StringComparison.Ordinal),
            "Repository-relative path must stay under repository root: " + relativePath);
        return combinedPath;
    }

    private static string[] EnumerateProjectFiles() =>
        Directory.EnumerateFiles(GetRepositoryRoot(), "*.csproj", SearchOption.AllDirectories)
            .Where(static path => !path.Contains($"{Path.DirectorySeparatorChar}bin{Path.DirectorySeparatorChar}", StringComparison.OrdinalIgnoreCase))
            .Where(static path => !path.Contains($"{Path.DirectorySeparatorChar}obj{Path.DirectorySeparatorChar}", StringComparison.OrdinalIgnoreCase))
            .Where(static path => !path.Contains($"{Path.DirectorySeparatorChar}Ignore{Path.DirectorySeparatorChar}", StringComparison.OrdinalIgnoreCase))
            .Where(static path => new FileInfo(path).Length > 0)
            .ToArray();

    private static IEnumerable<string> EnumerateDocumentImageRenderingSourceFiles() {
        foreach (string relativeRoot in DocumentImageRenderingRoots) {
            string root = GetRepositoryPath(relativeRoot);
            if (!Directory.Exists(root)) {
                continue;
            }

            foreach (string sourceFile in Directory.EnumerateFiles(root, "*.cs", SearchOption.AllDirectories)) {
                if (sourceFile.Contains($"{Path.DirectorySeparatorChar}bin{Path.DirectorySeparatorChar}", StringComparison.OrdinalIgnoreCase) ||
                    sourceFile.Contains($"{Path.DirectorySeparatorChar}obj{Path.DirectorySeparatorChar}", StringComparison.OrdinalIgnoreCase)) {
                    continue;
                }

                yield return sourceFile;
            }
        }
    }

    private static bool ContainsForbiddenDocumentImageRenderingSourceTerm(string source, string forbiddenTerm) {
        if (string.Equals(forbiddenTerm, "soffice", StringComparison.OrdinalIgnoreCase)) {
            return Regex.IsMatch(source, @"(?<![A-Za-z0-9_])soffice(?:\.exe)?(?![A-Za-z0-9_])", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        }

        return source.Contains(forbiddenTerm, StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsApprovedExternalCompatibilitySourceTerm(string relativePath, string forbiddenTerm) =>
        ApprovedExternalCompatibilitySourceTerms.TryGetValue(relativePath, out HashSet<string>? allowedTerms) &&
        allowedTerms.Contains(forbiddenTerm);

    private static IEnumerable<string> EnumerateDocumentImageRenderingProjectFiles() {
        foreach (string relativeRoot in DocumentImageRenderingRoots) {
            string projectPath = GetRepositoryPath(relativeRoot + "/" + relativeRoot + ".csproj");
            if (File.Exists(projectPath)) {
                yield return projectPath;
            }
        }
    }

    private static int CountJsonArrayEntries(string relativePath) {
        var path = GetRepositoryPath(relativePath);
        Assert.True(File.Exists(path), "Fixture file is missing: " + path);

        using JsonDocument document = JsonDocument.Parse(File.ReadAllText(path));
        Assert.Equal(JsonValueKind.Array, document.RootElement.ValueKind);
        return document.RootElement.GetArrayLength();
    }

    private static bool IsNonProductionProject(string projectPath) {
        string normalized = projectPath.Replace('\\', '/');
        return normalized.Contains("/OfficeIMO.Tests/", StringComparison.OrdinalIgnoreCase) ||
            normalized.Contains("/OfficeIMO.VerifyTests/", StringComparison.OrdinalIgnoreCase) ||
            normalized.Contains(".Tests/", StringComparison.OrdinalIgnoreCase) ||
            normalized.Contains(".Benchmarks", StringComparison.OrdinalIgnoreCase) ||
            normalized.Contains("/OfficeIMO.Examples/", StringComparison.OrdinalIgnoreCase);
    }

    private static string AppendRepositoryPathSegment(string basePath, string segment) =>
        basePath.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal)
            ? basePath + segment
            : basePath + Path.DirectorySeparatorChar + segment;

    private static string NormalizeProjectPath(string? path) =>
        (path ?? string.Empty).Replace('\\', '/');

    private static string GetRepositoryRelativePath(string path) {
        var repositoryRoot = Path.GetFullPath(GetRepositoryRoot());
        if (!repositoryRoot.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal)) {
            repositoryRoot += Path.DirectorySeparatorChar;
        }

        var rootUri = new Uri(repositoryRoot, UriKind.Absolute);
        var pathUri = new Uri(Path.GetFullPath(path), UriKind.Absolute);
        var relativePath = Uri.UnescapeDataString(rootUri.MakeRelativeUri(pathUri).ToString());
        Assert.False(
            relativePath == ".." || relativePath.StartsWith("../", StringComparison.Ordinal),
            "Path must stay under repository root: " + path);
        return NormalizeProjectPath(relativePath);
    }

    private static string[] GetProjectReferences(string projectPath) {
        var document = XDocument.Load(projectPath);
        var ns = document.Root?.Name.Namespace ?? XNamespace.None;

        return document
            .Descendants(ns + "ProjectReference")
            .Select(static e => NormalizeProjectPath((string?)e.Attribute("Include")))
            .Where(static include => !string.IsNullOrWhiteSpace(include))
            .ToArray();
    }

    private static string[] GetPackageReferences(string projectPath) {
        var document = XDocument.Load(projectPath);
        var ns = document.Root?.Name.Namespace ?? XNamespace.None;

        return document
            .Descendants(ns + "PackageReference")
            .Select(static e => (string?)e.Attribute("Include") ?? string.Empty)
            .Where(static include => !string.IsNullOrWhiteSpace(include))
            .ToArray();
    }

    private static bool ProjectReferencesImageSharp(string projectPath) {
        return ProjectReferencesPackages(projectPath, ["SixLabors.ImageSharp"]).Any();
    }

    private static bool ProjectReferencesSixLaborsFonts(string projectPath) {
        return ProjectReferencesPackages(projectPath, ["SixLabors.Fonts"]).Any();
    }

    private static bool SourceGrantsRetiredAggregateTestAccess(string sourcePath) {
        string source = File.ReadAllText(sourcePath);
        return Regex.IsMatch(
            source,
            @"\[\s*assembly\s*:\s*InternalsVisibleTo\s*\(\s*""OfficeIMO\.Tests""\s*\)\s*\]",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
    }

    private static IEnumerable<string> ProjectReferencesPackages(string projectPath, IReadOnlyCollection<string> packageIds) {
        var document = XDocument.Load(projectPath);
        var ns = document.Root?.Name.Namespace ?? XNamespace.None;

        return document
            .Descendants(ns + "PackageReference")
            .Select(static e => (string?)e.Attribute("Include") ?? string.Empty)
            .Where(include => packageIds.Contains(include, StringComparer.OrdinalIgnoreCase));
    }

    private static string? GetPackageReferenceVersion(string projectPath, string packageId) {
        var document = XDocument.Load(projectPath);
        var ns = document.Root?.Name.Namespace ?? XNamespace.None;

        return document
            .Descendants(ns + "PackageReference")
            .Where(element => string.Equals((string?)element.Attribute("Include"), packageId, StringComparison.OrdinalIgnoreCase))
            .Select(static element => (string?)element.Attribute("Version"))
            .SingleOrDefault();
    }
}
