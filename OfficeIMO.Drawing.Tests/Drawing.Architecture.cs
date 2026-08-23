using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public class DrawingArchitectureTests {
    private static readonly string RepositoryRoot = LocateRepositoryRoot();

    [Fact]
    public void OfficeDrawingProjectHasNoRuntimeDependencies() {
        XDocument project = LoadProject("OfficeIMO.Core", "OfficeIMO.Core.csproj");

        Assert.Empty(GetReferencedItems(project, "PackageReference"));
        Assert.Empty(GetReferencedItems(project, "ProjectReference"));
        Assert.DoesNotContain("System.Drawing", ReadProjectSource("OfficeIMO.Core"));
    }

    [Fact]
    public void LifecycleContractsAreOwnedByDependencyFreeCore() {
        Assert.Same(typeof(OfficeColor).Assembly, typeof(DocumentAccessMode).Assembly);
        Assert.Same(typeof(OfficeColor).Assembly, typeof(DocumentPersistenceMode).Assembly);
        Assert.Same(typeof(OfficeColor).Assembly, typeof(DocumentCreateOptions).Assembly);
        Assert.Same(typeof(OfficeColor).Assembly, typeof(DocumentLoadOptions).Assembly);
        Assert.True(File.Exists(Path.Combine(RepositoryRoot, "OfficeIMO.Core", "OfficeIMO.Core.csproj")));
    }

    [Fact]
    public void PrimaryImageExportOwnersReferenceOfficeDrawing() {
        string[] projectFolders = {
            "OfficeIMO.Excel",
            "OfficeIMO.Excel.Pdf",
            "OfficeIMO.Word",
            "OfficeIMO.Visio",
            "OfficeIMO.Pdf",
            "OfficeIMO.Markdown.Pdf",
            "OfficeIMO.PowerPoint.Pdf",
            "OfficeIMO.Word.Pdf"
        };

        foreach (string projectFolder in projectFolders) {
            XDocument project = LoadProject(projectFolder, projectFolder + ".csproj");
            Assert.Contains(
                GetReferencedItems(project, "ProjectReference"),
                reference => reference.Replace('\\', '/').EndsWith("/OfficeIMO.Core/OfficeIMO.Core.csproj", StringComparison.OrdinalIgnoreCase));
        }
    }

    [Fact]
    public void ProjectsUsingOfficeDrawingReferenceOfficeDrawingDirectly() {
        foreach (string projectPath in Directory.GetFiles(RepositoryRoot, "OfficeIMO.*.csproj", SearchOption.AllDirectories)) {
            if (IsNonProductionProject(projectPath) || IsOfficeDrawingProject(projectPath)) {
                continue;
            }

            string projectFolder = Path.GetDirectoryName(projectPath)!;
            if (!ProjectSourceUsesOfficeDrawing(projectFolder)) {
                continue;
            }

            XDocument project = XDocument.Load(projectPath);
            Assert.Contains(
                GetReferencedItems(project, "ProjectReference"),
                reference => reference.Replace('\\', '/').EndsWith("/OfficeIMO.Core/OfficeIMO.Core.csproj", StringComparison.OrdinalIgnoreCase));
        }
    }

    [Fact]
    public void ProductionProjectsDoNotReferenceThirdPartyImageRenderingPackages() {
        string[] bannedPackages = {
            "Aspose.Cells",
            "Aspose.PDF",
            "Aspose.Slides",
            "Aspose.Words",
            "Docnet.Core",
            "GemBox.Document",
            "GemBox.Pdf",
            "GemBox.Presentation",
            "GemBox.Spreadsheet",
            "Magick.NET-Q8-AnyCPU",
            "Magick.NET-Q16-AnyCPU",
            "Microsoft.Office.Interop.Excel",
            "Microsoft.Office.Interop.PowerPoint",
            "Microsoft.Office.Interop.Word",
            "Microsoft.Playwright",
            "PdfiumViewer",
            "PDFiumSharp",
            "PDFtoImage",
            "PuppeteerSharp",
            "Selenium.WebDriver",
            "System.Drawing.Common",
            "SixLabors.ImageSharp",
            "SixLabors.Fonts",
            "SkiaSharp",
            "SkiaSharp.Extended",
            "SkiaSharp.Views",
            "Spire.Doc",
            "Spire.PDF",
            "Spire.Presentation",
            "Spire.XLS",
            "Svg.Skia",
            "Svg",
            "Syncfusion.DocIO.Net.Core",
            "Syncfusion.EJ2.PdfViewer",
            "Syncfusion.Pdf.Net.Core",
            "Syncfusion.PdfToImageConverter.Net",
            "Syncfusion.Presentation.Net.Core",
            "Syncfusion.XlsIO.Net.Core"
        };

        foreach (string projectPath in Directory.GetFiles(RepositoryRoot, "OfficeIMO.*.csproj", SearchOption.AllDirectories)) {
            if (IsNonProductionProject(projectPath) || IsIsolatedExternalDrawingAdapter(projectPath)) {
                continue;
            }

            XDocument project = XDocument.Load(projectPath);
            IReadOnlyList<string> packageReferences = GetReferencedItems(project, "PackageReference");
            foreach (string package in bannedPackages) {
                Assert.DoesNotContain(
                    packageReferences,
                    reference => string.Equals(reference, package, StringComparison.OrdinalIgnoreCase));
            }
        }
    }

    [Fact]
    public void SixLaborsFontAdapterKeepsItsDependencyOutsideTheCoreRenderingOwners() {
        XDocument project = LoadProject("OfficeIMO.Drawing.SixLabors", "OfficeIMO.Drawing.SixLabors.csproj");

        Assert.Equal(new[] { "SixLabors.Fonts" }, GetReferencedItems(project, "PackageReference"));
        Assert.Contains(
            GetReferencedItems(project, "ProjectReference"),
            reference => reference.Replace('\\', '/').EndsWith("/OfficeIMO.Core/OfficeIMO.Core.csproj", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void ImageRenderingOwnerCodeDoesNotUseSystemDrawingOrExternalImageNamespaces() {
        string[] bannedSourceTokens = {
            "Aspose.",
            "GemBox.",
            "ImageMagick.",
            "Microsoft.Office.Interop",
            "Microsoft.Playwright",
            "Microsoft.Web.WebView2",
            "Pdfium",
            "PDFium",
            "pdftocairo",
            "pdftoppm",
            "Poppler",
            "PuppeteerSharp",
            "Selenium.",
            "using System.Drawing",
            "System.Drawing.",
            "using SixLabors.",
            "SixLabors.",
            "using SkiaSharp",
            "SkiaSharp.",
            "using ImageMagick",
            "Spire.",
            "Syncfusion."
        };

        foreach (string filePath in EnumerateImageRenderingOwnerSource()) {
            string source = File.ReadAllText(filePath);
            foreach (string token in bannedSourceTokens) {
                Assert.DoesNotContain(token, source, StringComparison.Ordinal);
            }
        }
    }

    [Fact]
    public void VisioPngRendererUsesSharedDrawingRasterStack() {
        string rasterAdapter = File.ReadAllText(Path.Combine(RepositoryRoot, "OfficeIMO.Visio", "VisioPngRenderer.RasterCanvas.cs"));
        string renderer = File.ReadAllText(Path.Combine(RepositoryRoot, "OfficeIMO.Visio", "VisioPngRenderer.cs"));

        Assert.Contains("OfficeRasterRenderTarget", rasterAdapter, StringComparison.Ordinal);
        Assert.Contains("OfficeRasterCanvas", rasterAdapter, StringComparison.Ordinal);
        Assert.Contains("OfficeTextBlockRenderer.DrawRasterTextBox", rasterAdapter, StringComparison.Ordinal);
        Assert.Contains("OfficeRasterImage RenderRaster", renderer, StringComparison.Ordinal);
        Assert.Contains("OfficeRasterImage.FromRgba32", renderer, StringComparison.Ordinal);
        Assert.Contains("OfficeRasterImageEncoder.Encode(", renderer, StringComparison.Ordinal);
        Assert.Contains("DpiX = options.PixelsPerInch", renderer, StringComparison.Ordinal);
        Assert.Contains("DpiY = options.PixelsPerInch", renderer, StringComparison.Ordinal);
    }

    [Fact]
    public void ExcelImageExportOptionsReuseSharedOfficeImageExportOptions() {
        Assert.True(typeof(OfficeImageExportOptions).IsAssignableFrom(typeof(ExcelImageExportOptions)));

        var options = new ExcelImageExportOptions {
            Scale = 2D,
            BackgroundColor = OfficeColor.Transparent
        };

        Assert.Equal(2D, options.Scale);
        Assert.Equal(OfficeColor.Transparent, options.BackgroundColor);
        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeImageExportOptions.ValidateScale(0D, "options"));
    }

    [Fact]
    public void WordImageExportOptionsReuseSharedOfficeImageExportOptions() {
        Assert.True(typeof(OfficeImageExportOptions).IsAssignableFrom(typeof(WordImageExportOptions)));

        var options = new WordImageExportOptions {
            Scale = 2D,
            BackgroundColor = OfficeColor.Transparent
        };

        Assert.Equal(2D, options.Scale);
        Assert.Equal(OfficeColor.Transparent, options.BackgroundColor);
        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeImageExportOptions.ValidateScale(0D, "options"));
    }

    private static XDocument LoadProject(string folder, string fileName) =>
        XDocument.Load(Path.Combine(RepositoryRoot, folder, fileName));

    private static IReadOnlyList<string> GetReferencedItems(XDocument project, string elementName) =>
        project
            .Descendants()
            .Where(element => string.Equals(element.Name.LocalName, elementName, StringComparison.Ordinal))
            .Select(element => (string?)element.Attribute("Include"))
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Select(value => value!)
            .ToArray();

    private static string ReadProjectSource(string projectFolder) =>
        string.Join(
            Environment.NewLine,
            Directory.GetFiles(Path.Combine(RepositoryRoot, projectFolder), "*.cs", SearchOption.AllDirectories)
                .Select(File.ReadAllText));

    private static IEnumerable<string> EnumerateImageRenderingOwnerSource() {
        string[] projectFolders = {
            "OfficeIMO.Core",
            "OfficeIMO.Excel",
            "OfficeIMO.Word",
            "OfficeIMO.Visio",
            "OfficeIMO.Pdf",
            "OfficeIMO.Markdown.Pdf",
            "OfficeIMO.PowerPoint.Pdf",
            "OfficeIMO.Word.Pdf"
        };

        foreach (string projectFolder in projectFolders) {
            foreach (string filePath in Directory.GetFiles(Path.Combine(RepositoryRoot, projectFolder), "*.cs", SearchOption.AllDirectories)) {
                string normalized = filePath.Replace('\\', '/');
                if (!normalized.Contains("/obj/", StringComparison.OrdinalIgnoreCase) &&
                    !normalized.EndsWith("/Properties/AssemblyInfo.cs", StringComparison.OrdinalIgnoreCase)) {
                    yield return filePath;
                }
            }
        }
    }

    private static bool IsNonProductionProject(string projectPath) {
        string normalized = projectPath.Replace('\\', '/');
        return normalized.Contains("/OfficeIMO.Tests/", StringComparison.OrdinalIgnoreCase) ||
            normalized.Contains("/OfficeIMO.VerifyTests/", StringComparison.OrdinalIgnoreCase) ||
            normalized.Contains(".Tests/", StringComparison.OrdinalIgnoreCase) ||
            normalized.Contains(".Benchmarks", StringComparison.OrdinalIgnoreCase) ||
            normalized.Contains("/OfficeIMO.Examples/", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsOfficeDrawingProject(string projectPath) =>
        string.Equals(Path.GetFileNameWithoutExtension(projectPath), "OfficeIMO.Core", StringComparison.OrdinalIgnoreCase);

    private static bool IsIsolatedExternalDrawingAdapter(string projectPath) =>
        string.Equals(
            Path.GetFileNameWithoutExtension(projectPath),
            "OfficeIMO.Drawing.SixLabors",
            StringComparison.OrdinalIgnoreCase);

    private static bool ProjectSourceUsesOfficeDrawing(string projectFolder) {
        foreach (string filePath in Directory.GetFiles(projectFolder, "*.cs", SearchOption.AllDirectories)) {
            string normalized = filePath.Replace('\\', '/');
            if (normalized.Contains("/obj/", StringComparison.OrdinalIgnoreCase)) {
                continue;
            }

            string source = File.ReadAllText(filePath);
            if (source.Contains("using OfficeIMO.Drawing", StringComparison.Ordinal) ||
                source.Contains("OfficeIMO.Drawing.", StringComparison.Ordinal) ||
                source.Contains("OfficeIMO.Drawing;", StringComparison.Ordinal)) {
                return true;
            }
        }

        return false;
    }

    private static string LocateRepositoryRoot() {
        DirectoryInfo? directory = new(AppContext.BaseDirectory);
        while (directory != null) {
            if (File.Exists(Path.Combine(directory.FullName, "OfficeIMO.sln")) ||
                Directory.Exists(Path.Combine(directory.FullName, "OfficeIMO.Core"))) {
                return directory.FullName;
            }

            directory = directory.Parent;
        }

        throw new InvalidOperationException("Unable to locate the OfficeIMO repository root.");
    }
}
