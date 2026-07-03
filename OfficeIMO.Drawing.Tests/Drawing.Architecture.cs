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
    public void OfficeDrawingProjectRemainsDependencyFree() {
        XDocument project = LoadProject("OfficeIMO.Drawing", "OfficeIMO.Drawing.csproj");

        Assert.Empty(GetReferencedItems(project, "PackageReference"));
        Assert.Empty(GetReferencedItems(project, "ProjectReference"));
        Assert.DoesNotContain("System.Drawing", ReadProjectSource("OfficeIMO.Drawing"));
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
                reference => reference.Replace('\\', '/').EndsWith("/OfficeIMO.Drawing/OfficeIMO.Drawing.csproj", StringComparison.OrdinalIgnoreCase));
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
                reference => reference.Replace('\\', '/').EndsWith("/OfficeIMO.Drawing/OfficeIMO.Drawing.csproj", StringComparison.OrdinalIgnoreCase));
        }
    }

    [Fact]
    public void ProductionProjectsDoNotReferenceThirdPartyImageRenderingPackages() {
        string[] bannedPackages = {
            "System.Drawing.Common",
            "SixLabors.ImageSharp",
            "SkiaSharp",
            "Magick.NET-Q8-AnyCPU",
            "Magick.NET-Q16-AnyCPU",
            "Svg.Skia",
            "Svg"
        };

        foreach (string projectPath in Directory.GetFiles(RepositoryRoot, "OfficeIMO.*.csproj", SearchOption.AllDirectories)) {
            if (IsNonProductionProject(projectPath)) {
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
    public void ImageRenderingOwnerCodeDoesNotUseSystemDrawingOrExternalImageNamespaces() {
        string[] bannedSourceTokens = {
            "using System.Drawing",
            "System.Drawing.",
            "using SixLabors.",
            "SixLabors.",
            "using SkiaSharp",
            "SkiaSharp.",
            "using ImageMagick",
            "ImageMagick."
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
        Assert.Contains("OfficePngWriter.EncodeRgba", renderer, StringComparison.Ordinal);
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
            "OfficeIMO.Drawing",
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
                if (!filePath.Replace('\\', '/').Contains("/obj/", StringComparison.OrdinalIgnoreCase)) {
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
        string.Equals(Path.GetFileNameWithoutExtension(projectPath), "OfficeIMO.Drawing", StringComparison.OrdinalIgnoreCase);

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
                Directory.Exists(Path.Combine(directory.FullName, "OfficeIMO.Drawing"))) {
                return directory.FullName;
            }

            directory = directory.Parent;
        }

        throw new InvalidOperationException("Unable to locate the OfficeIMO repository root.");
    }
}
