using OfficeIMO.Excel;
using OfficeIMO.Excel.Html;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Html;
using OfficeIMO.OneNote;
using OfficeIMO.OneNote.Html;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Html;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Html;
using OfficeIMO.Rtf;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using System.Reflection;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Html {
    [Fact]
    public void HtmlReadmeQuickStarts_UseCompiledPreparedDocumentApis() {
        HtmlConversionDocument source = HtmlConversionDocument.Parse("<h1>Hello</h1><p>Body</p>");

        string markdown = source.ToMarkdown();
        MarkdownDoc markdownDocument = source.ToMarkdownDocument();
        RtfDocument rtfDocument = source.ToRtfDocument();
        using WordDocument wordDocument = source.ToWordDocumentResult().RequireValue();

        Assert.Contains("Hello", markdown, StringComparison.Ordinal);
        Assert.NotNull(markdownDocument);
        Assert.Contains("Hello", rtfDocument.ToRtf(), StringComparison.Ordinal);
        Assert.NotEmpty(wordDocument.Paragraphs);

        using ExcelDocument workbook = ExcelDocument.Create(new MemoryStream());
        workbook.AddWorksheet("Data").CellValue(1, 1, "value");
        HtmlConversionDocument excelHtml = HtmlConversionDocument.Parse(workbook.ToHtml());
        using ExcelDocument importedWorkbook = excelHtml.ToExcelDocumentResult().RequireValue();
        Assert.Single(importedWorkbook.Sheets);

        using PowerPointPresentation presentation = PowerPointPresentation.Create(new MemoryStream());
        presentation.AddSlide().AddTextBoxPoints("Hello", 20, 20, 200, 40);
        HtmlConversionDocument powerPointHtml = HtmlConversionDocument.Parse(presentation.ToHtml());
        using PowerPointPresentation importedPresentation = powerPointHtml
            .ToPowerPointPresentationResult()
            .RequireValue();
        Assert.Single(importedPresentation.Slides);

        OneNoteSection oneNoteSection = source.ToOneNoteSectionResult().RequireValue();
        HtmlTextConversionResult oneNoteHtml = oneNoteSection.ToHtmlDocumentResult();
        Assert.Contains("Hello", oneNoteHtml.RequireValue(), StringComparison.Ordinal);

        OfficeIMO.Pdf.PdfDocumentConversionResult pdf = source.ToPdfDocumentResult();
        OfficeImageExportResult image = source.ExportImage(OfficeImageExportFormat.Png);
        Assert.NotEmpty(pdf.ToBytes());
        Assert.NotEmpty(image.Bytes);

        var reader = new OfficeDocumentReaderBuilder().AddHtmlHandler().Build();
        using var htmlStream = new MemoryStream(Encoding.UTF8.GetBytes("<h1>Hello</h1><p>Body</p>"));
        OfficeDocumentReadResult readerDocument = reader.ReadDocument(htmlStream, "readme.html");
        Assert.Contains("Hello", readerDocument.Markdown ?? string.Empty, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlTargetCapabilityContracts_AdvertisePublicCompiledApis() {
        var publicApis = new Dictionary<HtmlConversionTarget, ApiSurface> {
            [HtmlConversionTarget.Word] = new(
                typeof(WordHtmlConverterExtensions), "ToWordDocument", "ToWordDocumentResult",
                typeof(WordHtmlConverterExtensions), "ToHtml", "ToHtmlResult"),
            [HtmlConversionTarget.Excel] = new(
                typeof(HtmlExcelConverterExtensions), "ToExcelDocument", "ToExcelDocumentResult",
                typeof(ExcelHtmlConverterExtensions), "ToHtml", "ToHtmlResult"),
            [HtmlConversionTarget.PowerPoint] = new(
                typeof(HtmlPowerPointConverterExtensions), "ToPowerPointPresentation", "ToPowerPointPresentationResult",
                typeof(PowerPointHtmlConverterExtensions), "ToHtml", "ToHtmlResult"),
            [HtmlConversionTarget.OneNote] = new(
                typeof(HtmlOneNoteConverterExtensions), "ToOneNoteSection", "ToOneNoteSectionResult",
                typeof(OneNoteHtmlConverterExtensions), "ToHtmlDocument", "ToHtmlDocumentResult"),
            [HtmlConversionTarget.Markdown] = new(
                typeof(HtmlMarkdownConverterExtensions), "ToMarkdownDocument", "ToMarkdownDocumentResult",
                typeof(MarkdownDoc), "ToHtmlDocument", null),
            [HtmlConversionTarget.Rtf] = new(
                typeof(HtmlRtfConverterExtensions), "ToRtfDocument", "ToRtfDocumentResult",
                typeof(HtmlRtfConverterExtensions), "ToHtml", "ToHtmlResult"),
            [HtmlConversionTarget.Pdf] = new(
                typeof(HtmlPdfConverterExtensions), "ToPdfDocument", "ToPdfDocumentResult", null, null, null),
            [HtmlConversionTarget.Image] = new(
                typeof(HtmlImageExportExtensions), "ToPng", "ExportImage", null, null, null),
            [HtmlConversionTarget.Reader] = new(
                typeof(OfficeDocumentReader), "ReadDocument", "ReadDocument", null, null, null)
        };

        Assert.Equal(HtmlTargetCapabilityContracts.All.Count, publicApis.Count);
        foreach (HtmlTargetCapabilityContract contract in HtmlTargetCapabilityContracts.All) {
            ApiSurface surface = publicApis[contract.Target];
            AssertPublicMethod(surface.ImportOwner, surface.ImportMethod);
            AssertPublicMethod(surface.ResultOwner, surface.ResultMethod);
            Assert.Equal(surface.ExportOwner != null, contract.SupportsReverseHtml);
            if (surface.ExportOwner != null) {
                AssertPublicMethod(surface.ExportOwner, surface.ExportMethod!);
                if (surface.ExportResultMethod != null) {
                    AssertPublicMethod(surface.ExportOwner, surface.ExportResultMethod);
                }
            }

            Assert.Contains(surface.ImportMethod, contract.ImportEntryPoint, StringComparison.Ordinal);
            if (surface.ExportMethod != null) {
                Assert.Contains(surface.ExportMethod, contract.ExportEntryPoint!, StringComparison.Ordinal);
            }
        }
    }

    [Fact]
    public void HtmlReadmes_DoNotAdvertiseRemovedStringOrResultApis() {
        string repositoryRoot = FindRepositoryRoot();
        string[] readmePaths = {
            "OfficeIMO.Html/README.md",
            "OfficeIMO.Word.Html/README.md",
            "OfficeIMO.Excel.Html/README.md",
            "OfficeIMO.PowerPoint.Html/README.md",
            "OfficeIMO.OneNote.Html/README.md",
            "OfficeIMO.Markdown.Html/README.md",
            "OfficeIMO.Reader.Html/README.md"
        };

        foreach (string relativePath in readmePaths) {
            string readme = File.ReadAllText(Path.Combine(repositoryRoot, relativePath));
            Assert.DoesNotContain("GetArtifactOrThrow", readme, StringComparison.Ordinal);
            Assert.DoesNotContain("html.ToExcelDocument", readme, StringComparison.Ordinal);
            Assert.DoesNotContain("html.ToPowerPointPresentation", readme, StringComparison.Ordinal);
            Assert.DoesNotContain("html.ToMarkdown", readme, StringComparison.Ordinal);
            Assert.DoesNotContain("\".ToMarkdown", readme, StringComparison.Ordinal);
            Assert.DoesNotContain("\".ToRtfDocument", readme, StringComparison.Ordinal);
        }

        string excelReadme = File.ReadAllText(Path.Combine(repositoryRoot, "OfficeIMO.Excel.Html/README.md"));
        string powerPointReadme = File.ReadAllText(Path.Combine(repositoryRoot, "OfficeIMO.PowerPoint.Html/README.md"));
        Assert.Contains("HtmlConversionDocument.Parse(html)", excelReadme, StringComparison.Ordinal);
        Assert.Contains("HtmlConversionDocument.Parse(html)", powerPointReadme, StringComparison.Ordinal);
        Assert.Contains("HtmlConversionDocument.Load", excelReadme, StringComparison.Ordinal);
        Assert.Contains("HtmlConversionDocument.Load", powerPointReadme, StringComparison.Ordinal);
        Assert.Contains("result.RequireValue()", excelReadme, StringComparison.Ordinal);
        Assert.Contains("result.RequireValue()", powerPointReadme, StringComparison.Ordinal);
    }

    private static string FindRepositoryRoot() {
        for (DirectoryInfo? directory = new DirectoryInfo(AppContext.BaseDirectory); directory != null; directory = directory.Parent) {
            if (File.Exists(Path.Combine(directory.FullName, "Directory.Build.props"))
                && Directory.Exists(Path.Combine(directory.FullName, "OfficeIMO.Html"))) {
                return directory.FullName;
            }
        }

        throw new DirectoryNotFoundException("Could not locate the OfficeIMO repository root.");
    }

    private static void AssertPublicMethod(Type owner, string methodName) {
        Assert.Contains(owner.GetMethods(BindingFlags.Public | BindingFlags.Static | BindingFlags.Instance),
            method => string.Equals(method.Name, methodName, StringComparison.Ordinal));
    }

    private sealed record ApiSurface(
        Type ImportOwner,
        string ImportMethod,
        string ResultMethod,
        Type? ExportOwner,
        string? ExportMethod,
        string? ExportResultMethod) {
        public Type ResultOwner => ImportOwner;
    }
}
