using System.Text;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Markdown.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfCore = OfficeIMO.Pdf;
using WordPdf = OfficeIMO.Word.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfConversionTypographyTests {
    private const string FamilyName = "OfficeIMO Converter Unicode";

    [Theory]
    [InlineData("word-to-pdf")]
    [InlineData("excel-to-pdf")]
    [InlineData("markdown-to-pdf")]
    [InlineData("html-to-pdf")]
    [InlineData("powerpoint-to-pdf")]
    public void ConversionAdapters_PreserveEmbeddedMultilingualTextWithToUnicodeMaps(string conversionPath) {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        byte[] pdf;
        try {
            pdf = conversionPath switch {
                "word-to-pdf" => CreateWordPdf(fontPath),
                "excel-to-pdf" => CreateExcelPdf(fontPath),
                "markdown-to-pdf" => CreateMarkdownPdf(fontPath),
                "html-to-pdf" => CreateHtmlPdf(fontPath),
                "powerpoint-to-pdf" => CreatePowerPointPdf(fontPath),
                _ => throw new ArgumentOutOfRangeException(nameof(conversionPath), conversionPath, null)
            };
        } catch (ArgumentException exception) when (IsMissingEmbeddedGlyphFailure(exception)) {
            return;
        }

        string raw = Encoding.ASCII.GetString(pdf);
        string text = PdfCore.PdfReadDocument.Open(pdf).ExtractText();

        Assert.Contains("/Subtype /Type0", raw, StringComparison.Ordinal);
        Assert.Contains("/Encoding /Identity-H", raw, StringComparison.Ordinal);
        Assert.Contains("/ToUnicode", raw, StringComparison.Ordinal);
        Assert.Contains("/FontFile2", raw, StringComparison.Ordinal);
        Assert.Contains("Zażółć gęślą jaźń", text, StringComparison.Ordinal);
        Assert.Contains("Ελλάδα", text, StringComparison.Ordinal);
        Assert.Contains("Київ", text, StringComparison.Ordinal);

        WriteReviewArtifact("multilingual-" + conversionPath + ".pdf", pdf);
    }

    [Fact]
    public void MarkdownConversion_EmbeddedFontSubsetOutputIsDeterministicAcrossRuns() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        byte[] first;
        byte[] second;
        try {
            first = CreateMarkdownPdf(fontPath);
            second = CreateMarkdownPdf(fontPath);
        } catch (ArgumentException exception) when (IsMissingEmbeddedGlyphFailure(exception)) {
            return;
        }

        Assert.Equal(first, second);
    }

    [Fact]
    public void MarkdownConversion_ReportsMissingEmbeddedGlyphThroughPreflightExceptionBeforeThrowing() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        string unsupportedScalar = char.ConvertFromUtf32(0x10FFFF);
        var options = new MarkdownPdfSaveOptions {
            ApplyDefaultTheme = false,
            PdfOptions = CreatePdfOptions(fontPath)
        };

        ArgumentException exception = Assert.ThrowsAny<ArgumentException>(() => OfficeIMO.Markdown.MarkdownReader.Parse("# Missing Glyph\n\nUnsupported " + unsupportedScalar).ToPdf(options));

        Assert.True(IsMissingEmbeddedGlyphFailure(exception));
        Assert.Contains((string?)exception.Data["code"], new[] { "missing-embedded-font-glyph", "missing-embedded-font-fallback-glyph", "unsupported-text-glyph" });
        Assert.Equal("U+10FFFF", exception.Data["codePoint"]);
    }

    [Theory]
    [InlineData("word-to-pdf", "OfficeIMO.Word.Pdf")]
    [InlineData("excel-to-pdf", "OfficeIMO.Excel.Pdf")]
    [InlineData("markdown-to-pdf", "OfficeIMO.Markdown.Pdf")]
    [InlineData("powerpoint-to-pdf", "OfficeIMO.PowerPoint.Pdf")]
    public void ConversionAdapters_ReportOpenTypeFeatureWarningsForConfiguredEmbeddedFonts(string conversionPath, string expectedConverter) {
        string? fontPath = PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
        if (fontPath == null) {
            return;
        }

        PdfCore.PdfConversionReport report = conversionPath switch {
            "word-to-pdf" => CreateWordOpenTypeReport(fontPath),
            "excel-to-pdf" => CreateExcelOpenTypeReport(fontPath),
            "markdown-to-pdf" => CreateMarkdownOpenTypeReport(fontPath),
            "powerpoint-to-pdf" => CreatePowerPointOpenTypeReport(fontPath),
            _ => throw new ArgumentOutOfRangeException(nameof(conversionPath), conversionPath, null)
        };

        PdfCore.PdfConversionWarning warning = Assert.Single(report.Warnings.Where(item =>
            item.Code == "unsupported-font-ligature-substitution" &&
            item.Converter == expectedConverter).Take(1));
        Assert.Equal(PdfCore.PdfConversionWarningSeverity.Warning, warning.Severity);
        Assert.Equal("OpenType GSUB ligature", warning.Details["script"]);
        Assert.Equal("U+0066", warning.Details["codePoint"]);

        PdfCore.PdfConversionWarning markWarning = Assert.Single(report.Warnings.Where(item =>
            item.Code == "unsupported-mark-positioning-or-joiner-shaping" &&
            item.Converter == expectedConverter).Take(1));
        Assert.Equal(PdfCore.PdfConversionWarningSeverity.Warning, markWarning.Severity);
        Assert.Equal("combining-mark-or-joiner", markWarning.Details["script"]);
        Assert.Equal("U+0301", markWarning.Details["codePoint"]);
    }

    [Theory]
    [InlineData("word-to-pdf", "OfficeIMO.Word.Pdf")]
    [InlineData("excel-to-pdf", "OfficeIMO.Excel.Pdf")]
    [InlineData("markdown-to-pdf", "OfficeIMO.Markdown.Pdf")]
    [InlineData("powerpoint-to-pdf", "OfficeIMO.PowerPoint.Pdf")]
    public void ConversionAdapters_ReportComplexScriptWarningsForGeneratedText(string conversionPath, string expectedConverter) {
        PdfCore.PdfConversionReport report = conversionPath switch {
            "word-to-pdf" => CreateWordComplexScriptReport(allowMissingGlyphFailure: true),
            "excel-to-pdf" => CreateExcelComplexScriptReport(allowMissingGlyphFailure: true),
            "markdown-to-pdf" => CreateMarkdownComplexScriptReport(allowMissingGlyphFailure: true),
            "powerpoint-to-pdf" => CreatePowerPointComplexScriptReport(allowMissingGlyphFailure: true),
            _ => throw new ArgumentOutOfRangeException(nameof(conversionPath), conversionPath, null)
        };

        PdfCore.PdfConversionWarning bidiWarning = Assert.Single(report.Warnings.Where(item =>
            item.Code == "unsupported-bidirectional-text-layout" &&
            item.Converter == expectedConverter).Take(1));
        Assert.Equal(PdfCore.PdfConversionWarningSeverity.Warning, bidiWarning.Severity);
        Assert.Equal("right-to-left", bidiWarning.Details["script"]);
        Assert.Equal("U+0645", bidiWarning.Details["codePoint"]);

        PdfCore.PdfConversionWarning shapingWarning = Assert.Single(report.Warnings.Where(item =>
            item.Code == "unsupported-complex-script-shaping" &&
            item.Converter == expectedConverter).Take(1));
        Assert.Equal(PdfCore.PdfConversionWarningSeverity.Warning, shapingWarning.Severity);
        Assert.Equal("Arabic", shapingWarning.Details["script"]);
        Assert.Equal("U+0645", shapingWarning.Details["codePoint"]);
    }

    [Theory]
    [InlineData("word-to-pdf", "OfficeIMO.Word.Pdf")]
    [InlineData("excel-to-pdf", "OfficeIMO.Excel.Pdf")]
    [InlineData("markdown-to-pdf", "OfficeIMO.Markdown.Pdf")]
    [InlineData("powerpoint-to-pdf", "OfficeIMO.PowerPoint.Pdf")]
    public void ConversionAdapters_ReportScriptSpecificLineBreakingWarningsForGeneratedText(string conversionPath, string expectedConverter) {
        const string thaiText = "ภาษาไทย";
        PdfCore.PdfConversionReport report = conversionPath switch {
            "word-to-pdf" => CreateWordComplexScriptReport(thaiText, allowMissingGlyphFailure: true),
            "excel-to-pdf" => CreateExcelComplexScriptReport(thaiText, allowMissingGlyphFailure: true),
            "markdown-to-pdf" => CreateMarkdownComplexScriptReport(thaiText, allowMissingGlyphFailure: true),
            "powerpoint-to-pdf" => CreatePowerPointComplexScriptReport(thaiText, allowMissingGlyphFailure: true),
            _ => throw new ArgumentOutOfRangeException(nameof(conversionPath), conversionPath, null)
        };

        PdfCore.PdfConversionWarning warning = Assert.Single(report.Warnings.Where(item =>
            item.Code == "unsupported-script-specific-line-breaking" &&
            item.Converter == expectedConverter).Take(1));
        Assert.Equal(PdfCore.PdfConversionWarningSeverity.Warning, warning.Severity);
        Assert.Equal("Thai", warning.Details["script"]);
        Assert.Equal("U+0E20", warning.Details["codePoint"]);
    }

    private static PdfCore.PdfOptions CreatePdfOptions(string fontPath) =>
        new PdfCore.PdfOptions {
                CompressContentStreams = false,
                CompressEmbeddedFonts = false,
                PageWidth = 520,
                PageHeight = 420,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            }
            .UseFontFamily(FamilyName, fontPath);

    private static byte[] CreateWordPdf(string fontPath) {
        string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO.Pdf.Typography", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            string documentPath = Path.Combine(directory, "multilingual.docx");
            using WordDocument document = WordDocument.Create(documentPath);
            document.AddParagraph("Converter typography report");
            document.AddParagraph("Zażółć gęślą jaźń");
            WordTable table = document.AddTable(3, 2);
            table.Rows[0].Cells[0].Paragraphs[0].Text = "Region";
            table.Rows[0].Cells[1].Paragraphs[0].Text = "Signal";
            table.Rows[1].Cells[0].Paragraphs[0].Text = "Ελλάδα";
            table.Rows[1].Cells[1].Paragraphs[0].Text = "Athens";
            table.Rows[2].Cells[0].Paragraphs[0].Text = "Україна";
            table.Rows[2].Cells[1].Paragraphs[0].Text = "Київ";
            document.Save();

            return document.ToPdf(new WordPdf.PdfSaveOptions {
                PdfOptions = CreatePdfOptions(fontPath),
                IncludePageNumbers = false
            });
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    private static byte[] CreateExcelPdf(string fontPath) {
        string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO.Pdf.Typography", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            string workbookPath = Path.Combine(directory, "multilingual.xlsx");
            using ExcelDocument document = ExcelDocument.Create(workbookPath, "Report");
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "Converter typography report");
            sheet.Cell(2, 1, "Zażółć gęślą jaźń");
            sheet.Cell(3, 1, "Ελλάδα");
            sheet.Cell(4, 1, "Київ");
            document.Save();

            var options = new ExcelPdfSaveOptions {
                PdfOptions = CreatePdfOptions(fontPath),
                IncludeSheetHeadings = false
            };
            PdfCore.PdfDocumentConversionResult result = document.ToPdfDocumentResult(options);
            byte[] pdf = result.ToBytes();
            Assert.False(result.HasWarnings);
            return pdf;
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    private static byte[] CreateMarkdownPdf(string fontPath) {
        var options = new MarkdownPdfSaveOptions {
            ApplyDefaultTheme = false,
            PdfOptions = CreatePdfOptions(fontPath)
        };
        PdfCore.PdfDocumentConversionResult result = OfficeIMO.Markdown.MarkdownReader.Parse("""
# Converter typography report

Zażółć gęślą jaźń

| Region | Signal |
| --- | --- |
| Ελλάδα | Athens |
| Україна | Київ |
""").ToPdfDocumentResult(options);

        byte[] pdf = result.ToBytes();
        Assert.False(result.HasWarnings);
        return pdf;
    }

    private static byte[] CreateHtmlPdf(string fontPath) {
        var options = new HtmlPdfSaveOptions {
            FontFamily = PdfCore.PdfEmbeddedFontFamily.FromFiles(FamilyName, fontPath)
        };

        byte[] pdf = HtmlConversionDocument.Parse("""
<html>
  <body>
    <h1>Converter typography report</h1>
    <p>Zażółć gęślą jaźń</p>
    <table>
      <tr><th>Region</th><th>Signal</th></tr>
      <tr><td>Ελλάδα</td><td>Athens</td></tr>
      <tr><td>Україна</td><td>Київ</td></tr>
    </table>
  </body>
</html>
""").ToPdfDocumentResult(options).ToBytes();

        return pdf;
    }

    private static byte[] CreatePowerPointPdf(string fontPath) {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(520, 320);
        PowerPointSlide slide = presentation.AddSlide();
        PowerPointTextBox title = slide.AddTextBoxPoints("Converter typography report", 32, 26, 280, 28);
        title.FontSize = 14;
        PowerPointTextBox polish = slide.AddTextBoxPoints("Zażółć gęślą jaźń", 32, 70, 250, 28);
        polish.FontSize = 12;
        PowerPointTable table = slide.AddTablePoints(3, 2, 32, 118, 280, 98);
        table.GetCell(0, 0).Text = "Region";
        table.GetCell(0, 1).Text = "Signal";
        table.GetCell(1, 0).Text = "Ελλάδα";
        table.GetCell(1, 1).Text = "Athens";
        table.GetCell(2, 0).Text = "Україна";
        table.GetCell(2, 1).Text = "Київ";

        var options = new PowerPointPdfSaveOptions {
            PdfOptions = CreatePdfOptions(fontPath)
        };
        PdfCore.PdfDocumentConversionResult result = presentation.ToPdfDocumentResult(options);
        byte[] pdf = result.ToBytes();
        Assert.False(
            result.HasWarnings,
            string.Join(
                Environment.NewLine,
                result.Warnings.Select(static warning => warning.Code + ": " + warning.Message)));
        return pdf;
    }

    private static PdfCore.PdfConversionReport CreateWordOpenTypeReport(string fontPath) {
        string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO.Pdf.Typography", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            string documentPath = Path.Combine(directory, "opentype.docx");
            using WordDocument document = WordDocument.Create(documentPath);
            document.AddParagraph("office cafe\u0301");
            document.Save();

            var options = new WordPdf.PdfSaveOptions {
                PdfOptions = CreatePdfOptions(fontPath),
                IncludePageNumbers = false
            };
            PdfCore.PdfDocumentConversionResult result = document.ToPdfDocumentResult(options);
            _ = result.ToBytes();
            return result.Report;
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    private static PdfCore.PdfConversionReport CreateExcelOpenTypeReport(string fontPath) {
        string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO.Pdf.Typography", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            string workbookPath = Path.Combine(directory, "opentype.xlsx");
            using ExcelDocument document = ExcelDocument.Create(workbookPath, "Report");
            document.Sheets[0].Cell(1, 1, "office cafe\u0301");
            document.Save();

            var options = new ExcelPdfSaveOptions {
                PdfOptions = CreatePdfOptions(fontPath),
                IncludeSheetHeadings = false
            };
            PdfCore.PdfDocumentConversionResult result = document.ToPdfDocumentResult(options);
            _ = result.ToBytes();
            return result.Report;
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    private static PdfCore.PdfConversionReport CreateMarkdownOpenTypeReport(string fontPath) {
        var options = new MarkdownPdfSaveOptions {
            ApplyDefaultTheme = false,
            PdfOptions = CreatePdfOptions(fontPath)
        };
        PdfCore.PdfDocumentConversionResult result = OfficeIMO.Markdown.MarkdownReader.Parse("office cafe\u0301").ToPdfDocumentResult(options);
        _ = result.ToBytes();
        return result.Report;
    }

    private static PdfCore.PdfConversionReport CreatePowerPointOpenTypeReport(string fontPath) {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(520, 320);
        PowerPointTextBox textBox = presentation.AddSlide().AddTextBoxPoints("office cafe\u0301", 32, 32, 220, 32);
        textBox.FontSize = 14;

        var options = new PowerPointPdfSaveOptions {
            PdfOptions = CreatePdfOptions(fontPath)
        };
        PdfCore.PdfDocumentConversionResult result = presentation.ToPdfDocumentResult(options);
        _ = result.ToBytes();
        return result.Report;
    }

    private static PdfCore.PdfConversionReport CreateWordComplexScriptReport(string text = "مرحبا", bool allowMissingGlyphFailure = false) {
        string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO.Pdf.Typography", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            string documentPath = Path.Combine(directory, "complex-script.docx");
            using WordDocument document = WordDocument.Create(documentPath);
            document.AddParagraph(text);
            document.Save();

            var options = new WordPdf.PdfSaveOptions {
                IncludePageNumbers = false
            };
            PdfCore.PdfDocumentConversionResult? result = null;
            AssertRenderAttempt(() => (result = document.ToPdfDocumentResult(options)).ToBytes(), allowMissingGlyphFailure);
            return result?.Report ?? new PdfCore.PdfConversionReport();
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    private static PdfCore.PdfConversionReport CreateExcelComplexScriptReport(string text = "مرحبا", bool allowMissingGlyphFailure = false) {
        string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO.Pdf.Typography", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            string workbookPath = Path.Combine(directory, "complex-script.xlsx");
            using ExcelDocument document = ExcelDocument.Create(workbookPath, "Report");
            document.Sheets[0].Cell(1, 1, text);
            document.Save();

            var options = new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false
            };
            PdfCore.PdfDocumentConversionResult? result = null;
            AssertRenderAttempt(() => (result = document.ToPdfDocumentResult(options)).ToBytes(), allowMissingGlyphFailure);
            return result?.Report ?? new PdfCore.PdfConversionReport();
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    private static PdfCore.PdfConversionReport CreateMarkdownComplexScriptReport(string text = "مرحبا", bool allowMissingGlyphFailure = false) {
        var options = new MarkdownPdfSaveOptions {
            ApplyDefaultTheme = false
        };
        PdfCore.PdfDocumentConversionResult? result = null;
        AssertRenderAttempt(() => (result = OfficeIMO.Markdown.MarkdownReader.Parse(text).ToPdfDocumentResult(options)).ToBytes(), allowMissingGlyphFailure);
        return result?.Report ?? new PdfCore.PdfConversionReport();
    }

    private static PdfCore.PdfConversionReport CreatePowerPointComplexScriptReport(string text = "مرحبا", bool allowMissingGlyphFailure = false) {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(520, 320);
        PowerPointTextBox textBox = presentation.AddSlide().AddTextBoxPoints(text, 32, 32, 220, 32);
        textBox.FontSize = 14;

        var options = new PowerPointPdfSaveOptions();
        PdfCore.PdfDocumentConversionResult? result = null;
        AssertRenderAttempt(() => (result = presentation.ToPdfDocumentResult(options)).ToBytes(), allowMissingGlyphFailure);
        return result?.Report ?? new PdfCore.PdfConversionReport();
    }

    private static void AssertRenderAttempt(Func<byte[]> action, bool allowMissingGlyphFailure = false) {
        Exception? exception = Record.Exception(() => {
            byte[] pdf = action();
            Assert.NotEmpty(pdf);
        });

        if (exception == null) {
            return;
        }

        if (allowMissingGlyphFailure &&
            exception is ArgumentException argumentException &&
            IsMissingEmbeddedGlyphFailure(argumentException)) {
            return;
        }

        throw exception;
    }

    private static bool IsMissingEmbeddedGlyphFailure(ArgumentException exception) {
        if (exception.Data["code"] is string code &&
            (string.Equals(code, "missing-embedded-font-glyph", StringComparison.Ordinal) ||
             string.Equals(code, "missing-embedded-font-fallback-glyph", StringComparison.Ordinal))) {
            return true;
        }

        return exception.Message.Contains("not covered by the embedded TrueType font", StringComparison.Ordinal) ||
               exception.Message.Contains("cannot be encoded with embedded TrueType font", StringComparison.Ordinal) ||
               exception.Message.Contains("not covered by any embedded font fallback candidate", StringComparison.Ordinal) ||
               exception.Message.Contains("Embedded Unicode fonts are required for this text", StringComparison.Ordinal);
    }

    private static void WriteReviewArtifact(string fileName, byte[] bytes) {
        string? outputDirectory = Environment.GetEnvironmentVariable("OFFICEIMO_PDF_VISUAL_REVIEW_OUTPUT");
        if (string.IsNullOrWhiteSpace(outputDirectory)) {
            return;
        }

        Directory.CreateDirectory(outputDirectory);
        File.WriteAllBytes(Path.Combine(outputDirectory, fileName), bytes);
    }
}
