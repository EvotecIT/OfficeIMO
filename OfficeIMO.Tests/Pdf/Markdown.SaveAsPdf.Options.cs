using OfficeIMO.Markdown.Pdf;
using PdfCore = OfficeIMO.Pdf;
using System;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class MarkdownSaveAsPdfOptionsTests {
    [Fact]
    public void ToPdfDocument_Markdown_ClonesCallerPdfOptionsBeforeApplyingAdapterDefaults() {
        var pdfOptions = new PdfCore.PdfOptions();
        var options = new MarkdownPdfSaveOptions {
            PdfOptions = pdfOptions,
            CreateOutlineFromHeadings = true
        };

        "# Heading".ToPdfDocument(options).ToBytes();

        Assert.False(pdfOptions.CreateOutlineFromHeadings);
    }

    [Fact]
    public void ToPdfDocument_Markdown_FontFamilyUsesSharedOfficeFontMapping() {
        var options = new MarkdownPdfSaveOptions {
            FontFamily = "Georgia",
            PdfOptions = new PdfCore.PdfOptions {
                CompressContentStreams = false
            }
        };

        byte[] bytes = "# Heading\n\nBody".ToPdfDocument(options).ToBytes();
        string raw = Encoding.ASCII.GetString(bytes);

        Assert.True(
            raw.Contains("/BaseFont /Georgia-Regular", StringComparison.Ordinal) ||
            raw.Contains("/BaseFont /Times-Roman", StringComparison.Ordinal),
            "Expected Markdown font-family export to use an installed Georgia TrueType font or the mapped Times standard family.");
    }

    [Fact]
    public void ToPdfDocument_Markdown_DefaultsUseSharedUnicodeFallbackWhenAvailable() {
        var probe = new PdfCore.PdfOptions();
        if (!probe.TryUseDefaultDocumentFontFallback(requireEmbeddedFont: true)) {
            return;
        }

        const string polish = "Zażółć gęślą jaźń Łódź";
        byte[] bytes = ("# Faktura\n\n" + polish).ToPdfDocument(new MarkdownPdfSaveOptions()).ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        string text = PdfCore.PdfReadDocument.Load(bytes).ExtractText();

        Assert.Contains("/Subtype /Type0", raw, StringComparison.Ordinal);
        Assert.Contains("/Encoding /Identity-H", raw, StringComparison.Ordinal);
        Assert.Contains(polish, text, StringComparison.Ordinal);
    }
}
