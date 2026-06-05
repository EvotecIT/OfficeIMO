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
}
