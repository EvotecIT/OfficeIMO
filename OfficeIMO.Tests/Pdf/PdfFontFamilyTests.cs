using System;
using System.Globalization;
using System.IO;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfFontFamilyTests {
    [Fact]
    public void PdfEmbeddedFontFamily_SnapshotsFaceBytesForReusableRegistration() {
        byte[] regular = { 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
        byte[] bold = { 0, 1, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0 };
        var family = new PdfEmbeddedFontFamily("Snapshot Family", regular, bold);
        regular[0] = 255;
        bold[5] = 255;

        PdfOptions options = new PdfOptions().UseFontFamily(family);
        byte[] readback = family.Regular;
        readback[0] = 127;

        Assert.Equal("Snapshot Family", family.FamilyName);
        Assert.Equal(0, family.Regular[0]);
        Assert.Equal(1, family.Bold![5]);
        Assert.Equal("Snapshot Family-Regular", options.EmbeddedFonts[PdfStandardFont.Helvetica].FontName);
        Assert.Equal("Snapshot Family-Bold", options.EmbeddedFonts[PdfStandardFont.HelveticaBold].FontName);
        Assert.Equal(PdfStandardFont.Helvetica, options.DefaultFont);
        Assert.Equal(PdfStandardFont.Helvetica, options.HeaderFont);
        Assert.Equal(PdfStandardFont.Helvetica, options.FooterFont);
        Assert.Throws<ArgumentException>(() => new PdfEmbeddedFontFamily("Broken", Array.Empty<byte>()));
    }

    [Fact]
    public void PdfDoc_UseFontFamilyObjectReusesTrueTypeFamilyForGeneratedText() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        byte[] fontData = File.ReadAllBytes(fontPath);
        var family = new PdfEmbeddedFontFamily("OfficeIMO Object Font", fontData);
        byte[] bytes = PdfDoc.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .UseFontFamily(family)
            .Header(header => header.Text("Object font header"))
            .Paragraph(paragraph => paragraph
                .Text("Object regular ")
                .Bold("object bold ")
                .Italic("object italic"))
            .Footer(footer => footer.Text("Object font footer {page}/{pages}"))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        string text = PdfReadDocument.Load(bytes).ExtractText();

        Assert.Contains("/Subtype /TrueType", raw, StringComparison.Ordinal);
        Assert.Contains("/BaseFont /OfficeIMOObjectFont-Regular", raw, StringComparison.Ordinal);
        Assert.Contains("/BaseFont /OfficeIMOObjectFont-Bold", raw, StringComparison.Ordinal);
        Assert.Contains("/BaseFont /OfficeIMOObjectFont-Italic", raw, StringComparison.Ordinal);
        Assert.Contains("/Length1 " + fontData.Length.ToString(CultureInfo.InvariantCulture), raw, StringComparison.Ordinal);
        Assert.Contains("Object regular object bold object italic", text, StringComparison.Ordinal);
        Assert.Contains("Object font header", text, StringComparison.Ordinal);
        Assert.Contains("Object font footer 1/1", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ComposePage_UseFontFamilyScopesFamilyToComposedPageHeaderFooterAndBody() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        var family = PdfEmbeddedFontFamily.FromFiles("OfficeIMO Page Font", fontPath);
        byte[] bytes = PdfDoc.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .Paragraph(paragraph => paragraph.Text("Before page font"))
            .Page(page => page
                .UseFontFamily(family)
                .Header(header => header.Text("Page font header"))
                .Content(content => content.Item(item => item.Paragraph(paragraph => paragraph.Text("Page font body"))))
                .Footer(footer => footer.Text("Page font footer {page}/{pages}")))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        string text = PdfReadDocument.Load(bytes).ExtractText();

        Assert.Contains("/BaseFont /OfficeIMOPageFont-Regular", raw, StringComparison.Ordinal);
        Assert.Contains("Before page font", text, StringComparison.Ordinal);
        Assert.Contains("Page font header", text, StringComparison.Ordinal);
        Assert.Contains("Page font body", text, StringComparison.Ordinal);
        Assert.Contains("Page font footer", text, StringComparison.Ordinal);
    }

    [Fact]
    public void DefaultTextStyle_FontFamilyDoesNotRewriteExistingHeaderOrFooterFonts() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        byte[] fontData = File.ReadAllBytes(fontPath);
        var family = new PdfEmbeddedFontFamily("OfficeIMO Style Font", fontData);
        byte[] bytes = PdfDoc.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .Header(header => header.Font(PdfStandardFont.TimesRoman).Text("Times header"))
            .Footer(footer => footer.Font(PdfStandardFont.Courier).Text("Courier footer"))
            .DefaultTextStyle(style => style.FontFamily(family).FontSize(12))
            .Paragraph(paragraph => paragraph.Text("Styled body"))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        string text = PdfReadDocument.Load(bytes).ExtractText();

        Assert.Contains("/BaseFont /OfficeIMOStyleFont-Regular", raw, StringComparison.Ordinal);
        Assert.Contains("/BaseFont /Times-Roman", raw, StringComparison.Ordinal);
        Assert.Contains("/BaseFont /Courier", raw, StringComparison.Ordinal);
        Assert.Contains("Times header", text, StringComparison.Ordinal);
        Assert.Contains("Styled body", text, StringComparison.Ordinal);
        Assert.Contains("Courier footer", text, StringComparison.Ordinal);
    }
}
