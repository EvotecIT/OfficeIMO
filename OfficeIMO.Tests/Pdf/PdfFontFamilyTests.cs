using System;
using System.Collections.Generic;
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

        Assert.Contains("/Subtype /Type0", raw, StringComparison.Ordinal);
        Assert.Contains("/Subtype /CIDFontType2", raw, StringComparison.Ordinal);
        Assert.Contains("/Encoding /Identity-H", raw, StringComparison.Ordinal);
        Assert.Contains("/CIDToGIDMap /Identity", raw, StringComparison.Ordinal);
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

    [Fact]
    public void PdfDoc_UseFontFamilyWritesUnicodeGlyphsAndToUnicodeExtraction() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        const string polish = "Zażółć gęślą jaźń Łódź";
        byte[] bytes = PdfDoc.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .UseFontFamily("OfficeIMO Unicode Font", fontPath)
            .Header(header => header.Text("Nagłówek Łódź"))
            .Paragraph(paragraph => paragraph
                .Text(polish + " ")
                .Bold("Śląsk ")
                .Italic("źrebak"))
            .Footer(footer => footer.Text("Stopka Łódź {page}/{pages}"))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        string text = PdfReadDocument.Load(bytes).ExtractText();

        Assert.Contains("/Subtype /Type0", raw, StringComparison.Ordinal);
        Assert.Contains("/Encoding /Identity-H", raw, StringComparison.Ordinal);
        Assert.Contains("/CMapName /OfficeIMO-Identity-Glyph-UCS", raw, StringComparison.Ordinal);
        Assert.Contains(polish, text, StringComparison.Ordinal);
        Assert.Contains("Nagłówek Łódź", text, StringComparison.Ordinal);
        Assert.Contains("Śląsk źrebak", text, StringComparison.Ordinal);
        Assert.Contains("Stopka Łódź 1/1", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfDoc_UseFontFamilyEncodesTextWatermarkWithEmbeddedGlyphs() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        byte[] pdf = PdfDoc.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .UseFontFamily("OfficeIMO Watermark Font", fontPath)
            .Watermark("DRAFT", fontSize: 32, opacity: 0.25)
            .Paragraph(paragraph => paragraph.Text("Watermark font proof"))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(pdf);

        Assert.Contains("/Subtype /Type0", raw, StringComparison.Ordinal);
        Assert.DoesNotContain("<4452414654> Tj", raw, StringComparison.Ordinal);
        Assert.Contains("/CMapName /OfficeIMO-Identity-Glyph-UCS", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfDoc_UseFontFamilyWrapsLongNonBmpTokensWithoutSplittingSurrogates() {
        foreach (string fontPath in EnumerateLocalNonBmpTrueTypeFonts()) {
            byte[] bytes;
            try {
                bytes = PdfDoc.Create(new PdfOptions {
                    CompressContentStreams = false,
                    PageWidth = 120,
                    MarginLeft = 24,
                    MarginRight = 24
                })
                .UseFontFamily("OfficeIMO NonBmp Font", fontPath)
                .Paragraph(paragraph => paragraph.Text("AAAAAAAAAAAA😀BBBBBBBBBBBB"))
                .ToBytes();
            } catch (ArgumentException exception) when (exception.Message.Contains("not covered by the embedded TrueType font", StringComparison.Ordinal)) {
                continue;
            }

            string raw = Encoding.ASCII.GetString(bytes);

            Assert.Contains("/Subtype /Type0", raw, StringComparison.Ordinal);
            Assert.Contains("/Encoding /Identity-H", raw, StringComparison.Ordinal);
            return;
        }
    }

    private static IEnumerable<string> EnumerateLocalNonBmpTrueTypeFonts() {
        string windowsFonts = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
        string[] candidates = {
            Path.Combine(windowsFonts, "seguiemj.ttf"),
            Path.Combine(windowsFonts, "seguisym.ttf"),
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
            "/usr/share/fonts/truetype/noto/NotoSansSymbols2-Regular.ttf"
        };

        foreach (string candidate in candidates) {
            if (File.Exists(candidate)) {
                yield return candidate;
            }
        }
    }
}
