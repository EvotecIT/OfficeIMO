using System;
using System.Linq;
using System.Text;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfFontFamilyTests {
    [Fact]
    public void PdfDocument_RegisteredFallbacksPreserveSelectedFontCoveredSpans() {
        byte[] bytes = PdfDocument.Create(new PdfOptions { CompressContentStreams = false })
            .EmbedStandardFont(PdfStandardFont.Helvetica, CreateRegistrationPrimaryFont(), "OfficeIMO Primary")
            .RegisterEmbeddedFontFallbacks(CreateRegistrationEmojiFallbacks(coverLatin: true))
            .Paragraph(paragraph => paragraph.Text("Invoice 😀 marker"))
            .ToBytes();

        AssertRegistrationFonts(bytes, "OfficeIMOPrimary", "EmojiFallback");
        string extracted = PdfReadDocument.Open(bytes).ExtractText();
        Assert.Contains("Invoice", extracted, StringComparison.Ordinal);
        Assert.Contains("😀", extracted, StringComparison.Ordinal);
        Assert.Contains("marker", extracted, StringComparison.Ordinal);
        var spans = PdfReadDocument.Open(bytes).Pages[0].GetTextSpans();
        Assert.Equal("Invoicemarker", string.Concat(spans
            .Where(span => span.BaseFont == "OfficeIMOPrimary" && !string.IsNullOrWhiteSpace(span.Text))
            .Select(span => span.Text)));
        Assert.Equal("😀", string.Concat(spans
            .Where(span => span.BaseFont == "EmojiFallback-Regular" && !string.IsNullOrWhiteSpace(span.Text))
            .Select(span => span.Text)));
    }

    [Fact]
    public void PdfDocument_RegisteredFallbacksSplitMixedUnsupportedTokenWithoutDroppingSelectedText() {
        byte[] bytes = PdfDocument.Create(new PdfOptions { CompressContentStreams = false })
            .EmbedStandardFont(PdfStandardFont.Helvetica, CreateRegistrationPrimaryFont(), "OfficeIMO Primary")
            .RegisterEmbeddedFontFallbacks(CreateRegistrationEmojiFallbacks())
            .Paragraph(paragraph => paragraph.Text("A😀B"))
            .ToBytes();

        AssertRegistrationFonts(bytes, "OfficeIMOPrimary", "EmojiFallback");
        Assert.Contains("A😀B", PdfReadDocument.Open(bytes).ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void PdfDocument_RegisteredFallbacksResolveOnlyUsedCandidates() {
        var fallbacks = new PdfEmbeddedFontFallbackSet(new[] {
            new PdfEmbeddedFontFallbackCandidate("Emoji Fallback", ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(' ', 0x1F600)),
            new PdfEmbeddedFontFallbackCandidate("Unused Secondary Fallback", CreateMinimalOpenTypeCffFont()),
            new PdfEmbeddedFontFallbackCandidate("Unused Tertiary Fallback", CreateMinimalOpenTypeCffFont())
        }, new[] { PdfStandardFont.Helvetica, PdfStandardFont.TimesRoman, PdfStandardFont.Courier });
        byte[] bytes = PdfDocument.Create(new PdfOptions { CompressContentStreams = false })
            .EmbedStandardFont(PdfStandardFont.Helvetica, CreateRegistrationPrimaryFont(), "OfficeIMO Primary")
            .RegisterEmbeddedFontFallbacks(fallbacks)
            .Paragraph(paragraph => paragraph.Text("Invoice 😀 marker"))
            .ToBytes();

        AssertRegistrationFonts(bytes, "OfficeIMOPrimary", "EmojiFallback");
        string extracted = PdfReadDocument.Open(bytes).ExtractText();
        Assert.Contains("Invoice", extracted, StringComparison.Ordinal);
        Assert.Contains("😀", extracted, StringComparison.Ordinal);
        Assert.Contains("marker", extracted, StringComparison.Ordinal);
        // Unused data stays unparsed; requesting a missing glyph still validates the
        // next candidate instead of silently accepting an unusable font.
        Assert.Throws<NotSupportedException>(() => fallbacks.PlanText("Ω"));
    }

    [Fact]
    public void PdfDocument_RegisteredFallbackReplacementDoesNotOverwriteDocumentDefaultFontSlot() {
        byte[] primary = CreateRegistrationPrimaryFont();
        var options = new PdfOptions { CompressContentStreams = false };
        options.RegisterFontFamily(PdfStandardFont.Helvetica, new PdfEmbeddedFontFamily("OfficeIMO Default", primary));
        options.RegisterFontFamily(PdfStandardFont.TimesRoman, new PdfEmbeddedFontFamily("OfficeIMO Primary", primary));
        options.RegisterEmbeddedFontFallbacks(CreateRegistrationEmojiFallbacks());
        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.Text("Plain text"))
            .Paragraph(paragraph => paragraph.Font(PdfStandardFont.TimesRoman).Text("A😀B"))
            .ToBytes();

        AssertRegistrationFonts(bytes, "OfficeIMODefault", "OfficeIMOPrimary", "EmojiFallback");
        string extracted = PdfReadDocument.Open(bytes).ExtractText();
        Assert.Contains("Plain text", extracted, StringComparison.Ordinal);
        Assert.Contains("A😀B", extracted, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfDocument_ReplacingRegisteredFallbacksReusesPriorFallbackOwnedSlot() {
        var options = new PdfOptions { CompressContentStreams = false }
            .RegisterFontFamily(PdfStandardFont.Helvetica, new PdfEmbeddedFontFamily("Document Primary", CreateRegistrationPrimaryFont()))
            .RegisterFontFamily(PdfStandardFont.Courier, new PdfEmbeddedFontFamily("Document Mono", CreateRegistrationPrimaryFont()))
            .RegisterEmbeddedFontFallbacks(new PdfEmbeddedFontFallbackSet(
                new[] { new PdfEmbeddedFontFallbackCandidate("Old Symbols", ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(' ', 0x2605)) },
                new[] { PdfStandardFont.TimesRoman }))
            .RegisterEmbeddedFontFallbacks(CreateRegistrationEmojiFallbacks());

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.Text("A😀B"))
            .ToBytes();
        Assert.Contains("A😀B", PdfReadDocument.Open(bytes).ExtractText(), StringComparison.Ordinal);
        AssertRegistrationFonts(bytes, "DocumentPrimary-Regular", "EmojiFallback-Regular");
        Assert.Equal("Emoji Fallback-Regular", options.EmbeddedFonts[PdfStandardFont.TimesRoman].FontName);
    }

    [Fact]
    public void PdfDocument_ReplacingFallbacksPreservesLaterCallerFamilyWithSameData() {
        byte[] symbols = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(' ', 0x2605);
        var options = new PdfOptions { CompressContentStreams = false }
            .RegisterFontFamily(PdfStandardFont.Helvetica, new PdfEmbeddedFontFamily("Document Primary", CreateRegistrationPrimaryFont()))
            .RegisterEmbeddedFontFallbacks(new PdfEmbeddedFontFallbackSet(
                new[] { new PdfEmbeddedFontFallbackCandidate("Old Symbols", symbols) },
                new[] { PdfStandardFont.TimesRoman }))
            .RegisterFontFamily(PdfStandardFont.TimesRoman, new PdfEmbeddedFontFamily("Caller Symbols", symbols))
            .RegisterEmbeddedFontFallbacks(CreateRegistrationEmojiFallbacks());

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.Font(PdfStandardFont.TimesRoman).Text("★"))
            .Paragraph(paragraph => paragraph.Text("A😀B"))
            .ToBytes();
        AssertRegistrationFonts(bytes, "CallerSymbols-Regular", "EmojiFallback-Regular");
        Assert.Equal("Caller Symbols-Regular", options.EmbeddedFonts[PdfStandardFont.TimesRoman].FontName);
        string extracted = PdfReadDocument.Open(bytes).ExtractText();
        Assert.Contains("★", extracted, StringComparison.Ordinal);
        Assert.Contains("A😀B", extracted, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PdfDocument_ReplacingFallbacksPreservesLaterCallerItalicFaceWithSameNameAndData(bool bold) {
        byte[] symbols = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(' ', 0x2605);
        var options = new PdfOptions { CompressContentStreams = false }
            .RegisterFontFamily(PdfStandardFont.Helvetica, new PdfEmbeddedFontFamily("Document Primary", CreateRegistrationPrimaryFont()))
            .RegisterEmbeddedFontFallbacks(new PdfEmbeddedFontFallbackSet(
                new[] { new PdfEmbeddedFontFallbackCandidate("Old Symbols", symbols) },
                new[] { PdfStandardFont.TimesRoman }))
            .RegisterFontFamily(PdfStandardFont.TimesRoman, new PdfEmbeddedFontFamily("Old Symbols", symbols,
                italic: bold ? null : symbols, boldItalic: bold ? symbols : null))
            .RegisterEmbeddedFontFallbacks(CreateRegistrationEmojiFallbacks());

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.Runs(new[] { new PdfTextRun("★", bold: bold, italic: true, font: PdfStandardFont.TimesRoman) }))
            .Paragraph(paragraph => paragraph.Text("A😀B"))
            .ToBytes();
        AssertRegistrationFonts(bytes, bold ? "OldSymbols-BoldItalic" : "OldSymbols-Italic", "EmojiFallback-Regular");
        string extracted = PdfReadDocument.Open(bytes).ExtractText();
        Assert.Contains("★", extracted, StringComparison.Ordinal);
        Assert.Contains("A😀B", extracted, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfDocument_PreservingPrimaryFallbackSpansAppliesHeaderOffsetOnce() {
        static PdfTextSpan[] Render(double offset) {
            byte[] bytes = PdfDocument.Create(new PdfOptions { CompressContentStreams = false })
                .EmbedStandardFont(PdfStandardFont.Helvetica, CreateRegistrationPrimaryFont(), "OfficeIMO Primary")
                .RegisterEmbeddedFontFallbacks(CreateRegistrationEmojiFallbacks(coverLatin: true))
                .Header(header => header.StyledZones(
                    left => left.Run(new PdfTextRun("A😀B").WithHorizontalOffset(offset)), null, null))
                .Paragraph(paragraph => paragraph.Text("Plain text"))
                .ToBytes();
            return PdfReadDocument.Open(bytes).Pages[0].GetTextSpans()
                .Where(span => span.Text == "A" || span.Text == "😀" || span.Text == "B").ToArray();
        }

        PdfTextSpan[] baseline = Render(0D);
        PdfTextSpan[] shifted = Render(24D);
        Assert.Equal(new[] { "A", "😀", "B" }, baseline.Select(span => span.Text));
        Assert.Equal(baseline.Select(span => span.Text), shifted.Select(span => span.Text));
        for (int index = 0; index < baseline.Length; index++) {
            Assert.InRange(shifted[index].X - baseline[index].X, 23.99D, 24.01D);
        }
    }

    // The selected font deliberately covers the Latin spans but not the emoji.
    // Installed fonts cannot establish that contract consistently across platforms.
    private static byte[] CreateRegistrationPrimaryFont() =>
        ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(
            "Invoice markerABPlain text".Select(character => (int)character).Distinct().ToArray());

    private static PdfEmbeddedFontFallbackSet CreateRegistrationEmojiFallbacks(bool coverLatin = false) => new(
        new[] { new PdfEmbeddedFontFallbackCandidate("Emoji Fallback",
            ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(
                (coverLatin ? "Invoice markerABPlain text" : " ").Select(character => (int)character)
                    .Append(0x1F600).ToArray())) },
        new[] { PdfStandardFont.TimesRoman });

    private static void AssertRegistrationFonts(byte[] bytes, params string[] names) {
        string raw = Encoding.ASCII.GetString(bytes);
        foreach (string name in names) Assert.Contains("/BaseFont /" + name, raw, StringComparison.Ordinal);
    }
}
