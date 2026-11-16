using System.Text;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfFontFamilyTests {
    private const int FillThenStroke = 2;

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FallbackBoldRunsPaintWithTheFamilysBoldFace(bool compatibilitySlots) {
        byte[] regular = ManagedTextShapingTestAssets.CreateFontWithUniformAdvance(1000, 500, ' ', 0x1F600, 0x2603);
        byte[] bold = ManagedTextShapingTestAssets.CreateFontWithUniformAdvance(1000, 700, ' ', 0x1F600, 0x2603);

        PdfTextSpan[] spans = RenderBoldFallbackRun(new PdfEmbeddedFontFamily("Styled Fallback", regular, bold), compatibilitySlots, "😀☃");

        Assert.Equal("😀☃", string.Concat(spans.Select(span => span.Text)));
        Assert.All(spans, span => Assert.NotEqual(FillThenStroke, span.TextRenderingMode));
        // Two glyphs at 10 pt with the bold face's 700-unit advance.
        Assert.Equal(14D, spans.Sum(span => span.Advance), 2);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FallbackBoldRunsSynthesizeOnlyGlyphsTheBoldFaceLacks(bool compatibilitySlots) {
        byte[] regular = ManagedTextShapingTestAssets.CreateFontWithUniformAdvance(1000, 500, ' ', 0x1F600, 0x2603);
        byte[] bold = ManagedTextShapingTestAssets.CreateFontWithUniformAdvance(1000, 700, ' ', 0x1F600);

        PdfTextSpan[] spans = RenderBoldFallbackRun(new PdfEmbeddedFontFamily("Partial Fallback", regular, bold), compatibilitySlots, "😀☃");

        PdfTextSpan real = Assert.Single(spans, span => span.Text == "😀");
        Assert.NotEqual(FillThenStroke, real.TextRenderingMode);
        Assert.Equal(7D, real.Advance, 2);
        PdfTextSpan synthesized = Assert.Single(spans, span => span.Text == "☃");
        Assert.Equal(FillThenStroke, synthesized.TextRenderingMode);
        Assert.Equal(5D, synthesized.Advance, 2);
    }

    [Fact]
    public void FallbackBoldFaceThatForbidsEmbeddingIsSynthesizedFromTheRegularFace() {
        byte[] regular = File.ReadAllBytes(Assert.IsType<string>(PdfComplianceTestFonts.FindBundledTrueTypeFont()));
        byte[] restricted = (byte[])regular.Clone();
        WriteOs2FsType(restricted, 0x0002);

        PdfTextSpan[] spans = RenderBoldFallbackRun(new PdfEmbeddedFontFamily("Licensed Fallback", regular, restricted), compatibilitySlots: true, "ЖЗ");

        Assert.Equal("ЖЗ", string.Concat(spans.Select(span => span.Text)));
        Assert.All(spans, span => Assert.Equal(FillThenStroke, span.TextRenderingMode));
    }

    [Fact]
    public void FallbackRegularFaceCompanionKeepsAnotherFallbackWithTheSameName() {
        byte[] regular = ManagedTextShapingTestAssets.CreateFontWithUniformAdvance(1000, 500, ' ', 0x2603);
        byte[] bold = ManagedTextShapingTestAssets.CreateFontWithUniformAdvance(1000, 700, ' ', 0x1F600);
        byte[] other = ManagedTextShapingTestAssets.CreateFontWithUniformAdvance(1000, 600, ' ', 0x2602);
        var fallbacks = new PdfEmbeddedFontFallbackSet(new[] {
            new PdfEmbeddedFontFallbackCandidate(new PdfEmbeddedFontFamily("Pair", regular, bold)),
            new PdfEmbeddedFontFallbackCandidate("Pair Regular Face", other)
        });
        byte[] bytes = PdfDocument.Create(new PdfCreateOptions { CompressContentStreams = false })
            .EmbedStandardFont(PdfStandardFont.Helvetica, CreateRegistrationPrimaryFont(), "OfficeIMO Primary")
            .RegisterEmbeddedFontFallbacks(fallbacks)
            .Paragraph(paragraph => paragraph.AddText("A☃☂").SetBold().SetFontSize(global::OfficeIMO.OfficeLength.Points(10)))
            .ToBytes();

        Assert.Contains("A☃☂", PdfReadDocument.Open(bytes).ExtractText(), StringComparison.Ordinal);
        PdfTextSpan[] spans = PdfReadDocument.Open(bytes).Pages[0].GetTextSpans().ToArray();
        Assert.Equal(FillThenStroke, Assert.Single(spans, span => span.Text == "☃").TextRenderingMode);
        Assert.Equal(6D, Assert.Single(spans, span => span.Text == "☂").Advance, 2);
    }

    [Fact]
    public void FallbackRegularFaceCompanionWithoutNamedFamilyCapacityIsNotLeftUnresolved() {
        byte[] regular = ManagedTextShapingTestAssets.CreateFontWithUniformAdvance(1000, 500, ' ', 0x1F600, 0x2603);
        byte[] bold = ManagedTextShapingTestAssets.CreateFontWithUniformAdvance(1000, 700, ' ', 0x1F600);
        var options = new PdfCreateOptions { CompressContentStreams = false }
            .EmbedStandardFont(PdfStandardFont.Helvetica, CreateRegistrationPrimaryFont(), "OfficeIMO Primary");
        for (int index = 0; index < PdfCreateOptions.MaximumNamedFontFamilies; index++) {
            options.RegisterNamedFontFamily(new PdfEmbeddedFontFamily("Caller " + index, CreateRegistrationPrimaryFont()));
        }
        options.RegisterEmbeddedFontFallbacks(new PdfEmbeddedFontFallbackSet(
            new[] { new PdfEmbeddedFontFallbackCandidate(new PdfEmbeddedFontFamily("Full Fallback", regular, bold)) },
            new[] { PdfStandardFont.TimesRoman }));

        // The companion cannot be registered, so the run is reported instead of pointing at an unknown family.
        Exception exception = Assert.ThrowsAny<ArgumentException>(() => PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.AddText("A☃").SetBold())
            .ToBytes());
        Assert.Contains("U+2603", exception.Message, StringComparison.Ordinal);
    }

    private static PdfTextSpan[] RenderBoldFallbackRun(PdfEmbeddedFontFamily family, bool compatibilitySlots, string fallbackText) {
        var candidates = new[] { new PdfEmbeddedFontFallbackCandidate(family) };
        PdfEmbeddedFontFallbackSet fallbacks = compatibilitySlots
            ? new PdfEmbeddedFontFallbackSet(candidates, new[] { PdfStandardFont.TimesRoman })
            : new PdfEmbeddedFontFallbackSet(candidates);
        byte[] bytes = PdfDocument.Create(new PdfCreateOptions { CompressContentStreams = false })
            .EmbedStandardFont(PdfStandardFont.Helvetica, CreateRegistrationPrimaryFont(), "OfficeIMO Primary")
            .RegisterEmbeddedFontFallbacks(fallbacks)
            .Paragraph(paragraph => paragraph.AddText("A" + fallbackText).SetBold().SetFontSize(global::OfficeIMO.OfficeLength.Points(10)))
            .ToBytes();

        Assert.Contains("A" + fallbackText, PdfReadDocument.Open(bytes).ExtractText(), StringComparison.Ordinal);
        return PdfReadDocument.Open(bytes).Pages[0].GetTextSpans()
            .Where(span => !string.IsNullOrWhiteSpace(span.Text) && span.Text != "A")
            .ToArray();
    }

    private static void WriteOs2FsType(byte[] font, ushort fsType) {
        int count = (font[4] << 8) | font[5];
        for (int index = 0; index < count; index++) {
            int record = 12 + index * 16;
            if (Encoding.ASCII.GetString(font, record, 4) != "OS/2") continue;
            int offset = (font[record + 8] << 24) | (font[record + 9] << 16) | (font[record + 10] << 8) | font[record + 11];
            font[offset + 8] = (byte)(fsType >> 8);
            font[offset + 9] = (byte)fsType;
            return;
        }
        throw new InvalidOperationException("The test font has no OS/2 table.");
    }
}
