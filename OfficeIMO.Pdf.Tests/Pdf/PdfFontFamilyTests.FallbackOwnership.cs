using System;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfFontFamilyTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PdfDocument_FallbackPlanningPreservesEarlierDirectRunsInTheSameParagraph(bool fallbackOwned) {
        byte[] emoji = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(' ', 0x1F600);
        var options = new PdfOptions { CompressContentStreams = false }
            .RegisterFontFamily(PdfStandardFont.Helvetica, new PdfEmbeddedFontFamily("Document Primary", CreateRegistrationPrimaryFont()));
        options.RegisterEmbeddedFontFallbacks(fallbackOwned
            ? new PdfEmbeddedFontFallbackSet(new[] {
                new PdfEmbeddedFontFallbackCandidate("Direct Symbols", ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(' ', 0x2605)),
                new PdfEmbeddedFontFallbackCandidate("Run Emoji", emoji)
            }, new[] { PdfStandardFont.TimesRoman, PdfStandardFont.Helvetica })
            : new PdfEmbeddedFontFallbackSet(new[] { new PdfEmbeddedFontFallbackCandidate("Run Emoji", emoji) },
                new[] { PdfStandardFont.Helvetica }));
        string directText = fallbackOwned ? "★" : "AB";

        byte[] bytes = PdfDocument.Create(options).Paragraph(paragraph => paragraph.Runs(new[] {
            new PdfTextRun(directText, font: PdfStandardFont.TimesRoman), new PdfTextRun("A😀B")
        })).ToBytes();

        AssertRegistrationFonts(bytes, fallbackOwned ? "DirectSymbols-Regular" : "Times-Roman", "RunEmoji-Regular");
        Assert.Contains(directText + "A😀B", PdfReadDocument.Open(bytes).ExtractText(), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void PdfDocument_NestedFallbackDoesNotReplaceADirectlyPaintedFont(bool fallbackOwned, bool canvas) {
        byte[] emoji = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(' ', 0x1F600);
        var candidates = fallbackOwned
            ? new[] {
                new PdfEmbeddedFontFallbackCandidate("Direct Symbols", ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(' ', 0x2605)),
                new PdfEmbeddedFontFallbackCandidate("Panel Emoji", emoji)
            }
            : new[] { new PdfEmbeddedFontFallbackCandidate("Panel Emoji", emoji) };
        var slots = fallbackOwned
            ? new[] { PdfStandardFont.TimesRoman, PdfStandardFont.Helvetica }
            : new[] { PdfStandardFont.Helvetica };
        var options = new PdfOptions { CompressContentStreams = false }
            .RegisterFontFamily(PdfStandardFont.Helvetica, new PdfEmbeddedFontFamily("Document Primary", CreateRegistrationPrimaryFont()))
            .RegisterEmbeddedFontFallbacks(new PdfEmbeddedFontFallbackSet(candidates, slots));
        string directText = fallbackOwned ? "★" : "AB";
        var document = PdfDocument.Create(options);
        if (canvas) document.Canvas(paint => paint.Text(directText, 0, 0, 100, 20, font: PdfStandardFont.TimesRoman));
        else document.Paragraph(paragraph => paragraph.Font(PdfStandardFont.TimesRoman).Text(directText));

        byte[] bytes = document.Panel(content => content
            .Paragraph(paragraph => paragraph.Text("A"))
            .Paragraph(paragraph => paragraph.Text("A😀B")),
            new PdfPanelStyle { KeepTogether = false, KeepWithNext = false }).ToBytes();

        AssertRegistrationFonts(bytes, fallbackOwned ? "DirectSymbols-Regular" : "Times-Roman", "PanelEmoji-Regular");
        string text = PdfReadDocument.Open(bytes).ExtractText();
        Assert.Contains(directText, text, StringComparison.Ordinal);
        Assert.Contains("A😀B", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfDocument_NestedPanelPreservesFallbackSlotsAlreadyUsedByTheParent() {
        var options = new PdfOptions { CompressContentStreams = false }
            .RegisterFontFamily(PdfStandardFont.Helvetica, new PdfEmbeddedFontFamily("Document Primary", CreateRegistrationPrimaryFont()))
            .RegisterEmbeddedFontFallbacks(new PdfEmbeddedFontFallbackSet(
                new[] {
                    new PdfEmbeddedFontFallbackCandidate("Parent Emoji", ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(' ', 0x1F600)),
                    new PdfEmbeddedFontFallbackCandidate("Panel Symbols", ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(' ', 0x2605))
                },
                new[] { PdfStandardFont.Helvetica, PdfStandardFont.TimesRoman }));

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.Text("A😀B"))
            .Panel(content => content
                .Paragraph(paragraph => paragraph.Text("A"))
                .Paragraph(paragraph => paragraph.Text("A★B")),
                new PdfPanelStyle { KeepTogether = false, KeepWithNext = false })
            .ToBytes();

        AssertRegistrationFonts(bytes, "ParentEmoji-Regular", "PanelSymbols-Regular");
        string text = PdfReadDocument.Open(bytes).ExtractText();
        Assert.Contains("A😀B", text, StringComparison.Ordinal);
        Assert.Contains("A★B", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfDocument_ClearingFallbacksReleasesOwnedSlotsForTheNextSet() {
        var options = new PdfOptions { CompressContentStreams = false }
            .RegisterFontFamily(PdfStandardFont.Helvetica, new PdfEmbeddedFontFamily("Document Primary", CreateRegistrationPrimaryFont()))
            .RegisterFontFamily(PdfStandardFont.Courier, new PdfEmbeddedFontFamily("Document Mono", CreateRegistrationPrimaryFont()))
            .RegisterEmbeddedFontFallbacks(new PdfEmbeddedFontFallbackSet(
                new[] { new PdfEmbeddedFontFallbackCandidate("Old Symbols", ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(' ', 0x2605)) },
                new[] { PdfStandardFont.TimesRoman }));
        options.EmbeddedFontFallbacks = null;
        options.RegisterEmbeddedFontFallbacks(new PdfEmbeddedFontFallbackSet(
            new[] { new PdfEmbeddedFontFallbackCandidate("New Emoji", ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(' ', 0x1F600)) },
            new[] { PdfStandardFont.Helvetica }));

        byte[] bytes = PdfDocument.Create(options).Paragraph(paragraph => paragraph.Text("A😀B")).ToBytes();
        AssertRegistrationFonts(bytes, "DocumentPrimary-Regular", "NewEmoji-Regular");
        Assert.Contains("A😀B", PdfReadDocument.Open(bytes).ExtractText(), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PdfDocument_NestedPanelRetainsFallbackResolvedAfterFirstChild(bool followingFallback) {
        var options = new PdfOptions { CompressContentStreams = false }
            .RegisterFontFamily(PdfStandardFont.Helvetica, new PdfEmbeddedFontFamily("Document Primary", CreateRegistrationPrimaryFont()))
            .RegisterEmbeddedFontFallbacks(new PdfEmbeddedFontFallbackSet(
                new[] { new PdfEmbeddedFontFallbackCandidate("Panel Emoji", ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(' ', 0x1F600)) },
                new[] { PdfStandardFont.Helvetica }));

        byte[] bytes = PdfDocument.Create(options)
            .Panel(content => content
                .Paragraph(paragraph => paragraph.Text("A"))
                .Panel(nested => nested
                    .Paragraph(paragraph => paragraph.Text("A"))
                    .Paragraph(paragraph => paragraph.Text("A😀B")),
                    new PdfPanelStyle { KeepTogether = false, KeepWithNext = false }),
                new PdfPanelStyle { KeepTogether = false, KeepWithNext = false })
            .Paragraph(paragraph => paragraph.Text(followingFallback ? "B😀A" : "B"))
            .ToBytes();

        AssertRegistrationFonts(bytes, "DocumentPrimary-Regular", "PanelEmoji-Regular");
        string text = PdfReadDocument.Open(bytes).ExtractText();
        Assert.Contains("A😀B", text, StringComparison.Ordinal);
        Assert.Contains(followingFallback ? "B😀A" : "B", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfDocument_IdenticalExplicitFaceRegistrationClaimsPreviousFallbackFamily() {
        byte[] symbols = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(' ', 0x2605);
        var options = new PdfOptions { CompressContentStreams = false }
            .RegisterFontFamily(PdfStandardFont.Helvetica, new PdfEmbeddedFontFamily("Primary", CreateRegistrationPrimaryFont()))
            .RegisterEmbeddedFontFallbacks(new PdfEmbeddedFontFallbackSet(
                new[] { new PdfEmbeddedFontFallbackCandidate("Old Symbols", symbols) },
                new[] { PdfStandardFont.TimesRoman }))
            .EmbedStandardFont(PdfStandardFont.TimesRoman, symbols, "Old Symbols-Regular")
            .RegisterEmbeddedFontFallbacks(CreateRegistrationEmojiFallbacks());

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.Font(PdfStandardFont.TimesRoman).Text("★"))
            .Paragraph(paragraph => paragraph.Text("A😀B"))
            .ToBytes();
        AssertRegistrationFonts(bytes, "OldSymbols-Regular", "EmojiFallback-Regular");
        Assert.Contains("★", PdfReadDocument.Open(bytes).ExtractText(), StringComparison.Ordinal);
        Assert.Contains("A😀B", PdfReadDocument.Open(bytes).ExtractText(), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PdfDocument_FallbackResolutionPreservesCallerFamilyWithIdenticalFontBytes(bool callerRegisteredAfterFallback) {
        byte[] emoji = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(' ', 0x1F600);
        var options = new PdfOptions { CompressContentStreams = false }
            .RegisterFontFamily(PdfStandardFont.Helvetica, new PdfEmbeddedFontFamily("Primary", CreateRegistrationPrimaryFont()));
        var caller = new PdfEmbeddedFontFamily("Caller Emoji", emoji, italic: emoji);
        var fallback = new PdfEmbeddedFontFallbackSet(
            new[] { new PdfEmbeddedFontFallbackCandidate("Emoji Fallback", emoji) },
            new[] { PdfStandardFont.TimesRoman });
        if (!callerRegisteredAfterFallback) options.RegisterFontFamily(PdfStandardFont.TimesRoman, caller);
        options.RegisterEmbeddedFontFallbacks(fallback);
        if (callerRegisteredAfterFallback) options.RegisterFontFamily(PdfStandardFont.TimesRoman, caller);

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.Runs(new[] { new PdfTextRun("A😀B", italic: true) }))
            .Paragraph(paragraph => paragraph.Runs(new[] { new PdfTextRun("😀", italic: true, font: PdfStandardFont.TimesRoman) }))
            .ToBytes();

        AssertRegistrationFonts(bytes, "CallerEmoji-Italic", "EmojiFallback-Italic");
        Assert.Contains("A😀B", PdfReadDocument.Open(bytes).ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void PdfDocument_ReplacingFallbacksReclaimsOwnedSlotsOutsideTheNewRequestedMapping() {
        var options = new PdfOptions { CompressContentStreams = false }
            .RegisterFontFamily(PdfStandardFont.Helvetica, new PdfEmbeddedFontFamily("Document Primary", CreateRegistrationPrimaryFont()))
            .RegisterFontFamily(PdfStandardFont.Courier, new PdfEmbeddedFontFamily("Document Mono", CreateRegistrationPrimaryFont()))
            .RegisterEmbeddedFontFallbacks(new PdfEmbeddedFontFallbackSet(
                new[] { new PdfEmbeddedFontFallbackCandidate("Old Symbols", ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(' ', 0x2605)) },
                new[] { PdfStandardFont.TimesRoman }))
            .RegisterEmbeddedFontFallbacks(new PdfEmbeddedFontFallbackSet(
                new[] { new PdfEmbeddedFontFallbackCandidate("New Emoji", ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(' ', 0x1F600)) },
                new[] { PdfStandardFont.Helvetica }));

        byte[] bytes = PdfDocument.Create(options).Paragraph(paragraph => paragraph.Text("A😀B")).ToBytes();
        AssertRegistrationFonts(bytes, "DocumentPrimary-Regular", "NewEmoji-Regular");
        Assert.Contains("A😀B", PdfReadDocument.Open(bytes).ExtractText(), StringComparison.Ordinal);
    }
}
