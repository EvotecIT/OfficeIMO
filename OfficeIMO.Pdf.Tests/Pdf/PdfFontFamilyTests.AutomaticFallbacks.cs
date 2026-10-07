using System;
using System.Linq;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfFontFamilyTests {
    [Theory]
    [InlineData(0)]
    [InlineData(63)]
    public void PdfOptions_ResettingAutomaticFallbacksReclaimsNamedResources(int callerFamilyCount) {
        byte[] font = CreateRegistrationPrimaryFont();
        var options = new PdfOptions();
        foreach (PdfStandardFont slot in new[] { PdfStandardFont.Helvetica, PdfStandardFont.TimesRoman, PdfStandardFont.Courier }) {
            options.RegisterFontFamily(slot, new PdfEmbeddedFontFamily("Caller " + slot, font));
        }
        for (int index = 0; index < callerFamilyCount; index++) {
            options.RegisterNamedFontFamily(new PdfEmbeddedFontFamily("Caller " + index, font));
        }
        var candidates = new[] { new PdfEmbeddedFontFallbackCandidate("Automatic", font) };
        for (int index = 0; index <= PdfOptions.MaximumNamedFontFamilies; index++) {
            Assert.True(options.RegisterAutomaticFontFallbackCandidates(candidates, Array.Empty<PdfStandardFont>()));
            options.EmbeddedFontFallbacks = null;
            Assert.Equal(callerFamilyCount, options.NamedFontFamilies.Count);
        }
        Assert.Contains("AB", PdfReadDocument.Open(PdfDocument.Create(options).Paragraph(paragraph => paragraph.Text("AB")).ToBytes()).ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void PdfOptions_ReplacingAutomaticFallbacksReclaimsCapacityOnlyAfterValidation() {
        byte[] font = CreateRegistrationPrimaryFont();
        var options = new PdfOptions();
        foreach (PdfStandardFont slot in new[] { PdfStandardFont.Helvetica, PdfStandardFont.TimesRoman, PdfStandardFont.Courier }) {
            options.RegisterFontFamily(slot, new PdfEmbeddedFontFamily("Caller " + slot, font));
        }
        for (int index = 0; index < PdfOptions.MaximumNamedFontFamilies - 1; index++) {
            options.RegisterNamedFontFamily(new PdfEmbeddedFontFamily("Caller " + index, font));
        }
        options.RegisterAutomaticFontFallbackCandidates(
            new[] { new PdfEmbeddedFontFallbackCandidate("Automatic", font) }, Array.Empty<PdfStandardFont>());
        string automaticName = Assert.Single(options.EmbeddedFontFallbacks!.FontFamilyNames);
        var tooMany = new PdfEmbeddedFontFallbackSet(new[] {
            new PdfEmbeddedFontFallbackCandidate("Replacement 1", font),
            new PdfEmbeddedFontFallbackCandidate("Replacement 2", font)
        });
        Assert.Throws<InvalidOperationException>(() => options.RegisterEmbeddedFontFallbacks(tooMany));
        Assert.Equal(automaticName, Assert.Single(options.EmbeddedFontFallbacks!.FontFamilyNames));
        Assert.Equal(PdfOptions.MaximumNamedFontFamilies, options.NamedFontFamilies.Count);
        Assert.False(options.HasNamedFontFamily("Replacement 1"));

        options.RegisterEmbeddedFontFallbacks(new PdfEmbeddedFontFallbackSet(
            new[] { new PdfEmbeddedFontFallbackCandidate("Replacement", font) }));
        Assert.False(options.HasNamedFontFamily(automaticName));
        Assert.Equal(PdfOptions.MaximumNamedFontFamilies, options.NamedFontFamilies.Count);
        options.EmbeddedFontFallbacks = null;
        Assert.True(options.HasNamedFontFamily("Replacement"));
        Assert.Contains("AB", PdfReadDocument.Open(PdfDocument.Create(options).Paragraph(paragraph => paragraph.Text("AB")).ToBytes()).ExtractText(), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PdfOptions_ResettingAutomaticFallbacksPreservesLaterCallerFamilyWithSameNameAndData(bool styled) {
        byte[] font = CreateRegistrationPrimaryFont();
        byte[] bold = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(' ', 'Q');
        var options = new PdfOptions();
        options.RegisterAutomaticFontFallbackCandidates(
            new[] { new PdfEmbeddedFontFallbackCandidate("Automatic", font) },
            new[] { PdfStandardFont.Helvetica, PdfStandardFont.TimesRoman, PdfStandardFont.Courier });
        string name = Assert.Single(options.EmbeddedFontFallbacks!.FontFamilyNames);
        options.RegisterNamedFontFamily(new PdfEmbeddedFontFamily(name, font, bold: styled ? bold : null));
        options.EmbeddedFontFallbacks = null;

        Assert.Equal(name, Assert.Single(options.NamedFontFamilies).Key);
        byte[] bytes = PdfDocument.Create(options).Paragraph(paragraph => paragraph.Runs(new[] {
            new PdfTextRun(styled ? "Q" : "AB", bold: styled, fontFamily: name)
        })).ToBytes();
        Assert.Contains(styled ? "Q" : "AB", PdfReadDocument.Open(bytes).ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void PdfOptions_ClonedAutomaticFallbackOwnershipIsIndependent() {
        var options = new PdfOptions();
        options.RegisterAutomaticFontFallbackCandidates(
            new[] { new PdfEmbeddedFontFallbackCandidate("Automatic", ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(' ', 0x25B8)) },
            new[] { PdfStandardFont.Helvetica, PdfStandardFont.TimesRoman, PdfStandardFont.Courier });
        PdfOptions clone = options.Clone();
        options.EmbeddedFontFallbacks = null;
        Assert.Empty(options.NamedFontFamilies);
        Assert.Contains("A▸B", PdfReadDocument.Open(PdfDocument.Create(clone).Paragraph(paragraph => paragraph.Text("A▸B")).ToBytes()).ExtractText(), StringComparison.Ordinal);
        clone.EmbeddedFontFallbacks = null;
        Assert.Empty(clone.NamedFontFamilies);
    }

    [Theory]
    [InlineData(1)]
    [InlineData(3)]
    public void PdfOptions_AutomaticFallbacksKeepAllTextCoverageWhenCompatibilitySlotsAreOccupied(int occupiedSlots) {
        byte[] primary = CreateRegistrationPrimaryFont();
        var options = new PdfOptions();
        foreach (PdfStandardFont slot in new[] { PdfStandardFont.Helvetica, PdfStandardFont.TimesRoman, PdfStandardFont.Courier }.Take(occupiedSlots)) {
            options.RegisterFontFamily(slot, new PdfEmbeddedFontFamily("Caller " + slot, primary));
        }
        var candidates = new[] {
            new PdfEmbeddedFontFallbackCandidate("Latin", primary),
            new PdfEmbeddedFontFallbackCandidate("Polish", ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(' ', 0x0119)),
            new PdfEmbeddedFontFallbackCandidate("Symbols", ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(' ', 0x25B8))
        };

        Assert.True(options.RegisterAutomaticFontFallbackCandidates(candidates, Array.Empty<PdfStandardFont>()));
        byte[] bytes = PdfDocument.Create(options).Paragraph(paragraph => paragraph.Text("Aę▸B")).ToBytes();

        Assert.Contains("Aę▸B", PdfReadDocument.Open(bytes).ExtractText(), StringComparison.Ordinal);
        foreach (PdfStandardFont slot in new[] { PdfStandardFont.Helvetica, PdfStandardFont.TimesRoman, PdfStandardFont.Courier }.Take(occupiedSlots)) {
            Assert.Equal(primary, options.EmbeddedFonts[slot].Data);
        }
    }

    [Fact]
    public void PdfOptions_AutomaticNamedFallbacksPreserveCallerStyledFacesAndCollidingResourceNames() {
        byte[] primary = CreateRegistrationPrimaryFont();
        byte[] symbols = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(' ', 0x25B8);
        byte[] bold = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(' ', 'Q');
        var options = new PdfOptions()
            .RegisterFontFamily(PdfStandardFont.Helvetica, new PdfEmbeddedFontFamily("Caller Default", primary))
            .RegisterFontFamily(PdfStandardFont.TimesRoman, new PdfEmbeddedFontFamily("Caller Serif", primary))
            .RegisterFontFamily(PdfStandardFont.Courier, new PdfEmbeddedFontFamily("Caller Mono", primary))
            .RegisterNamedFontFamily(new PdfEmbeddedFontFamily("Symbols", symbols, bold: bold))
            .RegisterNamedFontFamily(new PdfEmbeddedFontFamily("Symbols [compatibility Helvetica]", symbols, bold: bold));

        Assert.True(options.RegisterAutomaticFontFallbackCandidates(
            new[] { new PdfEmbeddedFontFallbackCandidate("Symbols", symbols) }, Array.Empty<PdfStandardFont>()));
        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.Runs(new[] { new PdfTextRun("Q", bold: true, fontFamily: "Symbols") }))
            .Paragraph(paragraph => paragraph.Text("A▸B"))
            .ToBytes();

        Assert.Contains("Q", PdfReadDocument.Open(bytes).ExtractText(), StringComparison.Ordinal);
        Assert.Contains("A▸B", PdfReadDocument.Open(bytes).ExtractText(), StringComparison.Ordinal);
        Assert.Equal(bold, options.NamedFontFamilies["Symbols"].Bold);
        Assert.Equal(bold, options.NamedFontFamilies["Symbols [compatibility Helvetica]"].Bold);
    }

    [Fact]
    public void PdfOptions_AutomaticFallbacksRespectNamedFamilyBudgetWithoutPartialRegistration() {
        byte[] font = CreateRegistrationPrimaryFont();
        var options = new PdfOptions();
        foreach (PdfStandardFont slot in new[] { PdfStandardFont.Helvetica, PdfStandardFont.TimesRoman, PdfStandardFont.Courier }) {
            options.RegisterFontFamily(slot, new PdfEmbeddedFontFamily("Caller " + slot, font));
        }
        for (int index = 0; index < PdfOptions.MaximumNamedFontFamilies; index++) {
            options.RegisterNamedFontFamily(new PdfEmbeddedFontFamily("Caller " + index, font));
        }

        Assert.Throws<InvalidOperationException>(() => options.RegisterAutomaticFontFallbackCandidates(
            new[] { new PdfEmbeddedFontFallbackCandidate("Automatic", font) }, Array.Empty<PdfStandardFont>()));

        Assert.Null(options.EmbeddedFontFallbacks);
        Assert.Equal(PdfOptions.MaximumNamedFontFamilies, options.NamedFontFamilies.Count);
        Assert.Contains("AB", PdfReadDocument.Open(PdfDocument.Create(options).Paragraph(paragraph => paragraph.Text("AB")).ToBytes()).ExtractText(), StringComparison.Ordinal);
    }
}
