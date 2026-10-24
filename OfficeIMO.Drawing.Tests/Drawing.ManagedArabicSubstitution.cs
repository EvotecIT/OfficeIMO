using System.Threading;
using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingManagedArabicSubstitutionTests {
    [Theory]
    [InlineData("ببب ب", new[] { 4, 6, 5, 7 })]
    [InlineData("بَب", new[] { 4, 5 })]
    [InlineData("ب\u200Dب", new[] { 4, 5 })]
    [InlineData("ب\u200Cب", new[] { 7, 7 })]
    public void ArabicJoiningUsesLogicalSourceMarksAndControls(string source, int[] expectedForms) {
        byte[] font = ManagedTextShapingTestAssets.CreateArabicJoiningFont();
        OfficeOpenTypeSubstitution substitution = Assert.IsType<OfficeOpenTypeSubstitution>(OfficeOpenTypeSubstitution.TryCreate(font));
        var tokens = new List<OfficeOpenTypeSubstitution.GlyphToken>();
        for (int index = 0; index < source.Length; index++) {
            char character = source[index];
            if (character is '\u200D' or '\u200C') continue;
            tokens.Add(new OfficeOpenTypeSubstitution.GlyphToken(character == 'ب' ? 1 : 2,
                character.ToString(), index, character));
        }

        Assert.True(substitution.TryApplyArabicDefaults(tokens, OfficeTextFeatureSettings.Default, "ar", default, source));
        Assert.Equal(expectedForms, tokens.Where(token => token.Scalar == 'ب').Select(token => token.GlyphId));
    }

    [Fact]
    public void ArabicRequiredLigatureRetainsLogicalSourceAcrossIgnoredMark() {
        byte[] font = ManagedTextShapingTestAssets.CreateFontWithLigature(0x0644, 0x0627,
            featureTag: "rlig", scriptTag: "arab", lookupFlags: 0x0009);
        OfficeOpenTypeSubstitution substitution = Assert.IsType<OfficeOpenTypeSubstitution>(OfficeOpenTypeSubstitution.TryCreate(font));
        var tokens = new List<OfficeOpenTypeSubstitution.GlyphToken> {
            new(1, "ل", 0, 0x0644), new(4, "َ", 1, 0x064E, clusterStart: 0), new(2, "ا", 2, 0x0627)
        };

        Assert.True(substitution.TryApplyArabicDefaults(tokens, OfficeTextFeatureSettings.Default,
            "ar", default, "لَا"));

        Assert.Equal(2, tokens.Count);
        Assert.Equal(3, tokens[0].GlyphId);
        Assert.Equal("لَا", tokens[0].UnicodeText);
        Assert.Equal(new[] { 0, 2 }, tokens[0].LigatureComponentTextIndexes);
        Assert.Equal(4, tokens[1].GlyphId);
        Assert.Empty(tokens[1].UnicodeText);
        Assert.True(tokens[1].IsUnicodeContinuation);
        Assert.Equal(1, tokens[1].TextIndex);
        Assert.Equal(0, tokens[1].ClusterStart);
        Assert.Equal("لَا", string.Concat(tokens.Select(token => token.UnicodeText)));
        Assert.NotNull(tokens[0].LigatureIdentity);
        Assert.Same(tokens[0].LigatureIdentity, tokens[1].LigatureIdentity);
        Assert.Equal(1, tokens[1].LigatureComponent);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void NormalizedIgnoredMarks_AssociateOnlyInsideTheMatchedLigatureInputs(bool inside) {
        byte[] font = ManagedTextShapingTestAssets.CreateFontWithLigature(0x0644, 0x0627,
            featureTag: "rlig", scriptTag: "arab", lookupFlags: 0x0009);
        OfficeOpenTypeSubstitution substitution = Assert.IsType<OfficeOpenTypeSubstitution>(OfficeOpenTypeSubstitution.TryCreate(font));
        var mark = new OfficeOpenTypeSubstitution.GlyphToken(4, string.Empty, inside ? 0 : 1,
            0x064E, isUnicodeContinuation: true);
        var tokens = inside ? new List<OfficeOpenTypeSubstitution.GlyphToken> {
            new(1, "لَ", 0, 0x0644), mark, new(2, "ا", 2, 0x0627)
        } : new List<OfficeOpenTypeSubstitution.GlyphToken> {
            new(1, "ل", 0, 0x0644), new(2, "اَ", 1, 0x0627), mark
        };
        string source = inside ? "لَا" : "لاَ";

        Assert.True(substitution.TryApplyArabicDefaults(tokens, OfficeTextFeatureSettings.Default,
            "ar", default, source));

        Assert.Equal(2, tokens.Count);
        Assert.NotNull(tokens[0].LigatureIdentity);
        if (inside) {
            Assert.Same(tokens[0].LigatureIdentity, tokens[1].LigatureIdentity);
            Assert.Equal(1, tokens[1].LigatureComponent);
        } else {
            Assert.Null(tokens[1].LigatureIdentity);
            Assert.Equal(0, tokens[1].LigatureComponent);
        }
        Assert.True(tokens[1].IsUnicodeContinuation);
        Assert.Equal(source, string.Concat(tokens.Select(token => token.UnicodeText)));
    }

    [Fact]
    public void ArabicFeaturesDoNotUseUnrelatedLatinScript() {
        byte[] font = ManagedTextShapingTestAssets.CreateFontWithLigature(0x0644, 0x0627,
            featureTag: "rlig", scriptTag: "latn");
        OfficeOpenTypeSubstitution substitution = Assert.IsType<OfficeOpenTypeSubstitution>(OfficeOpenTypeSubstitution.TryCreate(font));
        var tokens = new List<OfficeOpenTypeSubstitution.GlyphToken> { new(1, "ل", 0, 0x0644), new(2, "ا", 1, 0x0627) };

        Assert.False(substitution.TryApplyArabicDefaults(tokens, OfficeTextFeatureSettings.Default,
            "ar", default, "لا"));
        Assert.Equal(new[] { 1, 2 }, tokens.Select(token => token.GlyphId));
        Assert.Equal("لا", string.Concat(tokens.Select(token => token.UnicodeText)));
    }

    [Fact]
    public void NormalizedMarkContinuation_DoesNotReceiveItsBaseJoiningMask() {
        byte[] font = ManagedTextShapingTestAssets.CreateArabicJoiningFont();
        OfficeOpenTypeSubstitution substitution = Assert.IsType<OfficeOpenTypeSubstitution>(OfficeOpenTypeSubstitution.TryCreate(font));
        // The nominal glyph deliberately also appears in the joining coverage:
        // scalar/mark classification must guard the shared original cluster index.
        var tokens = new List<OfficeOpenTypeSubstitution.GlyphToken> {
            new(1, "بَ", 0, 0x0628), new(1, string.Empty, 0, 0x064E, isUnicodeContinuation: true),
            new(1, "ب", 2, 0x0628)
        };

        Assert.True(substitution.TryApplyArabicDefaults(tokens, OfficeTextFeatureSettings.Default,
            "ar", default, "بَب"));

        Assert.Equal(new[] { 4, 1, 5 }, tokens.Select(token => token.GlyphId));
        Assert.True(tokens[1].IsUnicodeContinuation);
        Assert.Empty(tokens[1].UnicodeText);
        Assert.Equal("بَب", string.Concat(tokens.Select(token => token.UnicodeText)));
    }

    [Fact]
    public void AuthoredDisabledRequiredLigatureFeatureRetainsSeparateGlyphs() {
        byte[] font = ManagedTextShapingTestAssets.CreateFontWithLigature(0x0644, 0x0627,
            featureTag: "rlig", scriptTag: "arab");
        OfficeOpenTypeSubstitution substitution = Assert.IsType<OfficeOpenTypeSubstitution>(OfficeOpenTypeSubstitution.TryCreate(font));
        var tokens = new List<OfficeOpenTypeSubstitution.GlyphToken> { new(1, "ل", 0, 0x0644), new(2, "ا", 1, 0x0627) };

        Assert.True(substitution.TryApplyArabicDefaults(tokens, OfficeTextFeatureSettings.Default.With("rlig", 0), "ar", default, "لا"));
        Assert.Equal(new[] { 1, 2 }, tokens.Select(token => token.GlyphId));
    }

    [Theory]
    [InlineData("\u200C")]
    [InlineData("\u200D")]
    public void RemovedJoiningControls_DoNotCreateAnArtificialRequiredLigature(string control) {
        byte[] font = ManagedTextShapingTestAssets.CreateFontWithLigature(0x0644, 0x0627,
            featureTag: "rlig", scriptTag: "arab");
        OfficeOpenTypeSubstitution substitution = Assert.IsType<OfficeOpenTypeSubstitution>(OfficeOpenTypeSubstitution.TryCreate(font));
        var tokens = new List<OfficeOpenTypeSubstitution.GlyphToken> {
            new(1, "ل", 0, 0x0644), new(2, "ا", 2, 0x0627)
        };

        Assert.True(substitution.TryApplyArabicDefaults(tokens, OfficeTextFeatureSettings.Default,
            "ar", default, "ل" + control + "ا"));

        Assert.Equal(new[] { 1, 2 }, tokens.Select(token => token.GlyphId));
        Assert.Equal(new[] { 0, 2 }, tokens.Select(token => token.TextIndex));
    }

    [Theory]
    [InlineData("\u200C")]
    [InlineData("\u200D")]
    public void RetainedSameClusterControl_PreservesOwnershipAndTheIgnoredMarkSide(string control) {
        byte[] font = ManagedTextShapingTestAssets.CreateFontWithLigature(0x0644, 0x0627,
            featureTag: "rlig", scriptTag: "arab", lookupFlags: 0x0009);
        OfficeOpenTypeSubstitution substitution = Assert.IsType<OfficeOpenTypeSubstitution>(OfficeOpenTypeSubstitution.TryCreate(font));
        string source = "لَ" + control + "ا";
        var tokens = new List<OfficeOpenTypeSubstitution.GlyphToken> {
            new(1, source.Substring(0, 3), 0, 0x0644),
            new(4, string.Empty, 0, 0x064E, isUnicodeContinuation: true),
            new(4, string.Empty, 0, control[0], isUnicodeContinuation: true),
            new(2, "ا", 3, 0x0627)
        };

        Assert.True(substitution.TryApplyArabicDefaults(tokens, OfficeTextFeatureSettings.Default,
            "ar", default, source));

        Assert.Equal(new[] { 1, 4, 4, 2 }, tokens.Select(token => token.GlyphId));
        Assert.Equal(new[] { 0x0644, 0x064E, (int)control[0], 0x0627 }, tokens.Select(token => token.Scalar));
        Assert.Equal(new[] { 0, 0, 0, 3 }, tokens.Select(token => token.TextIndex));
        Assert.Equal(source, string.Concat(tokens.Select(token => token.UnicodeText)));
        Assert.True(tokens[1].IsUnicodeContinuation);
        Assert.True(tokens[2].IsUnicodeContinuation);
    }

    [Theory]
    [InlineData("\u200C", 7, 7)]
    [InlineData("\u200D", 4, 5)]
    public void RetainedControlBarriers_PreserveFullSourceJoiningFormSelection(string control, int first, int second) {
        byte[] font = ManagedTextShapingTestAssets.CreateArabicJoiningFont();
        OfficeOpenTypeSubstitution substitution = Assert.IsType<OfficeOpenTypeSubstitution>(OfficeOpenTypeSubstitution.TryCreate(font));
        var tokens = new List<OfficeOpenTypeSubstitution.GlyphToken> {
            new(1, "ب" + control, 0, 0x0628),
            new(2, string.Empty, 0, control[0], isUnicodeContinuation: true), new(1, "ب", 2, 0x0628)
        };

        Assert.True(substitution.TryApplyArabicDefaults(tokens, OfficeTextFeatureSettings.Default,
            "ar", default, "ب" + control + "ب"));

        Assert.Equal(new[] { first, 2, second }, tokens.Select(token => token.GlyphId));
        Assert.Equal("ب" + control + "ب", string.Concat(tokens.Select(token => token.UnicodeText)));
    }

    [Fact]
    public void LatinMarkCompositionRetainsTheOriginalDecomposedLogicalText() {
        byte[] font = ManagedTextShapingTestAssets.CreateFontWithLigature('a', 0x0301,
            featureTag: "ccmp", scriptTag: "latn");
        OfficeOpenTypeSubstitution substitution = Assert.IsType<OfficeOpenTypeSubstitution>(OfficeOpenTypeSubstitution.TryCreate(font));
        var tokens = new List<OfficeOpenTypeSubstitution.GlyphToken> { new(1, "a", 0, 'a'), new(2, "\u0301", 1, 0x0301, clusterStart: 0) };

        Assert.True(substitution.TryApplyMarkScriptDefaults(tokens, OfficeTextFeatureSettings.Default, "latn", "en", default, "a\u0301"));
        OfficeOpenTypeSubstitution.GlyphToken composed = Assert.Single(tokens);
        Assert.Equal(3, composed.GlyphId);
        Assert.Equal("a\u0301", composed.UnicodeText);
        Assert.Equal(0, composed.ClusterStart);
    }

    [Fact]
    public void ArabicSubstitutionCancellationDoesNotPublishPartialTokens() {
        byte[] font = ManagedTextShapingTestAssets.CreateFontWithLigature(0x0644, 0x0627,
            featureTag: "rlig", scriptTag: "arab");
        OfficeOpenTypeSubstitution substitution = Assert.IsType<OfficeOpenTypeSubstitution>(OfficeOpenTypeSubstitution.TryCreate(font));
        var tokens = new List<OfficeOpenTypeSubstitution.GlyphToken> { new(1, "ل", 0, 0x0644), new(2, "ا", 1, 0x0627) };
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() => substitution.TryApplyArabicDefaults(tokens,
            OfficeTextFeatureSettings.Default, "ar", cancellation.Token, "لا"));
        Assert.Equal(new[] { 1, 2 }, tokens.Select(token => token.GlyphId));
    }
}
