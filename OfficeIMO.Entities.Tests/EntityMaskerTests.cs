namespace OfficeIMO.Entities.Tests;

public sealed class EntityMaskerTests {
    [Fact]
    public void LabelAndFixedMaskingUseWrittenOffsets() {
        const string text = "PESEL: 44051401359; https://example.com";
        var matches = new EntityFinder().Find(text).Matches;
        Assert.Equal("PESEL: [NationalId]; https://example.com", EntityMasker.Mask(text, matches));
        Assert.Equal("PESEL: ***********; https://example.com",
            EntityMasker.Mask(text, matches, new EntityMaskOptions { Style = EntityMaskStyle.Fixed, FixedCharacter = '*' }));
        Assert.Equal("PESEL: ***; https://example.com",
            EntityMasker.Mask(text, matches, new EntityMaskOptions { Style = EntityMaskStyle.Fixed, FixedCharacter = '*', FixedLength = 3 }));
    }

    [Theory]
    [InlineData("IBAN PL61 1090 1014 0000 0712 1981 2874", "IBAN PL61 •••• •••• •••• •••• •••• 2874")]
    [InlineData("card 4111 1111 1111 1111", "card •••• •••• •••• 1111")]
    [InlineData("Email: jan.kowalski@example.com", "Email: j•••@example.com")]
    [InlineData("PESEL: 44051401359", "PESEL: •••••••••••")]
    public void PartialMaskShowsOnlySchemeApprovedFragments(string text, string expected) {
        var matches = new EntityFinder().Find(text).Matches;
        Assert.Single(matches);
        Assert.Equal(expected, EntityMasker.Mask(text, matches, new EntityMaskOptions { Style = EntityMaskStyle.Partial }));
    }

    [Theory]
    [InlineData(EntityMaskStyle.KindLabel, "PESEL: [NationalId]", "")]
    [InlineData(EntityMaskStyle.Fixed, "PESEL: ****", "*******")]
    [InlineData(EntityMaskStyle.Partial, "PESEL: ••••", "•••••••")]
    [InlineData(EntityMaskStyle.Pseudonym, "PESEL: [NationalId:5a19ed4883930bdc44b02fe3f4eb7e6e]", "")]
    public void EveryStyleMasksAllPartsOfASplitPesel(EntityMaskStyle style, string first, string second) {
        var segments = new[] {
            new EntityTextSegment("PESEL: 4405"),
            new EntityTextSegment("1401359", continuesPrevious: true),
        };
        var matches = new EntityFinder().Find(segments).Matches;
        Assert.Equal(2, Assert.Single(matches).Parts.Count);

        var options = new EntityMaskOptions {
            Style = style, FixedCharacter = '*', FixedLength = 2, PseudonymKey = new byte[] { 11, 12, 13, 14 },
        };
        Assert.Equal(new[] { first, second }, EntityMasker.Mask(segments, matches, options));
    }

    [Theory]
    [InlineData("IBAN PL61 1090 1014 0000 0712 1981 28", "74", "IBAN PL61 •••• •••• •••• •••• •••• 28", "74")]
    [InlineData("Email: jan.kowalski@", "EXAMPLE.COM", "Email: j•••@", "example.com")]
    public void SplitPartialMaskKeepsPermittedCharactersInTheirSourceSegments(
        string first, string second, string expectedFirst, string expectedSecond) {
        var segments = new[] {
            new EntityTextSegment(first),
            new EntityTextSegment(second, continuesPrevious: true),
        };
        var matches = new EntityFinder().Find(segments).Matches;
        Assert.Equal(2, Assert.Single(matches).Parts.Count);
        Assert.Equal(new[] { expectedFirst, expectedSecond },
            EntityMasker.Mask(segments, matches, new EntityMaskOptions { Style = EntityMaskStyle.Partial }));
    }

    [Fact]
    public void IndependentSegmentsKeepSeparateTokensAtTheirEdges() {
        var segments = new[] {
            new EntityTextSegment("alice@example.com"),
            new EntityTextSegment("bob@example.com"),
        };
        var matches = new EntityFinder().Find(segments).Matches;
        Assert.Equal(new[] { "[Email]", "[Email]" }, EntityMasker.Mask(segments, matches));
    }

    [Fact]
    public void DefaultMaskingProtectsAnEmailInsideAPublicMailtoUrl() {
        const string text = "mailto:Alice@example.com";
        var matches = new EntityFinder().Find(text).Matches;
        Assert.Same(InternetSchemes.Url, Assert.Single(matches).Scheme);
        Assert.Equal("mailto:[Email]", EntityMasker.Mask(text, matches));
    }

    [Fact]
    public void PseudonymForOverlappingReadingsUsesTheirCanonicalValues() {
        var finder = new EntityFinder(options: new EntityFinderOptions { OverlapPolicy = EntityOverlapPolicy.All });
        var options = new EntityMaskOptions {
            PseudonymKey = new byte[] { 11, 12, 13 },
            Sensitivities = { EntitySensitivity.Public },
        };
        const string first = "mailto:Alice@EXAMPLE.COM";
        const string second = "MAILTO:Alice@example.com";

        EntityPseudonymResult original = EntityMasker.Pseudonymize(first, finder.Find(first).Matches, options);
        EntityPseudonymResult equivalent = EntityMasker.Pseudonymize(second, finder.Find(second).Matches, options);
        Assert.Equal(original.Text, equivalent.Text);
        Assert.Equal(first, Assert.Single(original.Map).Value);
        Assert.Equal(second, Assert.Single(equivalent.Map).Value);
    }

    [Fact]
    public void PseudonymUsesTheCanonicalValueAndANondefaultKey() {
        const string text = "CARD 4111 1111 1111 1111; CARD 4111111111111111";
        var matches = new EntityFinder().Find(text).Matches;
        var options = new EntityMaskOptions { Style = EntityMaskStyle.Pseudonym, PseudonymKey = new byte[] { 11, 12, 13, 14 } };
        EntityPseudonymResult result = EntityMasker.Pseudonymize(text, matches, options);
        var token = Assert.Single(result.Map);
        Assert.Equal("4111111111111111", token.Value);
        Assert.Equal("CARD " + token.Key + "; CARD " + token.Key, result.Text);
        Assert.Equal(result.Text, EntityMasker.Mask(text, matches, options));

        options.PseudonymKey = new byte[] { 21, 22, 23, 24 };
        Assert.NotEqual(token.Key, Assert.Single(EntityMasker.Pseudonymize(text, matches, options).Map.Keys));
        options.PseudonymKey = null;
        Assert.Throws<System.ArgumentException>(() => EntityMasker.Pseudonymize(text, matches, options));
    }
}
