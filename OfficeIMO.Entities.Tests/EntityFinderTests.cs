using System;
using System.Threading;

namespace OfficeIMO.Entities.Tests;

public sealed class EntityFinderTests {
    [Fact]
    public void PeselRequiresContextAtTheDefaultThreshold() {
        const string unlabelled = "The reference is 44051401359.";
        var finder = new EntityFinder();
        Assert.Empty(finder.Find(unlabelled).Matches);
        var permissive = new EntityFinder(options: new EntityFinderOptions { MinimumConfidence = 0.30 });
        EntityMatch weak = Assert.Single(permissive.Find(unlabelled).Matches);
        Assert.Same(PolandSchemes.Pesel, weak.Scheme);
        Assert.Equal(0.40, weak.Confidence);

        EntityMatch labelled = Assert.Single(finder.Find("PESEL: 44051401359").Matches);
        Assert.Same(PolandSchemes.Pesel, labelled.Scheme);
        Assert.Equal(0.80, labelled.Confidence);
        Assert.Equal("PESEL", labelled.ContextKeyword);
        Assert.Equal("pl", labelled.ContextLanguage);
        Assert.True(labelled.Evidence.HasFlag(EntityEvidence.PlausibleDate));
        Assert.Equal("NationalId/PL.PESEL [7..18] 0.80", labelled.ToString());
        Assert.Empty(finder.Find("PESEL: 44051401358").Matches);
    }

    [Fact]
    public void OnlyWholeWordsAndAllowedLanguagesSupportCandidates() {
        var finder = new EntityFinder();
        Assert.Empty(finder.Find("antypesel: 44051401359").Matches);
        Assert.Empty(finder.Find("PESEL was compared with 44051401359").Matches);
        Assert.Empty(finder.Find("NIP: " + new string('.', 49) + " 1234563218").Matches);
        Assert.Empty(new EntityFinder(options: new EntityFinderOptions { Languages = new[] { "de" } })
            .Find("PESEL: 44051401359").Matches);

        EntityMatch krs = Assert.Single(finder.Find("KRAJOWY REJESTR SADOWY nr: 1").Matches);
        Assert.Same(PolandSchemes.Krs, krs.Scheme);
        Assert.Equal("0000000001", krs.Normalized);
        Assert.Equal("Krajowy Rejestr Sądowy", krs.ContextKeyword);
    }

    [Fact]
    public void WholeKeywordBeforeConnectorStillCountsWhenConnectorContainsIt() {
        EntityCatalog catalog = new EntityCatalogBuilder().AddPoland()
            .AddKeyword("PL.PESEL", "es", "de")
            .Build();
        EntityMatch match = Assert.Single(new EntityFinder(catalog).Find("de del 44051401359").Matches);
        Assert.Equal("de", match.ContextKeyword);
        Assert.Equal("es", match.ContextLanguage);
    }

    [Fact]
    public void ValidationOnlyVatRegistrationDoesNotSuppressPrefixedNipRecognition() {
        EntityCatalog catalog = new EntityCatalogBuilder().AddPoland().Add(EuVatSchemes.Get("PL")).Build();
        EntityMatch match = Assert.Single(new EntityFinder(catalog).Find("PL5260001246").Matches);
        Assert.Same(PolandSchemes.Nip, match.Scheme);
        Assert.Equal(0.50, match.Confidence);
    }

    [Fact]
    public void SegmentLabelAndStructuredSourceScoreWithoutChangingSourceOffsets() {
        var location = new EntitySegmentLocation("Sheet1", "B7");
        var result = new EntityFinder().Find(new[] {
            new EntityTextSegment("PESEL:"),
            new EntityTextSegment("44051401359"),
            new EntityTextSegment("1234563218", location, "NIP"),
            new EntityTextSegment("44051401359", structuredSource: true),
        });
        Assert.Collection(result.Matches, nip => {
            Assert.Same(PolandSchemes.Nip, nip.Scheme);
            Assert.Same(location, nip.Location);
            Assert.Equal(2, nip.SegmentIndex);
            Assert.Equal(0, nip.Start);
            Assert.Equal(0.90, nip.Confidence);
            Assert.True(nip.Evidence.HasFlag(EntityEvidence.Label));
        }, pesel => {
            Assert.Same(PolandSchemes.Pesel, pesel.Scheme);
            Assert.Equal(1, pesel.Confidence);
            Assert.True(pesel.Evidence.HasFlag(EntityEvidence.StructuredSource));
        });
    }

    [Fact]
    public void StructuredEvidenceRequiresTheEntireMarkedSourceSegment() {
        var finder = new EntityFinder();
        EntityMatch complete = Assert.Single(finder.Find(new[] {
            new EntityTextSegment("Note: "),
            new EntityTextSegment("44051401359", structuredSource: true, continuesPrevious: true),
        }).Matches);
        Assert.Equal(1, complete.Confidence);
        Assert.True(complete.Evidence.HasFlag(EntityEvidence.StructuredSource));
        Assert.Equal(1, Assert.Single(complete.Parts).SegmentIndex);

        var crossesBoundary = finder.Find(new[] {
            new EntityTextSegment("Note: 4405"),
            new EntityTextSegment("1401359", structuredSource: true, continuesPrevious: true),
        });
        Assert.Empty(crossesBoundary.Matches);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void SplitPeselRequiresOptInContinuation(bool continuesPrevious) {
        var locations = new[] { new EntitySegmentLocation("first"), new EntitySegmentLocation("second") };
        var result = new EntityFinder().Find(new[] {
            new EntityTextSegment("PESEL: 4405", locations[0]),
            new EntityTextSegment("1401359", locations[1], continuesPrevious: continuesPrevious),
        });
        if (!continuesPrevious) {
            Assert.Empty(result.Matches);
            return;
        }
        EntityMatch match = Assert.Single(result.Matches);
        Assert.Same(PolandSchemes.Pesel, match.Scheme);
        Assert.Equal("44051401359", match.Normalized);
        Assert.Equal("PESEL", match.ContextKeyword);
        Assert.Equal(0, match.SegmentIndex);
        Assert.Same(locations[0], match.Location);
        Assert.Equal((7, 4), (match.Start, match.Length));
        Assert.Collection(match.Parts,
            first => Assert.Equal((0, 7, 4), (first.SegmentIndex, first.Start, first.Length)),
            second => Assert.Equal((1, 0, 7), (second.SegmentIndex, second.Start, second.Length)));
    }

    [Fact]
    public void MatchInThirdInputSegmentUsesThatSegmentsCoordinates() {
        var result = new EntityFinder().Find(new[] {
            new EntityTextSegment("earlier"),
            new EntityTextSegment(""),
            new EntityTextSegment("NIP 5260001246"),
        });
        EntityMatch match = Assert.Single(result.Matches);
        Assert.Same(PolandSchemes.Nip, match.Scheme);
        Assert.Equal((2, 4, 10), (match.SegmentIndex, match.Start, match.Length));
        EntityMatchPart part = Assert.Single(match.Parts);
        Assert.Equal((2, 4, 10), (part.SegmentIndex, part.Start, part.Length));
    }

    [Theory]
    [InlineData("NIP 123-456-32-18 99", "1234563218")]
    [InlineData("NIP 123 456 32 18/99", "1234563218")]
    [InlineData("NIP 123 456 32 18 99", null)]
    [InlineData("NIP 123456321899", null)]
    public void NumericCandidatesUseGroupEdgesAndRequireASeparatorFenceOrPrintFormat(string text, string? expected) {
        var finder = new EntityFinder(options: new EntityFinderOptions { IncludeSchemes = { "PL.NIP" } });
        var result = finder.Find(text);
        if (expected == null) Assert.Empty(result.Matches);
        else Assert.Equal(expected, Assert.Single(result.Matches).Normalized);
    }

    [Fact]
    public void IssuerFormattingStrengthensTheReading() {
        var finder = new EntityFinder();
        EntityMatch nip = Assert.Single(finder.Find("NIP 123-456-32-18").Matches);
        Assert.Same(PolandSchemes.Nip, nip.Scheme);
        Assert.True(nip.Evidence.HasFlag(EntityEvidence.IssuerFormatting));
        Assert.Equal(0.90, nip.Confidence);
        Assert.Same(PolandSchemes.Nip, Assert.Single(finder.Find("NIP 5260001246").Matches).Scheme);

        EntityMatch ssn = Assert.Single(finder.Find("SSN 123-45-6789").Matches);
        Assert.Same(UnitedStatesSchemes.Ssn, ssn.Scheme);
        Assert.Equal(0.70, ssn.Confidence);
        Assert.True(ssn.Evidence.HasFlag(EntityEvidence.IssuerFormatting));
        Assert.Empty(finder.Find("123-45-6789").Matches);
    }

    [Fact]
    public void DeclaredNipAndRegonContainmentPreservesInnerOffsetsAndProvenance() {
        var finder = new EntityFinder();
        EntityMatch vat = Assert.Single(finder.Find("VAT: PL1234563218").Matches);
        Assert.Same(EuVatSchemes.Get("PL"), vat.Scheme);
        EntityMatch nip = Assert.Single(vat.Contained);
        Assert.Same(PolandSchemes.Nip, nip.Scheme);
        Assert.Equal(vat.Start + 2, nip.Start);
        Assert.Equal(vat.Length - 2, nip.Length);

        EntityMatch regon = Assert.Single(finder.Find("REGON 123456785").Matches);
        Assert.Equal(0.80, regon.Confidence);
        Assert.Equal(EntityCheckOutcome.Passed, regon.CheckDigit);
        Assert.Equal(EntityCheckProvenance.CommunityDocumented, regon.CheckProvenance);
        EntityMatch branch = Assert.Single(finder.Find("REGON 12345678500010").Matches);
        EntityMatch parent = Assert.Single(branch.Contained);
        Assert.Equal("123456785", parent.Normalized);
        Assert.Equal("123456785", PolandSchemes.Regon.Decode(branch.Normalized)!.Properties["parent"]);
    }

    [Theory]
    [InlineData("ATU10223006", "EU.VAT.AT")]
    [InlineData("DE111111125", "EU.VAT.DE")]
    [InlineData("GB434031494", "GB.VAT")]
    [InlineData("BE0776091951", "EU.VAT.BE")]
    [InlineData("FR40303265045", "EU.VAT.FR")]
    [InlineData("ES12345678Z", "EU.VAT.ES")]
    public void PrefixedVatUsesStrongEvidenceWhenTheCheckPasses(string written, string schemeId) {
        EntityMatch match = Assert.Single(new EntityFinder().Find(written).Matches);
        Assert.Equal(schemeId, match.Scheme.Id);
        Assert.Equal(0.85, match.Confidence);
        Assert.True(match.Evidence.HasFlag(EntityEvidence.CountryPrefix));
        Assert.Equal(EntityCheckOutcome.Passed, match.CheckDigit);
        if (schemeId == "EU.VAT.FR") Assert.Equal(EntityCheckProvenance.CommunityDocumented, match.CheckProvenance);
    }

    [Theory]
    [InlineData("PESEL: 12344051401359432109")]
    [InlineData("PESEL: 12.05.2023")]
    [InlineData("KRS: 2023-05-12")]
    [InlineData("KRS: 12/05/2023")]
    [InlineData("PESEL: 44051401359 USD")]
    [InlineData("PESEL: 44051401359.50")]
    [InlineData("PESEL: 44051401359,50")]
    [InlineData("PESEL: 44051401359€")]
    [InlineData("NIP: €1234563218")]
    [InlineData("card: 0000 0000 0000 0000")]
    [InlineData("PESEL: α44051401359")]
    [InlineData("PESEL: 44051401359β")]
    public void DigitGuardsRejectUnrelatedShapesEvenWhenTheLabelLooksRight(string text) {
        Assert.Empty(new EntityFinder(options: new EntityFinderOptions { MinimumConfidence = 0.30 }).Find(text).Matches);
    }

    [Fact]
    public void FindsAnIbanAndCardWithWrittenSpansAndFormattingEvidence() {
        const string iban = "PL61 1090 1014 0000 0712 1981 2874";
        const string text = "IBAN " + iban + "; card: 4111 1111 1111 1111.";
        Assert.Collection(new EntityFinder().Find(text).Matches, account => {
            Assert.Same(FinancialSchemes.Iban, account.Scheme);
            Assert.Equal(iban, text.Substring(account.Start, account.Length));
            Assert.Equal("IBAN", account.ContextKeyword);
            Assert.Equal("PL61109010140000071219812874", account.Normalized);
            Assert.True(account.Evidence.HasFlag(EntityEvidence.IssuerFormatting));
        }, card => {
            Assert.Same(FinancialSchemes.Card, card.Scheme);
            Assert.Equal(0.90, card.Confidence);
            Assert.True(card.Evidence.HasFlag(EntityEvidence.KnownIssuerRange));
        });
    }

    [Fact]
    public void ExactMatchLimitReadsOnlyOneSegmentBeyondTheCompletedChain() {
        static System.Collections.Generic.IEnumerable<EntityTextSegment> Segments() {
            yield return new EntityTextSegment("first@example.com");
            yield return new EntityTextSegment("second@example.com");
            throw new InvalidOperationException("The finder requested another segment after reaching its limit.");
        }

        EntityFindResult result = new EntityFinder(options: new EntityFinderOptions { MaxMatches = 1 }).Find(Segments());
        Assert.Single(result.Matches);
        Assert.True(result.Truncated);
    }

    [Fact]
    public void IbanRecognizerAcceptsCompactBankPrefixFollowedByGroups() {
        EntityMatch match = Assert.Single(new EntityFinder().Find("GB82WEST 1234 5698 7654 32").Matches);
        Assert.Same(FinancialSchemes.Iban, match.Scheme);
        Assert.Equal("GB82WEST12345698765432", match.Normalized);
    }

    [Fact]
    public void RepeatedRegonParentIsExcludedFromProseFindings() {
        // The parent is all zeroes. The first thirteen digits of the branch sum to 8 modulo 11.
        EntityMatch regon = Assert.Single(new EntityFinder().Find("REGON 00000000000018").Matches);
        Assert.Same(PolandSchemes.Regon, regon.Scheme);
        Assert.Empty(regon.Contained);
    }

    [Fact]
    public void CharacterLimitKeepsACompleteSegmentAndRejectsAnUnsearchedContinuation() {
        const string pesel = "PESEL: 44051401359";
        var finder = new EntityFinder(options: new EntityFinderOptions { MaxInputCharacters = pesel.Length });
        Assert.Single(finder.Find(pesel).Matches);
        var clipped = finder.Find(new[] {
            new EntityTextSegment(pesel),
            new EntityTextSegment("9", continuesPrevious: true),
        });
        Assert.True(clipped.Truncated);
        Assert.Empty(clipped.Matches);
    }

    [Fact]
    public void FiltersAndLimitsApplyAcrossSegmentsAndKeepDiagnosticsValueFree() {
        var options = new EntityFinderOptions { Countries = { "PL" }, MaxMatches = 1 };
        var finder = new EntityFinder(options: options);
        options.ExcludeSchemes.Add("EMAIL"); // Construction takes a snapshot.
        EntityFindResult result = finder.Find(new[] {
            new EntityTextSegment("DE111111125"),
            new EntityTextSegment("first@example.com; second@example.com"),
        });
        Assert.Equal("first@example.com", Assert.Single(result.Matches).Normalized);
        Assert.True(result.Truncated);
        Assert.DoesNotContain("example.com", string.Join(" ", result.Diagnostics));

        var clipped = new EntityFinder(options: new EntityFinderOptions { MaxInputCharacters = 18 });
        Assert.True(clipped.Find("PESEL: 440514013599").Truncated);
        Assert.Empty(clipped.Find("PESEL: 440514013599").Matches);
        var clippedAmount = new EntityFinder(options: new EntityFinderOptions { MaxInputCharacters = 19 });
        Assert.Empty(clippedAmount.Find("PESEL: 44051401359,50").Matches);
        Assert.Empty(clippedAmount.Find("PESEL: 44051401359-9").Matches);
        Assert.Empty(clippedAmount.Find("PESEL: 44051401359/9").Matches);
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => finder.Find("first@example.com", cancellation.Token));
    }
}
