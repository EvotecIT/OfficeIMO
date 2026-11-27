using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Entities.Tests;

public sealed class EntityOverlapTests {
    [Fact]
    public void SameSpanKeepsAlternativesAndContextRemovesTheAmbiguityPenaltyForTheSupportedReading() {
        var first = new SampleScheme("test.first");
        var second = new SampleScheme("test.second");
        var catalog = new EntityCatalogBuilder().Add(first).Add(second)
            .AddKeyword(second.Id, "en", "verified")
            .AddRecognizer(new SampleRecognizer(
                (first, 10, 5, EntityCandidateStrength.DistinctiveShape),
                (second, 10, 5, EntityCandidateStrength.DistinctiveShape))).Build();
        EntityMatch unlabelled = Assert.Single(new EntityFinder(catalog).Find("Witness:  value").Matches);
        Assert.Same(first, unlabelled.Scheme);
        Assert.Equal(0.65, unlabelled.Confidence);
        Assert.True(unlabelled.Evidence.HasFlag(EntityEvidence.AmbiguousReading));
        Assert.Same(second, Assert.Single(unlabelled.Alternatives).Scheme);

        EntityMatch labelled = Assert.Single(new EntityFinder(catalog).Find("Verified: value").Matches);
        Assert.Same(second, labelled.Scheme);
        Assert.Equal(1, labelled.Confidence);
        Assert.Equal(0.65, Assert.Single(labelled.Alternatives).Confidence);
    }

    [Fact]
    public void PartialOverlapFavorsConfidenceAndThenLengthAndAllKeepsEachReading() {
        var scheme = new SampleScheme("test");
        EntityCatalog catalog = new EntityCatalogBuilder().Add(scheme)
            .AddRecognizer(new SampleRecognizer(
                (scheme, 0, 6, EntityCandidateStrength.DistinctiveShape),
                (scheme, 4, 3, EntityCandidateStrength.StrongCheck),
                (scheme, 6, 5, EntityCandidateStrength.StrongCheck))).Build();
        EntityFindResult best = new EntityFinder(catalog).Find("abcdefghijk");
        Assert.Equal(1, best.DroppedOverlaps);
        Assert.Equal(new[] { 0, 6 }, best.Matches.Select(match => match.Start));
        Assert.Equal(new[] { 6, 5 }, best.Matches.Select(match => match.Length));

        var all = new EntityFinder(catalog, new EntityFinderOptions { OverlapPolicy = EntityOverlapPolicy.All });
        Assert.Equal(new[] { 0, 4, 6 }, all.Find("abcdefghijk").Matches.Select(match => match.Start));
    }

    [Fact]
    public void EligibleContainedReadingSurvivesWhenItsContainerLosesAnUnrelatedOverlap() {
        var other = new SampleScheme("sample");
        EntityCatalog catalog = new EntityCatalogBuilder().AddPoland().AddEuVat().Add(other)
            .AddRecognizer(new SampleRecognizer((other, 0, 6, EntityCandidateStrength.DistinctiveShape)))
            .Build();
        var finder = new EntityFinder(catalog, new EntityFinderOptions { MinimumConfidence = 0.40 });

        EntityFindResult result = finder.Find(new[] { new EntityTextSegment("ABC-PL1234563218", label: other.Id) });
        Assert.Equal(1, result.DroppedOverlaps);
        Assert.Collection(result.Matches,
            reading => Assert.Same(other, reading.Scheme),
            reading => {
                Assert.Same(PolandSchemes.Nip, reading.Scheme);
                Assert.Equal(6, reading.Start);
            });
    }

    [Fact]
    public void MaskerReplacesTheUnionOfOverlappingAndAdjacentSpansOnce() {
        const string text = "abcdefghijk!";
        var scheme = new SampleScheme("test");
        EntityCatalog catalog = new EntityCatalogBuilder().Add(scheme)
            .AddRecognizer(new SampleRecognizer(
                (scheme, 0, 3, EntityCandidateStrength.StrongCheck),
                (scheme, 2, 4, EntityCandidateStrength.StrongCheck),
                (scheme, 6, 5, EntityCandidateStrength.StrongCheck))).Build();
        var all = new EntityFinder(catalog, new EntityFinderOptions { OverlapPolicy = EntityOverlapPolicy.All });
        var matches = all.Find(text).Matches;
        Assert.Equal("[Email]!", EntityMasker.Mask(text, matches));
        Assert.Equal("***********!", EntityMasker.Mask(text, matches,
            new EntityMaskOptions { Style = EntityMaskStyle.Fixed, FixedCharacter = '*' }));

        EntityPseudonymResult pseudonym = EntityMasker.Pseudonymize(text, matches,
            new EntityMaskOptions { PseudonymKey = new byte[] { 11, 12, 13 } });
        var replacement = Assert.Single(pseudonym.Map);
        Assert.Equal("abcdefghijk", replacement.Value);
        Assert.Equal(replacement.Key + "!", pseudonym.Text);
    }

    private sealed class SampleScheme : EntityScheme {
        internal SampleScheme(string id) => Id = id;
        public override string Id { get; }
        public override EntityKind Kind => new("Email");
        public override string DisplayName => Id;
        public override EntityValidationResult Validate(string? value) =>
            new(this, value == null ? EntityValidationStatus.Empty : EntityValidationStatus.Valid, value);
    }

    private sealed class SampleRecognizer : IEntityRecognizer {
        private readonly (EntityScheme Scheme, int Start, int Length, EntityCandidateStrength Strength)[] _readings;

        internal SampleRecognizer(params (EntityScheme Scheme, int Start, int Length, EntityCandidateStrength Strength)[] readings) {
            _readings = readings;
            Schemes = readings.Select(reading => reading.Scheme).Distinct().ToArray();
        }

        public string Id => "test-recognizer";
        public IReadOnlyList<EntityScheme> Schemes { get; }
        public EntityTokenShape Shapes => EntityTokenShape.AlphaNumericRun;

        public void Recognize(EntityRecognitionContext context) {
            foreach (var reading in _readings) {
                var result = reading.Scheme.Validate(context.Text.Substring(reading.Start, reading.Length));
                context.Add(new EntityCandidate(result, reading.Start, reading.Length, reading.Strength));
            }
        }
    }
}
