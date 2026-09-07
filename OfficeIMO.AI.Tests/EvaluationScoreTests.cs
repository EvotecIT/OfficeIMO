using OfficeIMO.Reader;
using Xunit;

namespace OfficeIMO.AI.Tests;

public sealed class EvaluationScoreTests {
    [Fact]
    public void TableScoringPenalizesReorderedAndAdditionalCells() {
        var gold = new EvaluationGold(Table: new(new[] { "Code", "Count" }, new[] { (IReadOnlyList<string>)new[] { "A", "17" } }));
        OfficeAiResult correct = Result(new[] { "A", "17" });
        Assert.True(gold.Score(correct).Passed);
        Assert.True(gold.Score(correct).SemanticReviewRequired);
        EvaluationScore swapped = gold.Score(Result(new[] { "17", "A" }));
        Assert.False(swapped.Passed);
        Assert.Equal(0.5, swapped.TableCellRecall);
        Assert.Equal(0.5, swapped.TableCellPrecision);
        EvaluationScore duplicate = gold.Score(correct with { Tables = new[] { correct.Tables[0], correct.Tables[0] } });
        Assert.False(duplicate.Passed);
        Assert.Equal(1, duplicate.TableCellRecall);
        Assert.Equal(0.5, duplicate.TableCellPrecision);
    }

    [Fact]
    public void ScalarExtractionStillRequiresSemanticReviewAfterExactValueMatch() {
        var gold = new EvaluationGold(Fields: new[] { new EvaluationFieldGold("count", OfficeAiFieldStatus.Present, "17") });
        var result = Result(new[] { "A", "17" }) with { Tables = Array.Empty<OfficeAiTable>(), Fields = new[] {
            new OfficeAiField("count", OfficeAiFieldType.Integer, OfficeAiFieldStatus.Present, "17", "17", Array.Empty<OfficeAiCitation>())
        } };
        Assert.True(gold.Score(result).Passed);
        Assert.True(gold.Score(result).SemanticReviewRequired);
    }

    [Fact]
    public void FactMarkersAreExplicitlySeparateFromSemanticReview() {
        var gold = new EvaluationGold(FactMarkers: new[] { "17" });
        var result = Result(new[] { "A", "17" }) with { Claims = new[] {
            new OfficeAiClaim("This sentence mentions 17 but does not prove its meaning.", Array.Empty<OfficeAiCitation>())
        } };
        EvaluationScore score = gold.Score(result);
        Assert.Equal(1, score.FactMarkerRecall);
        Assert.True(score.SemanticReviewRequired);
    }

    private static OfficeAiResult Result(string[] cells) => new() {
        RequestId = "test", SourceHash = "test", SnapshotHash = "test", Status = OfficeAiResultStatus.Completed,
        Profile = new() { Id = "test", Provider = "test", Model = "test", IsLocal = true },
        Tables = new[] { new OfficeAiTable(new ReaderTable { Columns = new[] { "Code", "Count" }, Rows = new[] { cells } }, Array.Empty<OfficeAiCitation>()) }
    };
}
