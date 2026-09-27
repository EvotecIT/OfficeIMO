using OfficeIMO.Markup;
using OfficeIMO.Markup.Word;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests.Markup;

public sealed class OfficeMarkupWordChartTests {
    [Theory]
    [InlineData("pie", WordChartSnapshotKind.Pie)]
    [InlineData("doughnut", WordChartSnapshotKind.Doughnut)]
    [InlineData("donut", WordChartSnapshotKind.Doughnut)]
    public void WordChart_PreservesTheRequestedNativeFamily(string type, WordChartSnapshotKind expected) {
        string markup = "::chart type=" + type + " title=Status\nCategory,Value\nPass,8\nUnknown,2\nFail,1\n";
        var parsed = OfficeMarkupParser.Parse(markup);
        Assert.False(parsed.HasErrors);
        using WordDocument document = parsed.Document.ToWordDocument();
        Assert.True(Assert.Single(document.Charts).TryGetSnapshot(out WordChartSnapshot snapshot));
        Assert.Equal(expected, snapshot.ChartKind);
        Assert.Equal(new[] { "Pass", "Unknown", "Fail" }, snapshot.Data.Categories);
        Assert.Equal(new double[] { 8, 2, 1 }, snapshot.Data.Series[0].Values);
        Assert.Empty(document.ValidateDocument());
    }
}
