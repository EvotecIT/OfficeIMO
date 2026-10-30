using OfficeIMO.Markup;
using OfficeIMO.Markup.Word;
using OfficeIMO.Word;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests.Markup;

public sealed class OfficeMarkupWordChartTests {
    [Theory]
    [InlineData("pie", OfficeChartType.Pie)]
    [InlineData("doughnut", OfficeChartType.Doughnut)]
    [InlineData("donut", OfficeChartType.Doughnut)]
    public void WordChart_PreservesTheRequestedNativeFamily(string type, OfficeChartType expected) {
        string markup = "::chart type=" + type + " title=Status\nCategory,Value\nPass,8\nUnknown,2\nFail,1\n";
        var parsed = OfficeMarkupParser.Parse(markup);
        Assert.False(parsed.HasErrors);
        using WordDocument document = parsed.Document.ToWordDocument();
        Assert.True(Assert.Single(document.Charts).TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
        Assert.Equal(expected, snapshot.ChartKind);
        Assert.Equal(new[] { "Pass", "Unknown", "Fail" }, snapshot.Data.Categories);
        Assert.Equal(new double[] { 8, 2, 1 }, snapshot.Data.Series[0].Values);
        Assert.Empty(document.ValidateDocument());
    }
}
