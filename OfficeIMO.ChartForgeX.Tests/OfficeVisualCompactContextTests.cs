using global::ChartForgeX.Core;
using OfficeIMO.ChartForgeX;
using Xunit;

namespace OfficeIMO.ChartForgeX.Tests;

public class OfficeVisualCompactContextTests {
    [Theory]
    [InlineData(64D, 18D, 6D)]
    [InlineData(18D, 64D, 6D)]
    [InlineData(120D, 16D, 16D / 3D)]
    [InlineData(450D, 270D, 24D)]
    public void DocumentContextPreparesCompactSparklinesWithoutExhaustingContent(double width, double height, double padding) {
        var context = OfficeVisualDocumentStyle.Default.CreateContext(width, height);
        Assert.Equal(padding, context.Layout.Padding);
        var prepared = new Sparkline(SparklineData.FromValues(new[] { 2D, 5D, 3D })).Prepare(context);
        Assert.NotNull(prepared);
        Assert.Equal(width / .75D, context.Layout.Size.Width);
        Assert.Equal(height / .75D, context.Layout.Size.Height);
    }
}
