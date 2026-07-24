using OfficeIMO.Drawing;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class VisioAllSeverityBatch20SecurityTests {
    [Fact]
    public void ResizeToTextWithoutMaximumWidthPreservesAuthoredSpaces() {
        var compact = new VisioShape("compact", 2D, 2D, 0.2D, 0.2D, "A B");
        var spaced = new VisioShape("spaced", 2D, 2D, 0.2D, 0.2D, "A   B");
        var font = new OfficeFontInfo("Aptos", 12D);

        compact.ResizeToText(font, horizontalPadding: 0D, verticalPadding: 0D, minimumWidth: 0D, minimumHeight: 0D);
        spaced.ResizeToText(font, horizontalPadding: 0D, verticalPadding: 0D, minimumWidth: 0D, minimumHeight: 0D);

        Assert.True(spaced.Width > compact.Width);
    }
}
