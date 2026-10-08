using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Visio.Tests;

public sealed class VisioDrawingOptionsSnapshotTests {
    [Fact]
    public void OperationSnapshotsRetainRuntimeFontContextInDetachedCollections() {
        var options = new VisioDrawingOptions();
        Func<OfficeFontProgramLoadRequest, IReadOnlyDictionary<string, float>?> resolver = _ =>
            new Dictionary<string, float> { ["opsz"] = 18 };
        options.Fonts.FontVariationResolver = resolver;
        VisioDrawingOptions copy = options.Clone();
        Assert.NotSame(options.Fonts, copy.Fonts);
        Assert.Same(resolver, copy.Fonts.FontVariationResolver);
        options.Fonts.FontVariationResolver = null;
        Assert.Same(resolver, copy.Fonts.FontVariationResolver);
    }
}
