using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherWrapEditingTests {
    [Theory]
    [InlineData("reverse")]
    [InlineData("rectangle")]
    [InlineData("picture")]
    [InlineData("swap")]
    [InlineData("restore")]
    public void Square_wrapped_ordering_rebuilds_native_and_model_relationships(string caseName) {
        var (control, order, expected) = Case(caseName);
        byte[] source = Source(control); var d = PublisherDocument.Load(source); var before = d.Pages[0];
        uint[] oldWraps = Frame(d).WrapObjectIds.ToArray();
        var stories = d.TextStories.Select(s => (s.Id, s.Text)).ToArray();
        d.SetPageObjectOrder(before.Id, order);
        Assert.Equal(expected, Frame(d).WrapObjectIds); Assert.Equal(order, d.Pages[0].ObjectIds);
        Assert.Equal(oldWraps, before.TextFrames.Single(f => f.Id == 293).WrapObjectIds);
        Assert.Throws<OfficeConversionException>(() => d.ToPubBytes());
        var result = d.ToPubBytesResult(Allow());
        Assert.False(result.Report.HasErrors);
        Assert.Contains(result.Report.FidelityDiagnostics, item => item.Code == "PUB_NATIVE_WRAP_REFERENCES_UNQUALIFIED"
            && item.LossKind == OfficeConversionLossKind.Unassessed);
        var loaded = PublisherDocument.Load(result.Value);
        Assert.Equal(expected, Frame(loaded).WrapObjectIds); Assert.Equal(order, loaded.Pages[0].ObjectIds);
        Assert.Equal(stories, loaded.TextStories.Select(s => (s.Id, s.Text)).ToArray());
        Assert.Equal(d.ToSvg(), loaded.ToSvg());
        PublisherNativeWritingTests.AssertPreservedStorage(source, result.Value, "Contents", "Escher/EscherStm");
        Assert.Equal(result.Value, loaded.ToPubBytes()); Assert.Equal(result.Value, d.ToPubBytes(Allow()));
    }

    [Fact]
    public void Moving_square_wrapped_shape_and_picture_rebuilds_membership_and_can_return_to_exact_source() {
        byte[] source = Source("WrapObjectOrderOutside"); var d = PublisherDocument.Load(source); var before = d.Pages[0];
        Assert.Empty(Frame(d).WrapObjectIds);
        d.SetShapeBounds(294, 180, 132, 72, 72); Assert.Equal(new uint[] { 294 }, Frame(d).WrapObjectIds);
        d.SetPictureBounds(295, 144, 264, 96, 96); Assert.Equal(new uint[] { 294, 295 }, Frame(d).WrapObjectIds);
        var loaded = PublisherDocument.Load(d.ToPubBytes(Allow()));
        Assert.Equal(new uint[] { 294, 295 }, Frame(loaded).WrapObjectIds);
        Assert.Equal(d.ToSvg(), loaded.ToSvg());
        Assert.Empty(before.TextFrames.Single(f => f.Id == 293).WrapObjectIds);
        d.SetShapeBounds(294, 480, 132, 72, 72); Assert.Equal(new uint[] { 295 }, Frame(d).WrapObjectIds);
        d.SetPictureBounds(295, 480, 264, 96, 96); Assert.Empty(Frame(d).WrapObjectIds);
        Assert.Equal(source, d.ToPubBytes());
    }

    [Fact]
    public void Moving_the_text_frame_rebuilds_exclusions_without_changing_complete_stories() {
        var d = PublisherDocument.Load(Source("WrapObjectOrder")); var before = Frame(d);
        var stories = d.TextStories.Select(s => (s.Id, s.Text)).ToArray();
        d.SetTextFrameBounds(293, 480, 72, 100, 504); Assert.Empty(Frame(d).WrapObjectIds);
        var loaded = PublisherDocument.Load(d.ToPubBytes(Allow())); Assert.Empty(Frame(loaded).WrapObjectIds);
        Assert.Equal(stories, loaded.TextStories.Select(s => (s.Id, s.Text)).ToArray());
        Assert.Equal(new uint[] { 294, 295 }, before.WrapObjectIds);
        d.SetTextFrameBounds(293, 72, 72, 360, 504); Assert.Equal(new uint[] { 294, 295 }, Frame(d).WrapObjectIds);
    }

    [Fact]
    public void Wrap_reference_growth_composes_with_native_frame_layout_growth() {
        var d = PublisherDocument.Load(Source("WrapObjectOrderReversed"));
        d.SetTextFrameLayout(293, 2, 12, new OfficeTextPadding(4, 5, 6, 7), OfficeTextVerticalAlignment.Center);
        d.SetPageObjectOrder(d.Pages[0].Id, new uint[] { 293, 294, 295 });
        var loaded = PublisherDocument.Load(d.ToPubBytes(Allow()));
        Assert.Equal(new uint[] { 294, 295 }, Frame(loaded).WrapObjectIds);
        Assert.Equal(2, Frame(loaded).ColumnCount); Assert.Equal(12, Frame(loaded).ColumnSpacing);
        Assert.Equal(new OfficeTextPadding(4, 5, 6, 7), Frame(loaded).Padding);
        Assert.Equal(OfficeTextVerticalAlignment.Center, Frame(loaded).VerticalAlignment);
        Assert.Equal(d.ToSvg(), loaded.ToSvg());
    }

    [Fact]
    public void Native_default_rectangular_wrap_modes_remove_and_restore_exclusions() {
        var d = PublisherDocument.Load(Source("WrapObjectOrderDefaults"));
        Assert.Equal(new uint[] { 294, 295 }, Frame(d).WrapObjectIds);
        d.SetPageObjectOrder(d.Pages[0].Id, new uint[] { 295, 294, 293 }); Assert.Empty(Frame(d).WrapObjectIds);
        var loaded = PublisherDocument.Load(d.ToPubBytes(Allow())); Assert.Empty(Frame(loaded).WrapObjectIds);
        loaded.SetPageObjectOrder(loaded.Pages[0].Id, new uint[] { 293, 294, 295 });
        Assert.Equal(new uint[] { 294, 295 }, Frame(loaded).WrapObjectIds);
        Assert.Equal(new uint[] { 294, 295 }, Frame(PublisherDocument.Load(loaded.ToPubBytes(Allow()))).WrapObjectIds);
    }

    [Fact]
    public void Unsupported_native_wrap_side_refuses_order_changes_before_model_commit() {
        byte[] source = Source("WrapObjectOrderLeft"); var d = PublisherDocument.Load(source); var page = d.Pages[0];
        Assert.Equal(new uint[] { 294, 295 }, Frame(d).WrapObjectIds);
        Assert.Throws<NotSupportedException>(() => d.SetPageObjectOrder(page.Id, new uint[] { 295, 294, 293 }));
        Assert.Same(page, d.Pages[0]); Assert.Equal(source, d.ToPubBytes());
    }

    private static PublisherTextFrame Frame(PublisherDocument d) => d.Pages[0].TextFrames.Single(f => f.Id == 293);
    private static PublisherSaveOptions Allow() => new() { LossPolicy = OfficeConversionLossPolicy.Allow };
    private static byte[] Source(string control) => File.ReadAllBytes(PublisherNativeTests.Fixture(
        "NativeComparison/" + control + "/native-" + control.ToLowerInvariant() + ".pub"));

    private static (string Source, uint[] Order, uint[] Wraps) Case(string name) => name switch {
        "reverse" => ("WrapObjectOrder", new uint[] { 295, 294, 293 }, Array.Empty<uint>()),
        "rectangle" => ("WrapObjectOrder", new uint[] { 295, 293, 294 }, new uint[] { 294 }),
        "picture" => ("WrapObjectOrder", new uint[] { 294, 293, 295 }, new uint[] { 295 }),
        "swap" => ("WrapObjectOrder", new uint[] { 293, 295, 294 }, new uint[] { 295, 294 }),
        "restore" => ("WrapObjectOrderReversed", new uint[] { 293, 294, 295 }, new uint[] { 294, 295 }),
        _ => throw new ArgumentOutOfRangeException(nameof(name))
    };
}
