using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherObjectOrderWritingTests {
    [Theory]
    [InlineData(false, 0)]
    [InlineData(false, 2)]
    [InlineData(false, -1)]
    [InlineData(true, -1)]
    public void Existing_page_object_order_saves_complete_records_without_rebinding_content(bool reversed, int selectedPage) {
        byte[] source = Source(reversed); var d = PublisherDocument.Load(source);
        var snapshots = d.Pages.ToArray(); var stories = d.TextStories.Select(s => (s.Id, s.Text)).ToArray();
        uint[][] normal = { new uint[] { 299, 300, 302, 303 }, Array.Empty<uint>(), new uint[] { 304, 305, 307, 308 } };
        uint[][] target = normal.Select((ids, i) => selectedPage == -1 || selectedPage == i
            ? (reversed ? ids : ids.Reverse()).ToArray() : (reversed ? ids.Reverse() : ids).ToArray()).ToArray();
        for (int i = 0; i < d.Pages.Count; i++) {
            if (selectedPage == -1 || selectedPage == i) d.SetPageObjectOrder(d.Pages[i].Id, target[i]);
            Assert.Equal(target[i], d.Pages[i].ObjectIds);
            Assert.Equal(target[i], PublisherMixedObjectOrderTests.Order(d.Pages[i]));
            Assert.Equal(reversed ? normal[i].Reverse() : normal[i], snapshots[i].ObjectIds);
        }
        Assert.Throws<OfficeConversionException>(() => d.ToPubBytes());
        var result = d.ToPubBytesResult(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
        Assert.False(result.Report.HasErrors); Assert.False(result.Report.IsExactSourceCopy);
        Assert.Contains(result.Report.FidelityDiagnostics, x => x.Code == "PUB_NATIVE_OBJECT_ORDER_UNQUALIFIED");
        var loaded = PublisherDocument.Load(result.Value);
        for (int i = 0; i < target.Length; i++) Assert.Equal(target[i], loaded.Pages[i].ObjectIds);
        Assert.Equal(stories, loaded.TextStories.Select(s => (s.Id, s.Text)).ToArray());
        PublisherNativeWritingTests.AssertPreservedStorage(source, result.Value);
        var original = ShapeRecords(source); var saved = ShapeRecords(result.Value);
        Assert.Equal(original.Keys.OrderBy(id => id), saved.Keys.OrderBy(id => id));
        foreach (var item in original) Assert.Equal(item.Value, saved[item.Key]);
        Assert.Equal(result.Value, loaded.ToPubBytes());
        Assert.Equal(result.Value, d.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
    }

    [Fact]
    public void Object_order_composes_with_native_frame_layout_growth() {
        var d = PublisherDocument.Load(Source(false)); var page = d.Pages[0];
        uint[] order = page.ObjectIds.Reverse().ToArray();
        PublisherTextFrame frame = page.TextFrames.Single(f => f.Id == 302);
        d.SetTextFrameLayout(frame.Id, 2, 12, new OfficeTextPadding(4, 5, 6, 7), OfficeTextVerticalAlignment.Center);
        d.SetPageObjectOrder(page.Id, order);
        var loaded = PublisherDocument.Load(d.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
        Assert.Equal(order, loaded.Pages[0].ObjectIds);
        var current = loaded.Pages[0].TextFrames.Single(f => f.Id == frame.Id);
        Assert.Equal(2, current.ColumnCount); Assert.Equal(12, current.ColumnSpacing);
        Assert.Equal(new OfficeTextPadding(4, 5, 6, 7), current.Padding);
        Assert.Equal(OfficeTextVerticalAlignment.Center, current.VerticalAlignment);
        Assert.Equal(frame.StoryId, current.StoryId);
        Assert.Equal(d.TextStories.Select(s => (s.Id, s.Text)), loaded.TextStories.Select(s => (s.Id, s.Text)));
    }

    [Theory]
    [InlineData("missing")]
    [InlineData("duplicate")]
    [InlineData("foreign")]
    public void Invalid_object_orders_leave_native_model_and_snapshots_unchanged(string kind) {
        byte[] source = Source(false); var d = PublisherDocument.Load(source); var page = d.Pages[0];
        uint[] order = page.ObjectIds.ToArray();
        if (kind == "missing") order = order.Take(order.Length - 1).ToArray();
        else order[0] = kind == "duplicate" ? order[1] : d.Pages[2].ObjectIds[0];
        Assert.Throws<ArgumentException>(() => d.SetPageObjectOrder(page.Id, order));
        Assert.Same(page, d.Pages[0]); Assert.Equal(source, d.ToPubBytes());
    }

    [Fact]
    public void Cancellation_and_invalid_page_do_not_publish_order_changes() {
        byte[] source = Source(false); var d = PublisherDocument.Load(source); var page = d.Pages[0];
        var order = page.ObjectIds.Reverse().ToArray();
        Assert.Throws<OperationCanceledException>(() => d.SetPageObjectOrder(page.Id, order, new CancellationToken(true)));
        Assert.Throws<ArgumentOutOfRangeException>(() => d.SetPageObjectOrder(uint.MaxValue, order));
        Assert.Throws<ArgumentNullException>(() => d.SetPageObjectOrder(page.Id, null!));
        d.SetPageObjectOrder(d.Pages[1].Id, Array.Empty<uint>());
        Assert.Equal(page.ObjectIds, d.Pages[0].ObjectIds); Assert.Equal(source, d.ToPubBytes());
    }

    [Fact]
    public void Authored_object_order_drives_the_scene_and_retains_old_snapshots() {
        var d = PublisherDocument.Create(300, 300); uint pageId = d.Pages[0].Id;
        var red = OfficeShape.Rectangle(100, 100); red.FillColor = OfficeColor.Red;
        var blue = OfficeShape.Rectangle(100, 100); blue.FillColor = OfficeColor.Blue;
        uint a = d.AddShape(pageId, red, 20, 20).Id, b = d.AddShape(pageId, blue, 40, 40).Id;
        var snapshot = d.Pages[0]; uint[] order = { b, a };
        d.SetPageObjectOrder(pageId, order); order[0] = a;
        Assert.Equal(new[] { b, a }, d.Pages[0].ObjectIds);
        Assert.Equal(new[] { b, a }, PublisherMixedObjectOrderTests.Order(d.Pages[0]));
        Assert.Equal(new[] { a, b }, PublisherMixedObjectOrderTests.Order(snapshot));
        Assert.Throws<OfficeConversionException>(() => d.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
    }

    private static byte[] Source(bool reversed) => File.ReadAllBytes(PublisherNativeTests.Fixture(reversed
        ? "NativeComparison/MixedObjectOrderReversed/native-mixedobjectorderreversed.pub"
        : "NativeComparison/MixedObjectOrder/native-mixedobjectorder.pub"));

    private static Dictionary<uint, byte[]> ShapeRecords(byte[] bytes) {
        Assert.True(OfficeCompoundFileReader.TryRead(bytes, out OfficeCompoundFile? compound, out string? error), error);
        byte[] stream = compound!.Streams["Escher/EscherStm"];
        var art = new PublisherEscherReader(new PublisherBinaryData(stream, "Escher/EscherStm"),
            compound.Streams.TryGetValue("Escher/EscherDelayStm", out byte[]? delay) ? new PublisherBinaryData(delay, "delay") : null,
            new PublisherParseContext(new PublisherReadOptions(), CancellationToken.None)).Read();
        return art.Shapes.ToDictionary(s => s.Id, s => stream.Skip(s.NativeBinding!.BodyOffset - 8).Take(s.NativeBinding.BodyLength + 8).ToArray());
    }
}
