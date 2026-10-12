using OfficeIMO.Drawing;
using System.Xml.Linq;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherGradientFillTests {
    [Theory]
    [InlineData(4U, 0, 0.5, 1, 0.5, 0)]
    [InlineData(4U, -90, 0, 0.5, 1, 0.5)]
    [InlineData(7U, 0, 0.5, 1, 0.5, 0)]
    [InlineData(7U, -90, 0, 0.5, 1, 0.5)]
    [InlineData(7U, -45, 0, 1, 1, 0)]
    public void Native_linear_fill_preserves_angle_and_endpoint_colors(uint type, int angle,
        double x1, double y1, double x2, double y2) {
        PublisherDocument document = Load(type, angle);
        OfficeShape shape = Artwork(document);
        Assert.NotNull(shape.FillGradient);
        OfficeLinearGradient gradient = shape.FillGradient!;
        Assert.Equal(x1, gradient.StartX, 6); Assert.Equal(y1, gradient.StartY, 6);
        Assert.Equal(x2, gradient.EndX, 6); Assert.Equal(y2, gradient.EndY, 6);
        Assert.Equal(new[] { OfficeColor.Blue, OfficeColor.Red }, gradient.Stops.Select(stop => stop.Color));
        Assert.DoesNotContain(document.ReadReport.FidelityDiagnostics, item => item.Code == "PUB_FILL_APPROXIMATED");
        XElement svg = XElement.Parse(document.ToSvg());
        Assert.Contains(svg.Descendants(), item => item.Name.LocalName == "linearGradient");
    }

    [Theory]
    [InlineData(80, 0.2, false)]
    [InlineData(-20, 0.2, true)]
    public void Focus_preserves_the_native_interior_color_and_reflected_ramp(int focus, double middle, bool reverse) {
        OfficeLinearGradient gradient = Artwork(Load(4, focus: focus)).FillGradient!;
        Assert.NotNull(gradient);
        Assert.Equal(new[] { 0, middle, 1 }, gradient.Stops.Select(stop => stop.Offset));
        OfficeColor edge = reverse ? OfficeColor.Red : OfficeColor.Blue;
        OfficeColor center = reverse ? OfficeColor.Blue : OfficeColor.Red;
        Assert.Equal(new[] { edge, center, edge }, gradient.Stops.Select(stop => stop.Color));
    }

    [Fact]
    public void Foreground_and_background_opacity_are_carried_by_stops_once() {
        OfficeShape shape = Artwork(Load(4, extra: new() { [0x182] = 32768, [0x184] = 0 }));
        Assert.NotNull(shape.FillGradient);
        Assert.Equal(255, shape.FillGradient!.Stops[0].Color.A);
        Assert.Equal(0, shape.FillGradient.Stops[1].Color.A);
        Assert.Equal(0.5D, shape.FillOpacity);
    }

    [Fact]
    public void Unrepresentable_stop_alpha_ratios_have_explicit_loss_evidence() {
        PublisherDocument document = Load(4, extra: new() { [0x182] = 32768, [0x184] = 65536 });
        OfficeShape shape = Artwork(document);
        Assert.NotNull(shape.FillGradient);
        Assert.Equal(128, shape.FillGradient!.Stops[0].Color.A);
        Assert.Equal(1D, shape.FillOpacity);
        Assert.Contains(document.ReadReport.FidelityDiagnostics, item => item.Code == "PUB_GRADIENT_OPACITY_APPROXIMATED");
    }

    [Theory]
    [InlineData(32768U, null, true)]
    [InlineData(32768U, 65536U, true)]
    [InlineData(null, 32768U, true)]
    [InlineData(65536U, 32768U, true)]
    [InlineData(null, null, false)]
    [InlineData(65536U, 65536U, false)]
    [InlineData(32768U, 32768U, false)]
    public void Multi_color_opacity_evidence_uses_effective_native_defaults(uint? foreground, uint? background, bool expectedLoss) {
        var properties = new Dictionary<ushort, uint>();
        if (foreground.HasValue) properties[0x182] = foreground.Value;
        if (background.HasValue) properties[0x184] = background.Value;
        PublisherDocument document = Load(7, extra: properties,
            complex: Stops((0x00FF0000, 0), (0x0000FF00, 32768), (0x000000FF, 65536)));
        OfficeShape shape = Artwork(document);
        Assert.NotNull(shape.FillGradient);
        Assert.Equal((foreground ?? 65536U) / 65536D, shape.FillOpacity);
        Assert.All(shape.FillGradient!.Stops, stop => Assert.Equal(255, stop.Color.A));
        Assert.Equal(expectedLoss, document.ReadReport.FidelityDiagnostics.Any(item => item.Code == "PUB_GRADIENT_OPACITY_APPROXIMATED"));
    }

    [Theory]
    [InlineData(100)]
    [InlineData(-100)]
    public void Reversed_focus_retains_two_distinct_endpoints(int focus) {
        OfficeLinearGradient gradient = Artwork(Load(4, focus: focus)).FillGradient!;
        Assert.NotNull(gradient);
        Assert.Equal(new[] { 0D, 1D }, gradient.Stops.Select(stop => stop.Offset));
        Assert.Equal(new[] { OfficeColor.Red, OfficeColor.Blue }, gradient.Stops.Select(stop => stop.Color));
    }

    [Fact]
    public void Unscaled_native_angle_retains_its_physical_direction_on_a_wide_frame() {
        OfficeShape physical = Artwork(Load(4, -45));
        OfficeShape scaled = Artwork(Load(7, -45));
        Assert.NotNull(physical.FillGradient); Assert.NotNull(scaled.FillGradient);
        OfficeLinearGradient gradient = physical.FillGradient!;
        Assert.Equal(physical.Width / physical.Height,
            -(gradient.EndX - gradient.StartX) / (gradient.EndY - gradient.StartY), 6);
        Assert.Equal(1D, -(scaled.FillGradient!.EndX - scaled.FillGradient.StartX)
            / (scaled.FillGradient.EndY - scaled.FillGradient.StartY), 6);
    }

    [Fact]
    public void Complex_native_color_stops_preserve_positions_and_hard_transitions() {
        byte[] table = Stops((0x00FF0000, 0), (0x0000FF00, 32768), (0x000000FF, 32768), (0x000000FF, 65536));
        OfficeLinearGradient gradient = Artwork(Load(7, complex: table)).FillGradient!;
        Assert.NotNull(gradient);
        Assert.Equal(new[] { 0, 0.5, 0.5, 1 }, gradient.Stops.Select(stop => stop.Offset));
        Assert.Equal(new[] { OfficeColor.Blue, OfficeColor.Lime, OfficeColor.Red, OfficeColor.Red }, gradient.Stops.Select(stop => stop.Color));
    }

    [Fact]
    public void Focused_multi_color_stops_preserve_duplicate_endpoints_without_invalid_order() {
        byte[] table = Stops((0x00FF0000, 0), (0x0000FF00, 65536), (0x000000FF, 65536));
        OfficeLinearGradient gradient = Artwork(Load(7, focus: 80, complex: table)).FillGradient!;
        Assert.NotNull(gradient);
        Assert.Equal(new[] { 0, 0.2, 0.2, 0.2, 1 }, gradient.Stops.Select(stop => stop.Offset));
        Assert.Equal(new[] { OfficeColor.Blue, OfficeColor.Lime, OfficeColor.Red, OfficeColor.Lime, OfficeColor.Blue }, gradient.Stops.Select(stop => stop.Color));
    }

    [Theory]
    [InlineData(5U)]
    [InlineData(6U)]
    [InlineData(8U)]
    [InlineData(2U)]
    public void Unsupported_fill_modes_retain_primary_paint_and_precise_loss(uint mode) {
        PublisherDocument document = Load(mode);
        Assert.Null(Artwork(document).FillGradient);
        Assert.Contains(document.ReadReport.FidelityDiagnostics, item => item.Code == "PUB_FILL_APPROXIMATED"
            && item.LossKind == OfficeConversionLossKind.Approximation);
    }

    [Fact]
    public void Invalid_focus_and_color_arrays_preserve_recoverable_primary_paint() {
        foreach (PublisherDocument document in new[] { Load(4, focus: 101), Load(4, complex: new byte[] { 1, 0, 1, 0, 8, 0 }) }) {
            Assert.Null(Artwork(document).FillGradient);
            Assert.Contains(document.ReadReport.FidelityDiagnostics, item => item.Code == "PUB_GRADIENT_INVALID"
                && item.LossKind == OfficeConversionLossKind.Approximation);
        }
    }

    [Fact]
    public void Native_special_interpolation_is_reported_as_an_approximation() {
        PublisherDocument document = Load(4, extra: new() { [0x19C] = 0x40000003 });
        Assert.NotNull(Artwork(document).FillGradient);
        Assert.Contains(document.ReadReport.FidelityDiagnostics, item => item.Code == "PUB_GRADIENT_INTERPOLATION_APPROXIMATED"
            && item.LossKind == OfficeConversionLossKind.Approximation);
    }

    [Fact]
    public void Projected_gradient_expansion_honors_the_existing_item_work_limit() {
        byte[] table = Stops(Enumerable.Range(0, 64).Select(index => (0x000000FFU, (uint)(65536L * index / 63))).ToArray());
        byte[] input = Input(7, focus: 80, complex: table);
        var options = new PublisherReadOptions { Limits = new OfficeLegacyImportLimits { MaxItems = 100 } };
        InvalidDataException error = Assert.Throws<InvalidDataException>(() => PublisherDocument.Load(input, options));
        Assert.Contains("gradient stop work limit", error.Message);
    }

    [Fact]
    public void Fill_boolean_masks_control_visibility_and_custom_mapping() {
        PublisherDocument hidden = Load(4, extra: new() { [0x1BF] = 0x00100000 });
        Assert.DoesNotContain(hidden.ReadReport.FidelityDiagnostics, item => item.Code.StartsWith("PUB_GRADIENT", StringComparison.Ordinal));
        Assert.All(PublisherNativeTests.Elements(hidden.Pages[0].Drawing).OfType<OfficeDrawingShape>()
            .Where(item => item.SourceElementIds?.Contains("publisher-object-293") == true), item => Assert.Null(item.Shape.FillGradient));
        PublisherDocument custom = Load(4, extra: new() { [0x1BF] = 0x00120012 });
        Assert.Null(Artwork(custom).FillGradient);
        Assert.Contains(custom.ReadReport.FidelityDiagnostics, item => item.Code == "PUB_GRADIENT_MAPPING_UNSUPPORTED");
        PublisherDocument anchored = Load(4, extra: new() { [0x1BF] = 0x00140010 });
        Assert.Null(Artwork(anchored).FillGradient);
        Assert.Contains(anchored.ReadReport.FidelityDiagnostics, item => item.Code == "PUB_GRADIENT_MAPPING_UNSUPPORTED");
    }

    [Fact]
    public void Explicit_fill_rotation_is_distinct_from_unqualified_frame_following() {
        foreach (uint flags in new[] { 0x00100010U, 0x00300030U }) {
            PublisherDocument document = Load(4, extra: new() { [4] = 30 * 65536, [0x1BF] = flags });
            Assert.NotNull(Artwork(document).FillGradient);
            Assert.Equal(flags != 0x00300030U, document.ReadReport.FidelityDiagnostics.Any(item => item.Code == "PUB_GRADIENT_TRANSFORM_APPROXIMATED"));
        }
    }

    [Fact]
    public void Picture_fill_is_below_the_image_and_does_not_reappear_in_its_outline() {
        const string fixture = "SampleBrochure.pub";
        PublisherDocument original = PublisherDocument.Load(PublisherNativeTests.Fixture(fixture));
        OfficeDrawingImage picture = original.Pages.SelectMany(page => PublisherNativeTests.Elements(page.Drawing)).OfType<OfficeDrawingImage>().First();
        string sourceId = picture.SourceElementIds!.Single();
        uint id = uint.Parse(sourceId.Substring("publisher-object-".Length), System.Globalization.CultureInfo.InvariantCulture);
        var properties = new Dictionary<ushort, uint> {
            [0x180] = 7, [0x181] = 0x00FF0000, [0x183] = 0x000000FF, [0x19C] = 0,
            [0x1BF] = 0x00100010, [0x1C0] = 0, [0x1CB] = 12700, [0x1FF] = 0x00080008
        };
        PublisherDocument document = PublisherDocument.Load(PublisherDrawingFixture.Mutate(properties, 75, fixture: fixture, objectId: id));
        OfficeDrawingElement[] layers = document.Pages.SelectMany(page => PublisherNativeTests.Elements(page.Drawing))
            .Where(item => item.SourceElementIds?.Contains(sourceId) == true).ToArray();
        Assert.Collection(layers,
            item => Assert.NotNull(Assert.IsType<OfficeDrawingShape>(item).Shape.FillGradient),
            item => Assert.IsType<OfficeDrawingImage>(item),
            item => { Assert.Null(Assert.IsType<OfficeDrawingShape>(item).Shape.FillGradient); Assert.Null(((OfficeDrawingShape)item).Shape.FillColor); });
    }

    internal static PublisherDocument Load(uint type, int angle = 0, int focus = 0,
        Dictionary<ushort, uint>? extra = null, byte[]? complex = null) => PublisherDocument.Load(Input(type, angle, focus, extra, complex));

    internal static byte[] Input(uint type, int angle = 0, int focus = 0,
        Dictionary<ushort, uint>? extra = null, byte[]? complex = null) {
        var values = new Dictionary<ushort, uint> {
            [0x180] = type, [0x181] = 0x00FF0000, [0x183] = 0x000000FF,
            [0x18B] = unchecked((uint)(angle * 65536)), [0x18C] = unchecked((uint)focus),
            [0x19C] = 0, [0x1BF] = 0x00100010, [0x1FF] = 0x00080000
        };
        if (extra != null) foreach (var entry in extra) values[entry.Key] = entry.Value;
        return PublisherDrawingFixture.Mutate(values, 1, complex == null ? null : new() { [0x197] = complex });
    }

    internal static OfficeShape Artwork(PublisherDocument document) => PublisherNativeTests.Elements(document.Pages[0].Drawing)
        .OfType<OfficeDrawingShape>().Single(item => item.SourceElementIds?.Contains("publisher-object-293") == true).Shape;

    internal static byte[] Stops(params (uint Color, uint Position)[] entries) {
        using var stream = new MemoryStream(); using var writer = new BinaryWriter(stream);
        writer.Write((ushort)entries.Length); writer.Write((ushort)entries.Length); writer.Write((ushort)8);
        foreach (var entry in entries) { writer.Write(entry.Color); writer.Write(entry.Position); }
        return stream.ToArray();
    }
}
