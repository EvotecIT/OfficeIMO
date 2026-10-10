using OfficeIMO.Drawing;
using OfficeIMO.Publisher.Pdf;
using OfficeIMO.Pdf;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherCustomPathTests {
    [Theory]
    [InlineData(4)]
    [InlineData(0xFFF0)]
    public void Native_compact_vertices_keep_positive_coordinates_above_the_signed_word_boundary(int elementSize) {
        PublisherDocument publication = PublisherDocument.Load(Input(1,
            new[] { (32767, 32768), (40000, 65535), (65535, 40000) }, elementSize: elementSize,
            extra: new() { [0x140] = 0, [0x141] = 0, [0x142] = 65535, [0x143] = 65535 }));
        OfficeShape shape = Artwork(publication);
        Assert.Equal(shape.Width * 32767 / 65535, shape.PathCommands[0].Point.X, 6);
        Assert.Equal(shape.Height * 32768 / 65535, shape.PathCommands[0].Point.Y, 6);
        Assert.Equal(shape.Width * 40000 / 65535, shape.PathCommands[1].Point.X, 6);
        Assert.Equal(shape.Height, shape.PathCommands[1].Point.Y, 6);
        Assert.Equal(shape.Width, shape.PathCommands[2].Point.X, 6);
    }

    [Theory]
    [InlineData(0U, 8)]
    [InlineData(1U, 8)]
    [InlineData(2U, 8)]
    [InlineData(3U, 8)]
    [InlineData(1U, 4)]
    [InlineData(1U, 0xFFF0)]
    public void Native_vertices_retain_the_declared_canvas_and_implicit_connection_kind(uint pathKind, int elementSize) {
        (int X, int Y)[] points = pathKind < 2
            ? new[] { (35, 5), (85, 5), (35, 55) }
            : new[] { (35, 5), (35, 55), (85, 5), (85, 55) };
        PublisherDocument publication = Load(pathKind, points, elementSize: elementSize);
        OfficeShape shape = Artwork(publication);

        Assert.Equal(OfficeShapeKind.Path, shape.Kind);
        Assert.Equal(new OfficePoint(shape.Width / 4, shape.Height / 4), shape.PathCommands[0].Point);
        Assert.Equal(pathKind < 2 ? OfficePathCommandKind.LineTo : OfficePathCommandKind.CubicBezierTo,
            shape.PathCommands[1].Kind);
        Assert.Equal((pathKind & 1) != 0, shape.PathCommands.Last().Kind == OfficePathCommandKind.Close);
        Assert.DoesNotContain(publication.ReadReport.FidelityDiagnostics,
            item => item.Code == "PUB_SHAPE_GEOMETRY_APPROXIMATED");
    }

    [Fact]
    public void Native_segment_commands_override_the_implicit_shape_path_property() {
        PublisherDocument publication = Load(99, new[] { (10, -20), (110, -20), (10, 80) },
            new ushort[] { 0x4000, 0x0002, 0x6001, 0x8000 });
        OfficeShape shape = Artwork(publication);

        Assert.Equal(OfficeShapeKind.Path, shape.Kind);
        Assert.Equal(new[] { OfficePathCommandKind.MoveTo, OfficePathCommandKind.LineTo,
            OfficePathCommandKind.LineTo, OfficePathCommandKind.Close }, shape.PathCommands.Select(command => command.Kind));
        Assert.Contains("publisher-object-293", PublisherNativeTests.Elements(publication.Pages[0].Drawing)
            .OfType<OfficeDrawingShape>().Single(item => ReferenceEquals(item.Shape, shape)).SourceElementIds!);
    }

    [Fact]
    public void Native_custom_path_paint_flags_suppress_fill_and_stroke_independently() {
        PublisherDocument publication = Load(4, new[] { (10, -20), (110, -20), (10, 80) },
            new ushort[] { 0x4000, 0x0002, 0x6001, 0xAA00, 0x8000 });
        OfficeShape shape = Artwork(publication);

        Assert.Equal(OfficeShapeKind.Path, shape.Kind);
        Assert.Null(shape.FillColor);
        Assert.Null(shape.FillGradient);
        Assert.NotNull(shape.StrokeColor);
    }

    [Fact]
    public void Custom_path_preserves_the_full_gradient_field_and_source_reports_in_exports() {
        byte[] input = Input(1, new[] { (35, 5), (85, 5), (35, 55) }, extra: new() {
            [0x180] = 4, [0x183] = 0x00FF0000
        });
        PublisherDocument publication = PublisherDocument.Load(input);
        OfficeShape shape = Artwork(publication);
        Assert.NotNull(shape.FillGradient);
        Assert.Equal(0.5D, shape.FillGradient!.StartX);
        Assert.Equal(1D, shape.FillGradient.StartY);
        var svg = publication.ToSvgResult();
        Assert.Contains(svg.Report.FidelityDiagnostics, item => item.Code == "PUB_CUSTOM_PATH_RENDERING_UNQUALIFIED"
            && item.Location == "Contents/object/293");
        Assert.Contains("linearGradient", svg.Value);
        var pdf = publication.ToPdfDocumentResult();
        Assert.Contains(publication.ReadReport, pdf.SourceConversionReports);
        Assert.Equal(publication.Pages.Count, PdfReadDocument.Open(pdf.ToBytes()).Pages.Count);
        Assert.Throws<OfficeConversionException>(() => svg.RequireNoLoss());
    }

    [Theory]
    [InlineData(0xA902, "PUB_CUSTOM_PATH_COMMAND_UNSUPPORTED")]
    [InlineData(0xC000, "PUB_CUSTOM_PATH_COMMAND_UNSUPPORTED")]
    [InlineData(0x6000, "PUB_CUSTOM_PATH_INVALID")]
    public void Unsupported_or_malformed_path_commands_keep_explicit_fallback_evidence(ushort word, string code) {
        PublisherDocument publication = Load(4, new[] { (10, -20), (110, -20), (10, 80) },
            new ushort[] { 0x4000, word, 0x8000 });
        Assert.Equal(OfficeShapeKind.Rectangle, Artwork(publication).Kind);
        Assert.Contains(publication.ReadReport.FidelityDiagnostics, item => item.Code == code
            && item.LossKind == OfficeConversionLossKind.Approximation && item.Location == "Contents/object/293");
    }

    [Fact]
    public void Missing_geometry_guides_and_separate_paint_groups_are_not_guessed() {
        PublisherDocument guide = Load(1, new[] { (unchecked((int)0x8000007F), 5), (85, 5), (35, 55) });
        Assert.Contains(guide.ReadReport.FidelityDiagnostics, item => item.Code == "PUB_CUSTOM_PATH_GUIDES_INVALID");
        PublisherDocument groups = Load(4, new[] { (10, -20), (110, -20), (10, 80), (110, 80) },
            new ushort[] { 0x4000, 0x0001, 0x8000, 0x4000, 0x0001, 0x8000 });
        Assert.Contains(groups.ReadReport.FidelityDiagnostics, item => item.Code == "PUB_CUSTOM_PATH_PAINT_GROUPS_UNSUPPORTED");
        Assert.Equal(OfficeShapeKind.Rectangle, Artwork(groups).Kind);
    }

    [Fact]
    public void Invalid_geometry_space_and_truncated_vertices_keep_recoverable_text() {
        byte[] input = Input(1, new[] { (35, 5), (85, 5), (35, 55) }, extra: new() { [0x142] = 10 });
        PublisherDocument publication = PublisherDocument.Load(input);
        Assert.Contains(publication.ReadReport.FidelityDiagnostics, item => item.Code == "PUB_CUSTOM_PATH_INVALID");
        Assert.NotEmpty(publication.TextStories);
        byte[] malformed = Vertices(new[] { (10, -20), (110, 80) });
        malformed[0] = malformed[2] = 3;
        publication = PublisherDocument.Load(PublisherDrawingFixture.Mutate(new(), 0, new() { [0x145] = malformed }));
        Assert.Contains(publication.ReadReport.FidelityDiagnostics, item => item.Code == "PUB_CUSTOM_PATH_INVALID");
        Assert.NotEmpty(publication.TextStories);
    }

    [Fact]
    public void Native_path_expansion_obeys_cumulative_item_work_limits() {
        byte[] input = Input(1, Enumerable.Range(0, 80).Select(index => (index + 10, index - 20)).ToArray());
        InvalidDataException error = Assert.Throws<InvalidDataException>(() => PublisherDocument.Load(input,
            new PublisherReadOptions { Limits = new OfficeLegacyImportLimits { MaxItems = 100 } }));
        Assert.Contains("custom path item work limit", error.Message);
    }

    [Fact]
    public void Native_array_lengths_excluding_headers_preserve_both_commands_and_following_properties() {
        byte[] input = PublisherDrawingFixture.Mutate(new() { [0x144] = 4 }, 0, new() {
            [0x145] = Vertices(new[] { (0, 0), (21600, 0), (0, 21600) }),
            [0x146] = Segments(new ushort[] { 0x4000, 0x0002, 0x6001, 0x8000 }),
            [0x380] = System.Text.Encoding.Unicode.GetBytes("Custom triangle\0")
        }, arrayLengthsExcludeHeader: true);
        PublisherDocument publication = PublisherDocument.Load(input);
        Assert.Equal(OfficeShapeKind.Path, Artwork(publication).Kind);
        Assert.Equal(4, Artwork(publication).PathCommands.Count);
        Assert.DoesNotContain(publication.ReadReport.FidelityDiagnostics, item => item.Code == "PUB_CUSTOM_PATH_INVALID");
    }

    [Fact]
    public void Repeated_path_and_mask_projection_share_a_cumulative_command_budget() {
        var commands = new[] { OfficePathCommand.MoveTo(10, 10) }
            .Concat(Enumerable.Range(0, 10).Select(index => OfficePathCommand.LineTo(index * 5, index * 2)))
            .Append(OfficePathCommand.Close()).ToArray();
        var inner = new OfficeDrawing(100, 100).AddShape(OfficeShape.Path(100, 100, commands), 0, 0);
        var drawing = new OfficeDrawing(100, 100).AddClippedDrawing(inner, 0, 0, OfficeClipPath.Path(100, 100, commands));
        var context = new PublisherParseContext(new PublisherReadOptions {
            Limits = new OfficeLegacyImportLimits { MaxItems = 30 }
        }, default);
        context.AccountProjection(drawing);
        InvalidDataException error = Assert.Throws<InvalidDataException>(() => context.AccountProjection(drawing));
        Assert.Contains("projected path command work limit", error.Message);
    }

    private static PublisherDocument Load(uint pathKind, (int X, int Y)[] points,
        ushort[]? segments = null, int elementSize = 8) => PublisherDocument.Load(Input(pathKind, points, segments, elementSize));

    private static byte[] Input(uint pathKind, (int X, int Y)[] points,
        ushort[]? segments = null, int elementSize = 8, Dictionary<ushort, uint>? extra = null) {
        var complex = new Dictionary<ushort, byte[]> { [0x145] = Vertices(points, elementSize) };
        if (segments != null) complex[0x146] = Segments(segments);
        var values = new Dictionary<ushort, uint> {
            [0x140] = 10, [0x141] = unchecked((uint)-20), [0x142] = 110, [0x143] = 80,
            [0x144] = pathKind, [0x181] = 0x000000FF, [0x1BF] = 0x00100010,
            [0x1C0] = 0x00FF0000, [0x1CB] = 25400, [0x1FF] = 0x00080008
        };
        if (extra != null) foreach (var entry in extra) values[entry.Key] = entry.Value;
        return PublisherDrawingFixture.Mutate(values, 0, complex);
    }

    internal static byte[] Vertices((int X, int Y)[] points, int elementSize = 8) {
        using var stream = new MemoryStream(); using var writer = new BinaryWriter(stream);
        writer.Write(checked((ushort)points.Length)); writer.Write(checked((ushort)points.Length));
        writer.Write(checked((ushort)elementSize));
        foreach ((int x, int y) in points) {
            if (elementSize == 8) { writer.Write(x); writer.Write(y); }
            else { writer.Write(checked((ushort)x)); writer.Write(checked((ushort)y)); }
        }
        return stream.ToArray();
    }

    internal static byte[] Segments(ushort[] segments) {
        using var stream = new MemoryStream(); using var writer = new BinaryWriter(stream);
        writer.Write(checked((ushort)segments.Length)); writer.Write(checked((ushort)segments.Length));
        writer.Write((ushort)2);
        foreach (ushort segment in segments) writer.Write(segment);
        return stream.ToArray();
    }

    private static OfficeShape Artwork(PublisherDocument publication) => PublisherNativeTests.Elements(publication.Pages[0].Drawing)
        .OfType<OfficeDrawingShape>().Single(item => item.SourceElementIds?.Contains("publisher-object-293") == true).Shape;
}
