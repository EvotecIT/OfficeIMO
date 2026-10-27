using OfficeIMO.Drawing;
using OfficeIMO.Drawing.Binary;
using OfficeIMO.Pdf;
using OfficeIMO.Publisher.Internal;
using OfficeIMO.Publisher.Pdf;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherGeometryGuideTests {
    [Theory]
    [InlineData(0x2000, 0x04F8)]
    [InlineData(0x0001, 10)]
    public void Native_paths_ignore_unreferenced_guides_without_a_false_fallback(ushort operation, ushort operand) {
        var publication = PublisherDocument.Load(Input(new[] {
            ((ushort)0, (ushort)25, (ushort)0, (ushort)0),
            ((ushort)0x2000, (ushort)0x0400, (ushort)50, (ushort)0),
            (operation, operand, (ushort)0, (ushort)0) }));
        OfficeShape shape = Artwork(publication);
        Assert.Equal(OfficeShapeKind.Path, shape.Kind);
        Assert.Equal(new OfficePoint(shape.Width / 4, shape.Height / 4), shape.PathCommands[0].Point);
        Assert.Equal(shape.Width * 3 / 4, shape.PathCommands[1].Point.X, 10);
        Assert.DoesNotContain(publication.ReadReport.FidelityDiagnostics, item => item.Code.StartsWith("PUB_CUSTOM_PATH_GUIDE"));
    }

    [Theory]
    [InlineData(0x0145, false)]
    [InlineData(0x0146, true)]
    public void Native_cleared_default_paths_keep_preset_geometry_without_false_path_loss(int propertyId, bool tertiary) {
        byte[] input = PublisherDrawingFixture.Mutate(new() { [0x145] = 0, [0x146] = 0,
            [0x181] = 0x000000FF, [0x1BF] = 0x00100010 }, 1);
        Assert.True(OfficeCompoundFileReader.TryRead(input, out OfficeCompoundFile? source, out string? error), error);
        byte[] escher = source!.Streams["Escher/EscherStm"];
        int length = checked((int)BitConverter.ToUInt32(escher, 4));
        byte[] payload = propertyId == 0x145
            ? PublisherCustomPathTests.Vertices(new[] { (0, 0), (21600, 21600) })
            : PublisherCustomPathTests.Segments(new ushort[] { 0x4000, 0x0001, 0x8000 });
        using var data = new MemoryStream(); using var writer = new BinaryWriter(data);
        writer.Write((ushort)0x13); writer.Write((ushort)(tertiary ? 0xF122 : 0xF00B)); writer.Write(6 + payload.Length);
        writer.Write((ushort)(propertyId | 0x8000)); writer.Write((uint)payload.Length); writer.Write(payload);
        byte[] defaults = data.ToArray();
        byte[] replacement = escher.Take(8 + length).Concat(defaults).Concat(escher.Skip(8 + length)).ToArray();
        Array.Copy(BitConverter.GetBytes(length + defaults.Length), 0, replacement, 4, 4);
        input = OfficeCompoundFileWriter.Rewrite(source, new Dictionary<string, byte[]> { ["Escher/EscherStm"] = replacement });
        PublisherDocument publication = PublisherDocument.Load(input);
        Assert.Equal(OfficeShapeKind.Rectangle, Artwork(publication).Kind);
        Assert.DoesNotContain(publication.ReadReport.FidelityDiagnostics, item => item.Location == "Contents/object/293"
            && item.Code.StartsWith("PUB_CUSTOM_PATH_"));
    }

    [Fact]
    public void Drawing_default_application_is_bounded_without_recounting_container_headers() {
        byte[] Record(ushort kind, ushort initial, byte[] body) => BitConverter.GetBytes(initial)
            .Concat(BitConverter.GetBytes(kind)).Concat(BitConverter.GetBytes(body.Length)).Concat(body).ToArray();
        byte[] properties = BitConverter.GetBytes((ushort)0x0148).Concat(BitConverter.GetBytes(40U))
            .Concat(BitConverter.GetBytes((ushort)0x01FF)).Concat(BitConverter.GetBytes(0x00080000U)).ToArray();
        byte[] group = Record(0xF000, 15, Record(0xF00B, 0x23, properties));
        byte[] shape = Record(0xF004, 15, Record(0xF00A, 2, BitConverter.GetBytes(1U).Concat(BitConverter.GetBytes(1U)).ToArray()));
        byte[] bytes = group.Concat(new byte[4]).Concat(shape).Concat(shape).ToArray();
        var context = new PublisherParseContext(new PublisherReadOptions {
            Limits = new OfficeLegacyImportLimits { MaxRecords = 12 }
        }, default);
        new PublisherEscherReader(new PublisherBinaryData(bytes, "controlled defaults"), null, context).Read();
        Assert.Equal(12, context.Records);
        var limited = new PublisherParseContext(new PublisherReadOptions {
            Limits = new OfficeLegacyImportLimits { MaxRecords = 11 }
        }, default);
        InvalidDataException error = Assert.Throws<InvalidDataException>(() =>
            new PublisherEscherReader(new PublisherBinaryData(bytes, "controlled defaults"), null, limited).Read());
        Assert.Contains("native record limit", error.Message);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void Native_guides_inherit_drawing_group_defaults_and_keep_local_boolean_overrides(bool tertiary, bool localStroke) {
        byte[] input = Input(new[] { ((ushort)0x2000, (ushort)0x01FC, (ushort)0, (ushort)0),
            ((ushort)0x2000, (ushort)0x0148, (ushort)0, (ushort)0) });
        var local = new Dictionary<ushort, uint> { [0x01FF] = localStroke ? 0x00080008U : 0x00400000U };
        if (localStroke) local[0x0148] = 45;
        input = PublisherDrawingFixture.Mutate(input, local, 0);
        Assert.True(OfficeCompoundFileReader.TryRead(input, out OfficeCompoundFile? source, out string? error), error);
        byte[] escher = source!.Streams["Escher/EscherStm"];
        int length = checked((int)BitConverter.ToUInt32(escher, 4));
        Assert.Equal(0xF000, BitConverter.ToUInt16(escher, 2));
        using var data = new MemoryStream(); using var writer = new BinaryWriter(data);
        writer.Write((ushort)0x23); writer.Write((ushort)(tertiary ? 0xF122 : 0xF00B)); writer.Write(12);
        writer.Write((ushort)0x0148); writer.Write((uint)40);
        writer.Write((ushort)0x01FF); writer.Write((uint)0x00080000);
        byte[] defaults = data.ToArray();
        byte[] replacement = escher.Take(8 + length).Concat(defaults).Concat(escher.Skip(8 + length)).ToArray();
        Array.Copy(BitConverter.GetBytes(length + defaults.Length), 0, replacement, 4, 4);
        input = OfficeCompoundFileWriter.Rewrite(source, new Dictionary<string, byte[]> { ["Escher/EscherStm"] = replacement });
        PublisherDocument publication = PublisherDocument.Load(input);
        OfficeShape shape = Artwork(publication);
        Assert.Equal(OfficeShapeKind.Path, shape.Kind);
        Assert.Equal(shape.Width / 100 * (localStroke ? 1 : 0), shape.PathCommands[0].Point.X, 10);
        Assert.Equal(shape.Width * (localStroke ? 0.45 : 0.4), shape.PathCommands[1].Point.X, 10);
        Assert.DoesNotContain(publication.ReadReport.FidelityDiagnostics, item => item.Code == "PUB_CUSTOM_PATH_GUIDES_INVALID");
    }

    [Fact]
    public void Native_frame_guides_preserve_subpoint_extents_and_odd_emu_halves() {
        const int nativeWidth = 19051, nativeHeight = 31751;
        byte[] guides = GuideArray(new[] { ((ushort)0x2000, (ushort)0x04FC, (ushort)0, (ushort)0),
            ((ushort)0x2000, (ushort)0x04FD, (ushort)0, (ushort)0),
            ((ushort)0x2000, (ushort)0x04FE, (ushort)0, (ushort)0),
            ((ushort)0x2000, (ushort)0x04FF, (ushort)0, (ushort)0) });
        byte[] vertices = PublisherCustomPathTests.Vertices(new[] {
            (Guide(0), Guide(1)), (Guide(2), Guide(3)), (0, 0) });
        var source = new PublisherContents(612, 792, new Dictionary<uint, PublisherContentChunk>());
        source.Pages.Add(new PublisherSourcePage(1, "", null, null));
        source.Shapes.Add(2, new PublisherSourceShape(new PublisherContentChunk(2, 1, 0, 1), Array.Empty<PublisherBlock>()));
        source.ObjectPages.Add(2, 1);
        var escher = new PublisherEscherData();
        escher.Shapes.Add(new PublisherEscherShape(2, 0, 0, new[] {
            new OfficeArtProperty(0, 0x140, 0), new OfficeArtProperty(1, 0x141, 0),
            new OfficeArtProperty(2, 0x142, nativeWidth), new OfficeArtProperty(3, 0x143, nativeHeight),
            new OfficeArtProperty(4, 0x144, 1),
            new OfficeArtProperty(5, 0x8145, (uint)vertices.Length, vertices.Length, complexData: vertices),
            new OfficeArtProperty(6, 0x8156, (uint)guides.Length, guides.Length, complexData: guides),
            new OfficeArtProperty(7, 0x181, 0x000000FF), new OfficeArtProperty(8, 0x1BF, 0x00100010)
        }) { Bounds = new PublisherNativeRectangle(0, 0, nativeWidth, nativeHeight) });
        PublisherDocument publication = new PublisherPageProjector(source, new PublisherQuillText(), escher,
            new PublisherParseContext(new PublisherReadOptions(), CancellationToken.None)).Project();
        OfficeShape shape = PublisherNativeTests.Elements(publication.Pages[0].Drawing).OfType<OfficeDrawingShape>()
            .Single(item => item.SourceElementIds?.Contains("publisher-object-2") == true).Shape;
        Assert.Equal(OfficeShapeKind.Path, shape.Kind);
        Assert.Equal(nativeWidth / 12700D, shape.Width, 10);
        Assert.Equal(nativeHeight / 12700D, shape.Height, 10);
        Assert.Equal(shape.Width, shape.PathCommands[0].Point.X, 10);
        Assert.Equal(shape.Height, shape.PathCommands[0].Point.Y, 10);
        Assert.Equal((nativeWidth / 2) / 12700D, shape.PathCommands[1].Point.X, 10);
        Assert.Equal((nativeHeight / 2) / 12700D, shape.PathCommands[1].Point.Y, 10);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Native_guides_resolve_adjustments_and_previous_results_without_losing_following_complex_properties(bool excludeHeaders) {
        byte[] input = Input(new[] { ((ushort)0x2000, (ushort)0x0147, (ushort)0, (ushort)0),
            ((ushort)0x2000, (ushort)0x0400, (ushort)50, (ushort)0) }, excludeHeaders: excludeHeaders);
        PublisherDocument publication = PublisherDocument.Load(input);
        OfficeShape shape = Artwork(publication);
        Assert.Equal(OfficeShapeKind.Path, shape.Kind);
        Assert.Equal(new OfficePoint(shape.Width / 4, shape.Height / 4), shape.PathCommands[0].Point);
        Assert.Equal(new OfficePoint(shape.Width * 3 / 4, shape.Height / 4), shape.PathCommands[1].Point);
        Assert.DoesNotContain(publication.ReadReport.FidelityDiagnostics, item => item.Code.StartsWith("PUB_CUSTOM_PATH_GUIDE"));
        var svg = publication.ToSvgResult();
        Assert.Contains(svg.Report.FidelityDiagnostics, item => item.Code == "PUB_CUSTOM_PATH_RENDERING_UNQUALIFIED"
            && item.Message.Contains("producer rounding") && item.Location == "Contents/object/293");
        Assert.Throws<OfficeConversionException>(() => svg.RequireNoLoss());
        var pdf = publication.ToPdfDocumentResult();
        Assert.Contains(publication.ReadReport, pdf.SourceConversionReports);
        Assert.Equal(publication.Pages.Count, PdfDocument.Load(new System.IO.MemoryStream(pdf.ToBytes())).Pages.Count);
        Assert.NotEmpty(publication.TextStories);
    }

    [Theory]
    [InlineData(0x2000, 0x0400, "PUB_CUSTOM_PATH_GUIDES_INVALID")]
    [InlineData(0x0011, 10, "PUB_CUSTOM_PATH_GUIDE_FORMULA_UNSUPPORTED")]
    [InlineData(0x2000, 0x04F8, "PUB_CUSTOM_PATH_GUIDE_PARAMETER_UNSUPPORTED")]
    [InlineData(0x0001, 10, "PUB_CUSTOM_PATH_GUIDES_INVALID")]
    public void Native_invalid_or_unsupported_guides_preserve_text_and_report_the_geometry_fallback(
        ushort operation, ushort first, string code) {
        var publication = PublisherDocument.Load(Input(new[] { (operation, first, (ushort)0, (ushort)0) }));
        Assert.Equal(OfficeShapeKind.Rectangle, Artwork(publication).Kind);
        Assert.NotEmpty(publication.TextStories);
        Assert.Contains(publication.ReadReport.FidelityDiagnostics, item => item.Code == code
            && item.Location == "Contents/object/293" && item.LossKind == OfficeConversionLossKind.Approximation);
    }

    [Fact]
    public void Native_guides_retain_recovery_limits_in_the_complete_reader() {
        var records = Enumerable.Range(0, 128).Select(index => ((ushort)0, (ushort)index, (ushort)0, (ushort)0)).ToArray();
        var error = Assert.Throws<InvalidDataException>(() => PublisherDocument.Load(Input(records),
            new PublisherReadOptions { Limits = new OfficeLegacyImportLimits { MaxItems = 100 } }));
        Assert.Contains("custom path item work limit", error.Message);
    }

    [Fact]
    public void Limousine_scaling_keeps_an_explicit_fallback_even_for_literal_paths() {
        byte[] input = PublisherDrawingFixture.Mutate(new() {
            [0x0153] = 100, [0x0144] = 1, [0x181] = 0x000000FF, [0x1BF] = 0x00100010
        }, 0,
            new() { [0x0145] = PublisherCustomPathTests.Vertices(new[] { (0, 0), (21600, 0), (0, 21600) }) });
        var publication = PublisherDocument.Load(input);
        Assert.Equal(OfficeShapeKind.Rectangle, Artwork(publication).Kind);
        Assert.Contains(publication.ReadReport.FidelityDiagnostics, item => item.Code == "PUB_CUSTOM_PATH_SCALING_UNSUPPORTED");
    }

    [Fact]
    public void Guided_picture_masks_use_the_same_geometry_and_keep_original_assets() {
        byte[] original = PublisherPictureEffectTests.PictureInput(new());
        var values = new Dictionary<ushort, uint> { [0x140] = 0, [0x141] = 0, [0x142] = 100, [0x143] = 100,
            [0x144] = 1, [0x147] = 25 };
        byte[] input = PublisherDrawingFixture.Mutate(original, values, 0, Complex(new[] {
            ((ushort)0x2000, (ushort)0x0147, (ushort)0, (ushort)0),
            ((ushort)0x2000, (ushort)0x0400, (ushort)50, (ushort)0),
            ((ushort)0x2000, (ushort)0x04F8, (ushort)0, (ushort)0) }), 362);
        var options = new PublisherReadOptions { ImageCodec = new SolidCodec() };
        PublisherDocument publication = PublisherDocument.Load(input, options);
        PublisherDocument control = PublisherDocument.Load(original, options);
        Assert.Equal(control.Images.SelectMany(image => image.GetBytes()), publication.Images.SelectMany(image => image.GetBytes()));
        var group = publication.Pages.SelectMany(page => PublisherNativeTests.Elements(page.Drawing))
            .OfType<OfficeDrawingGroup>().Single(item => item.SourceElementIds?.Contains("publisher-object-362") == true);
        Assert.Equal(OfficeClipPathKind.Path, group.ClipPath.Kind);
        Assert.Equal(new OfficePoint(group.ClipPath.Width / 4, group.ClipPath.Height / 4), group.ClipPath.Commands[0].Point);
        var drawing = new OfficeDrawing(group.ClipPath.Width, group.ClipPath.Height)
            .AddClippedDrawing(group.InnerDrawing, 0, 0, group.ClipPath);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing);
        Assert.Equal(OfficeColor.Red, raster.GetPixel((int)(raster.Width * 0.3), (int)(raster.Height * 0.3)));
        Assert.Equal(0, raster.GetPixel((int)(raster.Width * 0.8), (int)(raster.Height * 0.8)).A);
        Assert.Equal(0, raster.GetPixel((int)(raster.Width * 0.1), (int)(raster.Height * 0.1)).A);
    }

    internal static byte[] Input((ushort Operation, ushort First, ushort Second, ushort Third)[] records,
        bool excludeHeaders = false) => PublisherDrawingFixture.Mutate(new() {
            [0x140] = 0, [0x141] = 0, [0x142] = 100, [0x143] = 100, [0x144] = 1,
            [0x147] = 25, [0x181] = 0x000000FF, [0x1BF] = 0x00100010
        }, 0, Complex(records), arrayLengthsExcludeHeader: excludeHeaders);

    private static Dictionary<ushort, byte[]> Complex((ushort Operation, ushort First, ushort Second, ushort Third)[] records) => new() {
        [0x0145] = PublisherCustomPathTests.Vertices(new[] { (Guide(0), Guide(0)), (Guide(1), Guide(0)), (Guide(0), Guide(1)) }),
        [0x0156] = GuideArray(records),
        [0x0380] = System.Text.Encoding.Unicode.GetBytes("Guided artwork\0")
    };
    internal static byte[] GuideArray((ushort Operation, ushort First, ushort Second, ushort Third)[] records) {
        using var stream = new MemoryStream(); using var writer = new BinaryWriter(stream);
        writer.Write(checked((ushort)records.Length)); writer.Write(checked((ushort)records.Length)); writer.Write((ushort)8);
        foreach (var record in records) {
            writer.Write(record.Operation); writer.Write(record.First); writer.Write(record.Second); writer.Write(record.Third);
        }
        return stream.ToArray();
    }
    private static int Guide(int index) => unchecked((int)0x80000000) | index;
    private static OfficeShape Artwork(PublisherDocument publication) => PublisherNativeTests.Elements(publication.Pages[0].Drawing)
        .OfType<OfficeDrawingShape>().Single(item => item.SourceElementIds?.Contains("publisher-object-293") == true).Shape;
    private sealed class SolidCodec : IOfficeRasterImageCodec {
        public bool TryDecode(byte[] encodedBytes, string? contentType, out OfficeRasterImage? image) {
            image = new OfficeRasterImage(20, 20, OfficeColor.Red); return true;
        }
    }
}
