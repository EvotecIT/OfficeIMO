using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Publisher.Pdf;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherLineStyleTests {
    private static readonly OfficeColor Ink = OfficeColor.FromRgb(204, 32, 64);

    [Theory]
    [InlineData(0U, 0U, OfficeStrokeLineCap.Round, OfficeStrokeLineJoin.Bevel)]
    [InlineData(1U, 1U, OfficeStrokeLineCap.Square, OfficeStrokeLineJoin.Miter)]
    [InlineData(2U, 2U, OfficeStrokeLineCap.Butt, OfficeStrokeLineJoin.Round)]
    public void Native_caps_joins_and_miter_limits_survive_scene_and_svg_projection(uint cap, uint join,
        OfficeStrokeLineCap expectedCap, OfficeStrokeLineJoin expectedJoin) {
        PublisherDocument publication = Load(new Dictionary<ushort, uint> {
            [0x01D7] = cap, [0x01D6] = join, [0x01CC] = 0x00044000
        }, shapeType: 5);
        OfficeShape shape = Artwork(publication);
        Assert.Equal(expectedCap, shape.StrokeLineCap);
        Assert.Equal(expectedJoin, shape.StrokeLineJoin);
        Assert.Equal(4.25, shape.StrokeMiterLimit);
        Assert.Contains("stroke-miterlimit=\"4.25\"", publication.ToSvg(0));
        Assert.Equal(publication.ToSvg(0), OfficeDrawingSvgExporter.ToSvg(publication.Pages[0].Drawing.Clone()));
    }

    [Fact]
    public void Missing_native_line_details_use_flat_caps_round_joins_and_no_markers() {
        OfficeShape shape = Artwork(Load(new Dictionary<ushort, uint>()));
        Assert.Equal(OfficeStrokeLineCap.Butt, shape.StrokeLineCap);
        Assert.Equal(OfficeStrokeLineJoin.Round, shape.StrokeLineJoin);
        Assert.Equal(8, shape.StrokeMiterLimit);
        Assert.Null(shape.StrokeStartMarker); Assert.Null(shape.StrokeEndMarker);
    }

    [Theory]
    [InlineData(5U, OfficeStrokeDashStyle.Dot, new double[] { 10, 30 })]
    [InlineData(6U, OfficeStrokeDashStyle.Dash, new double[] { 40, 30 })]
    [InlineData(8U, OfficeStrokeDashStyle.DashDot, new double[] { 40, 30, 10, 30 })]
    [InlineData(10U, OfficeStrokeDashStyle.DashDotDot, new double[] { 80, 30, 10, 30, 10, 30 })]
    public void Native_dash_vocabulary_preserves_dot_dash_order_and_relative_spacing(uint value, OfficeStrokeDashStyle expected, double[] pattern) {
        PublisherDocument publication = Load(new Dictionary<ushort, uint> { [0x01CE] = value });
        Assert.Equal(expected, Artwork(publication).StrokeDashStyle);
        Assert.Equal(pattern, Artwork(publication).StrokeDashArray);
    }

    [Theory]
    [InlineData(1U, OfficeLineMarkerKind.Triangle)]
    [InlineData(2U, OfficeLineMarkerKind.Stealth)]
    [InlineData(3U, OfficeLineMarkerKind.Diamond)]
    [InlineData(4U, OfficeLineMarkerKind.Oval)]
    [InlineData(5U, OfficeLineMarkerKind.Arrow)]
    public void Native_marker_kinds_and_relative_sizes_reach_svg_and_pdf_with_fidelity_evidence(uint nativeKind, OfficeLineMarkerKind kind) {
        PublisherDocument publication = Load(new Dictionary<ushort, uint> {
            [0x01D0] = nativeKind, [0x01D1] = nativeKind,
            [0x01D2] = 0, [0x01D3] = 0, [0x01D4] = 2, [0x01D5] = 2
        });
        OfficeShape shape = Artwork(publication);
        Assert.Equal(kind, shape.StrokeStartMarker!.Kind); Assert.Equal(kind, shape.StrokeEndMarker!.Kind);
        Assert.Equal(30, shape.StrokeStartMarker.Width); Assert.Equal(40, shape.StrokeStartMarker.Length);
        Assert.Equal(60, shape.StrokeEndMarker.Width); Assert.Equal(80, shape.StrokeEndMarker.Length);
        string svg = publication.ToSvg(0);
        Assert.Equal(svg, OfficeDrawingSvgExporter.ToSvg(publication.Pages[0].Drawing.Clone()));
        Assert.Contains(publication.ReadReport.FidelityDiagnostics, item => item.Code == "PUB_LINE_MARKER_APPROXIMATED"
            && item.LossKind == OfficeConversionLossKind.Approximation && item.Location == "Contents/object/293");
        var result = publication.ToPdfDocumentResult();
        OfficeShape[] painted = PublisherNativeTests.Elements(PdfReadDocument.Open(result.ToBytes()).Pages[0].ToDrawing())
            .OfType<OfficeDrawingShape>().Select(item => item.Shape).ToArray();
        if (kind == OfficeLineMarkerKind.Arrow) {
            Assert.Equal(3, painted.Count(item => item.StrokeColor == Ink));
            Assert.DoesNotContain(painted, item => item.FillColor == Ink);
            Assert.Contains("<polyline", svg);
        } else {
            Assert.Equal(2, painted.Count(item => item.FillColor == Ink));
        }
        Assert.Contains(publication.ReadReport, result.SourceConversionReports);
    }

    [Fact]
    public void Unsupported_line_details_are_reported_and_invisible_strokes_do_not_create_marker_losses() {
        var values = new Dictionary<ushort, uint> {
            [0x01D0] = 99, [0x01D1] = 1, [0x01D4] = 99,
            [0x01D6] = 99, [0x01D7] = 99, [0x01CC] = 0,
            [0x01CD] = 2
        };
        PublisherDocument publication = Load(values);
        Assert.Null(Artwork(publication).StrokeStartMarker);
        Assert.Equal(45, Artwork(publication).StrokeEndMarker!.Width);
        Assert.Contains(publication.ReadReport.FidelityDiagnostics, item => item.Code == "PUB_LINE_MARKER_UNSUPPORTED");
        Assert.Contains(publication.ReadReport.FidelityDiagnostics, item => item.Code == "PUB_LINE_STYLE_APPROXIMATED");
        Assert.Contains(publication.ReadReport.FidelityDiagnostics, item => item.Code == "PUB_LINE_DETAIL_UNASSESSED");
        values[0x01FF] = 0x00080000;
        PublisherDocument hiddenStroke = Load(values);
        Assert.DoesNotContain(hiddenStroke.ReadReport.FidelityDiagnostics, item => item.Code.StartsWith("PUB_LINE_"));
    }

    [Fact]
    public void Reserved_markers_are_ignored_and_closed_geometries_report_unsupported_decorations() {
        PublisherDocument reserved = Load(new Dictionary<ushort, uint> { [0x01D0] = 6, [0x01D1] = 7 });
        Assert.Null(Artwork(reserved).StrokeStartMarker); Assert.Null(Artwork(reserved).StrokeEndMarker);
        Assert.DoesNotContain(reserved.ReadReport.FidelityDiagnostics, item => item.Code == "PUB_LINE_MARKER_APPROXIMATED");
        PublisherDocument rectangle = Load(new Dictionary<ushort, uint> { [0x01D0] = 1 }, shapeType: 1);
        Assert.Null(Artwork(rectangle).StrokeStartMarker);
        Assert.Contains(rectangle.ReadReport.FidelityDiagnostics, item => item.Code == "PUB_LINE_MARKER_GEOMETRY_UNSUPPORTED"
            && item.LossKind == OfficeConversionLossKind.Omission);
    }

    private static PublisherDocument Load(Dictionary<ushort, uint> values, int shapeType = 20) => PublisherDocument.Load(Mutate(values, shapeType));
    private static OfficeShape Artwork(PublisherDocument publication) => PublisherNativeTests.Elements(publication.Pages[0].Drawing)
        .OfType<OfficeDrawingShape>().Single(item => item.SourceElementIds?.Contains("publisher-object-293") == true).Shape;

    internal static byte[] Mutate(Dictionary<ushort, uint> values, int shapeType = 20) {
        Assert.True(OfficeCompoundFileReader.TryRead(File.ReadAllBytes(PublisherNativeTests.Fixture("Simple.pub")), out OfficeCompoundFile? source, out string? error), error);
        var properties = new Dictionary<ushort, uint>(values) {
            [0x0181] = 0, [0x01BF] = 0x00100000, [0x01C0] = 0x004020CC,
            [0x01CB] = 127000
        };
        if (!properties.ContainsKey(0x01FF)) properties[0x01FF] = 0x00080008;
        bool found = false;
        byte[] Rewrite(byte[] bytes, int start, int end) {
            using var output = new MemoryStream(); using var writer = new BinaryWriter(output);
            for (int offset = start; offset < end;) {
                ushort initial = BitConverter.ToUInt16(bytes, offset), kind = BitConverter.ToUInt16(bytes, offset + 2);
                int content = offset + 8, boundary = content + checked((int)BitConverter.ToUInt32(bytes, offset + 4));
                byte[] body = bytes.Skip(content).Take(boundary - content).ToArray();
                if (kind == 0xF004) {
                    int client = content;
                    while (client < boundary) {
                        int length = checked((int)BitConverter.ToUInt32(bytes, client + 4));
                        if (BitConverter.ToUInt16(bytes, client + 2) == 0xF011 && length == 10 && BitConverter.ToUInt32(bytes, client + 14) == 293) {
                            found = true; body = RewriteShape(bytes, content, boundary, properties, shapeType); break;
                        }
                        client += 8 + length;
                    }
                } else if ((initial & 15) == 15) body = Rewrite(bytes, content, boundary);
                writer.Write(initial); writer.Write(kind); writer.Write(body.Length); writer.Write(body);
                offset = boundary;
                if (kind is 0xF000 or 0xF002 && boundary < end) { writer.Write(bytes, boundary, 4); offset += 4; }
            }
            return output.ToArray();
        }
        byte[] escher = source!.Streams["Escher/EscherStm"];
        byte[] replacement = Rewrite(escher, 0, escher.Length);
        Assert.True(found);
        return OfficeCompoundFileWriter.Rewrite(source, new Dictionary<string, byte[]> { ["Escher/EscherStm"] = replacement });
    }

    private static byte[] RewriteShape(byte[] bytes, int start, int end, Dictionary<ushort, uint> values, int shapeType) {
        using var output = new MemoryStream(); using var writer = new BinaryWriter(output);
        for (int offset = start; offset < end;) {
            ushort initial = BitConverter.ToUInt16(bytes, offset), kind = BitConverter.ToUInt16(bytes, offset + 2);
            int content = offset + 8, length = checked((int)BitConverter.ToUInt32(bytes, offset + 4));
            byte[] body = bytes.Skip(content).Take(length).ToArray();
            if (kind == 0xF00A) initial = (ushort)((shapeType << 4) | (initial & 15));
            if (kind is 0xF00B or 0xF122) {
                int count = initial >> 4;
                var entries = new List<(ushort Op, uint Value)>();
                for (int i = 0; i < count; i++) {
                    ushort op = BitConverter.ToUInt16(body, i * 6);
                    if (!values.ContainsKey((ushort)(op & 0x3FFF))) entries.Add((op, BitConverter.ToUInt32(body, i * 6 + 2)));
                }
                entries.AddRange(values.Select(item => (item.Key, item.Value)));
                using var data = new MemoryStream(); using var properties = new BinaryWriter(data);
                foreach (var entry in entries.OrderBy(item => item.Op & 0x3FFF)) { properties.Write(entry.Op); properties.Write(entry.Value); }
                properties.Write(body, count * 6, body.Length - count * 6);
                body = data.ToArray(); initial = (ushort)((entries.Count << 4) | (initial & 15));
            }
            writer.Write(initial); writer.Write(kind); writer.Write(body.Length); writer.Write(body);
            offset = content + length;
        }
        return output.ToArray();
    }
}
