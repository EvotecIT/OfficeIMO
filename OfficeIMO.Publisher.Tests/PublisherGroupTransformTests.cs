using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Publisher.Internal;
using OfficeIMO.Publisher.Pdf;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherGroupTransformTests {
    [Theory]
    [InlineData(0, true, false)]
    [InlineData(0, false, true)]
    [InlineData(90, false, false)]
    [InlineData(30, true, true)]
    [InlineData(-30, true, false)]
    [InlineData(270, false, true)]
    public void Native_group_transforms_reach_pictures_text_and_public_frame_geometry(double angle, bool horizontal, bool vertical) {
        byte[] input = MutateGroup(angle, horizontal, vertical);
        PublisherDocument publication = PublisherDocument.Load(input);
        PublisherPage page = publication.Pages.Single(p => p.TextFrames.Any(frame => frame.Id == 345));
        PublisherTextFrame frame = page.TextFrames.Single(item => item.Id == 345);
        double centerX = page.Width / 2 - 157823.5 / 12700D, centerY = page.Height / 2 - 1201391 / 12700D;
        foreach (OfficePoint point in new[] { new OfficePoint(0, 0), new OfficePoint(frame.Width, 0), new OfficePoint(frame.Width, frame.Height) }) {
            OfficePoint expected = Transform(new OfficePoint(frame.X + point.X, frame.Y + point.Y), centerX, centerY, angle, horizontal, vertical);
            Equal(expected, frame.PageTransform.TransformPoint(point));
        }
        OfficeDrawingEffectGroup artwork = PublisherNativeTests.Elements(page.Drawing).OfType<OfficeDrawingEffectGroup>()
            .Single(group => PublisherNativeTests.Elements(group.InnerDrawing).OfType<OfficeDrawingImage>()
                .Any(image => image.SourceElementIds?.Contains("publisher-object-346") == true));
        OfficeDrawingImage image = PublisherNativeTests.Elements(artwork.InnerDrawing).OfType<OfficeDrawingImage>()
            .Single(item => item.SourceElementIds?.Contains("publisher-object-346") == true);
        OfficeDrawingImage original = PublisherDocument.Load(PublisherNativeTests.Fixture("SampleNewsletter.pub"))
            .Pages.SelectMany(item => PublisherNativeTests.Elements(item.Drawing)).OfType<OfficeDrawingImage>()
            .Single(item => item.SourceElementIds?.Contains("publisher-object-346") == true);
        foreach (OfficePoint point in new[] { new OfficePoint(0, 0), new OfficePoint(1, 0), new OfficePoint(1, 1) }) {
            OfficePoint unrotated = original.Projection.CreateUnitSquareTransform().TransformPoint(point);
            // The fixture's native group anchor exchanges its width/height at
            // quarter turns; child coordinates map through that restored frame.
            if (angle is 90 or 270) unrotated = new OfficePoint(
                centerX + (unrotated.X - centerX) * 1543214D / 2050741D,
                centerY + (unrotated.Y - centerY) * 2050741D / 1543214D);
            OfficePoint expected = Transform(unrotated, centerX, centerY, angle, horizontal, vertical);
            Equal(expected, artwork.Transform.TransformPoint(image.Projection.CreateUnitSquareTransform().TransformPoint(point)));
        }
        Assert.Equal(publication.Pages.Count, PdfReadDocument.Open(publication.ToPdfBytes()).Pages.Count);
        Assert.Contains("matrix(", publication.ToSvg(publication.Pages.ToList().IndexOf(page)));
    }

    [Fact]
    public void Nested_groups_compose_about_distinct_centres_and_restore_quarter_turn_anchor_dimensions() {
        byte[] leaf = Shape(3, 0, 0, true, new[] { 100, 120, 250, 300 });
        byte[] inner = Group(2, 30, false, true, true, new[] { 200, 100, 800, 600 }, leaf);
        byte[] outer = Group(1, 90, true, false, false, new[] { -127000, -254000, 127000, 254000 }, inner);
        var context = new PublisherParseContext(new PublisherReadOptions(), CancellationToken.None);
        PublisherEscherData result = new PublisherEscherReader(new PublisherBinaryData(outer, "nested group records"), null, context).Read();
        PublisherEscherShape child = result.Shapes.Single(shape => shape.Id == 3);
        Assert.Equal(-91440, child.Bounds!.Value.X1, 6);
        Assert.Equal(-71120, child.Bounds.Value.Y1, 6);
        Assert.Equal(0, child.Bounds.Value.X2, 6);
        Assert.Equal(-25400, child.Bounds.Value.Y2, 6);
        var point = new OfficePoint(32000, -20000);
        OfficePoint expected = Transform(Transform(point, 0, -38100, 30, false, true), 0, 0, 90, true, false);
        Equal(expected, child.GroupTransform.TransformPoint(point));
    }

    [Fact]
    public void Hidden_group_content_is_inert_and_no_longer_excludes_printable_story_text() {
        PublisherDocument baseline = PublisherDocument.Load(PublisherNativeTests.Fixture("SampleNewsletter.pub"));
        PublisherDocument hidden = PublisherDocument.Load(MutateGroup(0, false, false, hidden: true));
        Assert.DoesNotContain(hidden.Pages.SelectMany(page => PublisherNativeTests.Elements(page.Drawing)),
            element => element.SourceElementIds?.Any(id => id is "publisher-object-345" or "publisher-object-346") == true);
        Assert.Contains(hidden.TextStories, story => story.Text == baseline.TextStories.Single(item => item.Id ==
            baseline.Pages.SelectMany(page => page.TextFrames).Single(frame => frame.Id == 345).StoryId).Text);
        Assert.DoesNotContain(hidden.Pages.SelectMany(page => page.TextFrames), frame => frame.Id == 345);
        PublisherTextFrame baselineWrap = baseline.Pages.SelectMany(page => page.TextFrames).Single(frame => frame.Id == 327);
        PublisherTextFrame hiddenWrap = hidden.Pages.SelectMany(page => page.TextFrames).Single(frame => frame.Id == 327);
        Assert.True(hiddenWrap.TextLength > baselineWrap.TextLength);
    }

    [Fact]
    public void Outside_story_wrap_exclusions_follow_the_transformed_group_picture() {
        PublisherDocument control = PublisherDocument.Load(PublisherNativeTests.Fixture("SampleNewsletter.pub"));
        PublisherDocument transformed = PublisherDocument.Load(MutateGroup(30, true, false));
        PublisherPage page = control.Pages.Single(item => item.TextFrames.Any(frame => frame.Id == 345));
        OfficeDrawingImage picture = PublisherNativeTests.Elements(page.Drawing).OfType<OfficeDrawingImage>()
            .Single(item => item.SourceElementIds?.Contains("publisher-object-346") == true);
        // This native picture declares the standard 36,576-EMU wrap distance on every side.
        double distance = 36576 / 12700D;
        double centerX = page.Width / 2 - 157823.5 / 12700D, centerY = page.Height / 2 - 1201391 / 12700D;
        OfficePoint[] corners = new[] {
            new OfficePoint(picture.Projection.X - distance, picture.Projection.Y - distance),
            new OfficePoint(picture.Projection.X + picture.Projection.Width + distance, picture.Projection.Y - distance),
            new OfficePoint(picture.Projection.X - distance, picture.Projection.Y + picture.Projection.Height + distance),
            new OfficePoint(picture.Projection.X + picture.Projection.Width + distance, picture.Projection.Y + picture.Projection.Height + distance)
        }.Select(point => Transform(point, centerX, centerY, 30, true, false)).ToArray();
        OfficeDrawingRichText[] regions = transformed.Pages.SelectMany(item => PublisherNativeTests.Elements(item.Drawing))
            .OfType<OfficeDrawingRichText>().Where(item => item.SourceElementIds?.Contains("publisher-object-327") == true)
            .OrderBy(item => item.Y).ToArray();
        Assert.Equal(2, regions.Length);
        Assert.Equal(corners.Min(point => point.Y), regions[0].Y + regions[0].Height, 6);
        Assert.Equal(corners.Max(point => point.X), regions[1].X, 6);
    }

    [Fact]
    public void Transformed_content_survives_scene_cloning() {
        byte[] input = MutateGroup(30, true, false);
        PublisherDocument document = PublisherDocument.Load(input);
        PublisherPage page = document.Pages.Single(item => item.TextFrames.Any(frame => frame.Id == 345));
        Assert.Equal(OfficeDrawingSvgExporter.ToSvg(page.Drawing), OfficeDrawingSvgExporter.ToSvg(page.Drawing.Clone()));
    }

    [Fact]
    public void Off_page_group_picture_is_retained_when_reflection_moves_it_onto_the_page() {
        PublisherDocument control = PublisherDocument.Load(PublisherNativeTests.Fixture("SampleNewsletter.pub"));
        OfficeDrawingImage original = control.Pages.SelectMany(page => PublisherNativeTests.Elements(page.Drawing))
            .OfType<OfficeDrawingImage>().Single(image => image.SourceElementIds?.Contains("publisher-object-346") == true);
        double offsetY = -original.Projection.Y - 20D;
        PublisherDocument reflected = PublisherDocument.Load(MutateGroup(0, false, true, yOffsetPoints: offsetY));
        PublisherPage page = reflected.Pages.Single(item => item.TextFrames.Any(frame => frame.Id == 345));
        OfficeDrawingEffectGroup artwork = PublisherNativeTests.Elements(page.Drawing).OfType<OfficeDrawingEffectGroup>()
            .Single(group => PublisherNativeTests.Elements(group.InnerDrawing).OfType<OfficeDrawingImage>()
                .Any(image => image.SourceElementIds?.Contains("publisher-object-346") == true));
        OfficeDrawingImage picture = PublisherNativeTests.Elements(artwork.InnerDrawing).OfType<OfficeDrawingImage>().Single();
        var bounds = artwork.Transform.TransformRectangleBounds(picture.Projection.X, picture.Projection.Y,
            picture.Projection.Width, picture.Projection.Height);
        Assert.True(bounds.Top >= 0 && bounds.Bottom < page.Height);
        var isolated = new OfficeDrawing(page.Width, page.Height).AddEffectDrawing(artwork.Drawing, artwork.Transform);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(isolated);
        // This band originated above the page before reflection. It must retain
        // its pixels until the final page clip, even without any backing artwork.
        Assert.Equal(255, raster.GetPixel((int)((bounds.Left + bounds.Right) / 2), (int)(bounds.Bottom - 10D)).A);
        Assert.Contains("data:image/", OfficeDrawingSvgExporter.ToSvg(isolated));
    }

    [Fact]
    public void Inherited_transform_wrappers_cannot_bypass_the_projected_element_budget() {
        var child = new OfficeDrawing(10, 10).AddShape(OfficeShape.Rectangle(2, 2), 1, 1);
        var drawing = new OfficeDrawing(10, 10).AddEffectDrawing(child, OfficeTransform.RotateDegrees(30));
        var context = new PublisherParseContext(new PublisherReadOptions {
            Limits = new OfficeLegacyImportLimits { MaxItems = 1 }
        }, CancellationToken.None);
        Assert.Throws<InvalidDataException>(() => context.AccountProjection(drawing));
    }

    private static byte[] MutateGroup(double angle, bool horizontal, bool vertical, bool hidden = false, double yOffsetPoints = 0D) =>
        PublisherInputContractTests.Mutate("Escher/EscherStm", bytes => {
            bool found = false;
            Visit(bytes, 0, bytes.Length, (children, id) => {
                if (id != 344) return;
                found = true;
                foreach ((int initial, int kind, int offset, int length) in children) {
                    if (kind == 0xF00A) {
                        uint flags = BitConverter.ToUInt32(bytes, offset + 4);
                        PublisherInputContractTests.WriteUInt32(bytes, offset + 4,
                            (flags & ~0xC0U) | (horizontal ? 0x40U : 0) | (vertical ? 0x80U : 0));
                    }
                    if (kind == 0xF00B) {
                        Assert.Equal(1, initial >> 4);
                        ushort property = hidden ? (ushort)0x03BF : (ushort)4;
                        bytes[offset] = (byte)property; bytes[offset + 1] = (byte)(property >> 8);
                        PublisherInputContractTests.WriteUInt32(bytes, offset + 2,
                            hidden ? 0x40004000U : unchecked((uint)(int)(angle * 65536)));
                    }
                    if (kind == 0xF010 && yOffsetPoints != 0D) {
                        for (int item = offset + 4; item < offset + length; item += 6) {
                            ushort key = BitConverter.ToUInt16(bytes, item);
                            if (key is 0x2002 or 0x2004) {
                                int y = BitConverter.ToInt32(bytes, item + 2);
                                PublisherInputContractTests.WriteUInt32(bytes, item + 2,
                                    unchecked((uint)checked(y + (int)Math.Round(yOffsetPoints * 12700D))));
                            }
                        }
                    }
                }
            });
            Assert.True(found);
        }, "SampleNewsletter.pub");

    private static void Visit(byte[] bytes, int start, int end, Action<List<(int Initial, int Kind, int Offset, int Length)>, uint> visitor) {
        for (int offset = start; offset < end;) {
            ushort initial = BitConverter.ToUInt16(bytes, offset), kind = BitConverter.ToUInt16(bytes, offset + 2);
            int content = offset + 8, boundary = content + checked((int)BitConverter.ToUInt32(bytes, offset + 4));
            if (kind == 0xF004) {
                var children = new List<(int, int, int, int)>(); uint id = 0;
                for (int child = content; child < boundary;) {
                    int value = child + 8, length = checked((int)BitConverter.ToUInt32(bytes, child + 4));
                    int type = BitConverter.ToUInt16(bytes, child + 2);
                    children.Add((BitConverter.ToUInt16(bytes, child), type, value, length));
                    if (type == 0xF011 && length == 10) id = BitConverter.ToUInt32(bytes, value + 6);
                    child = value + length;
                }
                visitor(children, id);
            } else if ((initial & 15) == 15) Visit(bytes, content, boundary, visitor);
            offset = boundary + (kind is 0xF000 or 0xF002 && boundary < end ? 4 : 0);
        }
    }

    private static byte[] Group(uint id, double angle, bool horizontal, bool vertical, bool child, int[] anchor, byte[] contents) =>
        Record(0xF003, 15, Shape(id, 1U | (horizontal ? 0x40U : 0) | (vertical ? 0x80U : 0), angle, child, anchor,
            child ? 500 : 1000), contents);

    private static byte[] Shape(uint id, uint flags, double angle, bool child, int[] anchor, int coordinateExtent = 500) {
        byte[] anchorRecord = child ? Record(0xF00F, 0, Integers(anchor)) : Record(0xF010, 0,
            Client(new ushort[] { 0x2001, 0x2002, 0x2003, 0x2004 }, anchor.Select(value => unchecked((uint)value)).ToArray()));
        byte[] coordinates = (flags & 1) == 0 ? Array.Empty<byte>() : Record(0xF009, 1,
            Integers(0, 0, coordinateExtent, coordinateExtent));
        return Record(0xF004, 15, coordinates, Record(0xF00A, 2, Integers(unchecked((int)id), unchecked((int)flags))),
            Record(0xF00B, 19, new byte[] { 4, 0 }.Concat(BitConverter.GetBytes((int)(angle * 65536))).ToArray()), anchorRecord,
            Record(0xF011, 0, Client(new ushort[] { 0x6801 }, new[] { id })));
    }

    private static byte[] Client(ushort[] keys, uint[] values) {
        using var stream = new MemoryStream(); using var writer = new BinaryWriter(stream);
        writer.Write(4 + keys.Length * 6);
        for (int i = 0; i < keys.Length; i++) { writer.Write(keys[i]); writer.Write(values[i]); }
        return stream.ToArray();
    }
    private static byte[] Integers(params int[] values) => values.SelectMany(BitConverter.GetBytes).ToArray();
    private static byte[] Record(ushort kind, ushort initial, params byte[][] contents) {
        byte[] body = contents.SelectMany(value => value).ToArray();
        using var stream = new MemoryStream(); using var writer = new BinaryWriter(stream);
        writer.Write(initial); writer.Write(kind); writer.Write(body.Length); writer.Write(body);
        return stream.ToArray();
    }
    private static OfficePoint Transform(OfficePoint point, double x, double y, double angle, bool horizontal, bool vertical) {
        double dx = (point.X - x) * (horizontal ? -1 : 1), dy = (point.Y - y) * (vertical ? -1 : 1);
        double radians = angle * Math.PI / 180;
        return new OfficePoint(x + dx * Math.Cos(radians) - dy * Math.Sin(radians), y + dx * Math.Sin(radians) + dy * Math.Cos(radians));
    }
    private static void Equal(OfficePoint expected, OfficePoint actual) {
        Assert.Equal(expected.X, actual.X, 6); Assert.Equal(expected.Y, actual.Y, 6);
    }
}
