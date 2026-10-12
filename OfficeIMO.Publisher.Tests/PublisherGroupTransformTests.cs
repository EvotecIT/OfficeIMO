using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Publisher.Internal;
using OfficeIMO.Publisher.Pdf;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherGroupTransformTests {
    [Fact]
    public void Client_anchors_keep_page_units_inside_a_scaled_group() {
        byte[] leaf = Shape(3, 0, 0, false, new[] { -127000, -254000, 127000, 254000 }, filled: true);
        PublisherDocument document = ProjectRecords(Record(0xF003, 15,
            Shape(1, 1, 30, false, new[] { -1270000, -635000, 1270000, 635000 }, 1270000), leaf));
        PublisherTextFrame frame = Assert.Single(document.Pages[0].TextFrames);
        Assert.Equal(20, frame.Width, 6); Assert.Equal(40, frame.Height, 6);
        Equal(Transform(new OfficePoint(296, 376), 306, 396, 30, false, false),
            frame.PageTransform.TransformPoint(new OfficePoint(0, 0)));
    }

    [Theory]
    [InlineData(30, 0x80008000U, true, false)]
    [InlineData(0, 0x40004000U, false, true)]
    [InlineData(30, 0xC000C000U, true, true)]
    [InlineData(0, 0x0000C000U, false, false)]
    public void Grouped_client_anchors_apply_relative_rotation_and_masked_flips(double angle, uint flips, bool horizontal, bool vertical) {
        byte[] leaf = Shape(3, flips == 0x0000C000U ? 0xC0U : 0, 20, false, new[] { -127000, -254000, 127000, 254000 },
            relativeAngle: angle, filled: true, relativeFlips: flips);
        PublisherDocument document = ProjectRecords(Record(0xF003, 15,
            Shape(1, 1, 30, false, new[] { -1270000, -635000, 1270000, 635000 }, 1270000), leaf));
        PublisherTextFrame frame = Assert.Single(document.Pages[0].TextFrames);
        foreach (OfficePoint point in new[] { new OfficePoint(0, 0), new OfficePoint(20, 0), new OfficePoint(20, 40) }) {
            OfficePoint local = Transform(new OfficePoint(296 + point.X, 376 + point.Y), 306, 396, angle, horizontal, vertical);
            Equal(Transform(local, 306, 396, 30, false, false), frame.PageTransform.TransformPoint(point));
        }
    }

    private static PublisherDocument CreateNonuniformGroup(bool relative, double angle) {
        byte[] leaf = Shape(3, 0, relative ? 20 : angle, true, new[] { 317500, 317500, 952500, 952500 },
            relativeAngle: relative ? angle : null, filled: true);
        byte[] records = Record(0xF003, 15,
            Shape(1, 1, 30, false, new[] { -1270000, -635000, 1270000, 635000 }, 1270000), leaf);
        return ProjectRecords(records);
    }

    private static PublisherDocument ProjectRecords(byte[] records) {
        var escher = new PublisherEscherReader(new PublisherBinaryData(records, "nonuniform group"), null,
            new PublisherParseContext(new PublisherReadOptions(), default)).Read();
        var source = new PublisherContents(612, 792, new Dictionary<uint, PublisherContentChunk>());
        source.Pages.Add(new PublisherSourcePage(10, "", null, null));
        source.Shapes.Add(1, new PublisherSourceShape(new PublisherContentChunk(1, 1, 0, 10), Array.Empty<PublisherBlock>()));
        source.Shapes.Add(3, new PublisherSourceShape(new PublisherContentChunk(3, 1, 0, 10),
            new[] { new PublisherBlock(0x27, 0, 0, 0, 4) }));
        source.ObjectPages.Add(1, 10); source.ObjectPages.Add(3, 10);
        var text = new PublisherQuillText();
        text.Stories.Add(4, new PublisherTextStory(4, "", Array.Empty<OfficeRichTextParagraph>()));
        return new PublisherPageProjector(source, text, escher,
            new PublisherParseContext(new PublisherReadOptions(), default)).Project();
    }

    [Theory]
    [InlineData(false, 45)]
    [InlineData(false, 90)]
    [InlineData(false, -30)]
    [InlineData(true, 45)]
    [InlineData(true, 0)]
    public void Nonuniform_group_scaling_follows_child_rotation_in_artwork_and_frame_corners(bool relative, double angle) {
        PublisherDocument document = CreateNonuniformGroup(relative, angle);
        PublisherTextFrame frame = Assert.Single(document.Pages[0].TextFrames);
        var artwork = PublisherNativeTests.Elements(document.Pages[0].Drawing).OfType<OfficeDrawingEffectGroup>()
            .Single(item => item.SourceElementIds?.Contains("publisher-object-3") == true
                || PublisherNativeTests.Elements(item.InnerDrawing).Any(element => element.SourceElementIds?.Contains("publisher-object-3") == true));
        foreach (OfficePoint unit in new[] { new OfficePoint(0, 0), new OfficePoint(1, 0), new OfficePoint(1, 1) }) {
            OfficePoint rotated = Transform(new OfficePoint(25 + unit.X * 50, 25 + unit.Y * 50), 50, 50, angle, false, false);
            OfficePoint expected = Transform(new OfficePoint(206 + rotated.X * 2, 346 + rotated.Y), 306, 396, 30, false, false);
            Equal(expected, frame.PageTransform.TransformPoint(new OfficePoint(unit.X * frame.Width, unit.Y * frame.Height)));
            OfficeDrawingShape shape = PublisherNativeTests.Elements(artwork.InnerDrawing).OfType<OfficeDrawingShape>().Single();
            Equal(expected, artwork.Transform.TransformPoint(new OfficePoint(shape.X + unit.X * shape.Shape.Width,
                shape.Y + unit.Y * shape.Shape.Height)));
        }
    }

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
        Assert.Equal(100, child.Bounds!.Value.X1, 6);
        Assert.Equal(120, child.Bounds.Value.Y1, 6);
        Assert.Equal(250, child.Bounds.Value.X2, 6);
        Assert.Equal(300, child.Bounds.Value.Y2, 6);
        var point = new OfficePoint(100, 120);
        OfficePoint inParent = Transform(new OfficePoint(200 + point.X * 1.2, 100 + point.Y), 500, 350, 30, false, true);
        OfficePoint inOuter = new(-254000 + inParent.X * 508, -127000 + inParent.Y * 254);
        OfficePoint expected = Transform(inOuter, 0, 0, 90, true, false);
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

    private static byte[] Shape(uint id, uint flags, double angle, bool child, int[] anchor, int coordinateExtent = 500, double? relativeAngle = null, bool filled = false, uint? relativeFlips = null) {
        byte[] anchorRecord = child ? Record(0xF00F, 0, Integers(anchor)) : Record(0xF010, 0,
            Client(new ushort[] { 0x2001, 0x2002, 0x2003, 0x2004 }, anchor.Select(value => unchecked((uint)value)).ToArray()));
        byte[] coordinates = (flags & 1) == 0 ? Array.Empty<byte>() : Record(0xF009, 1,
            Integers(0, 0, coordinateExtent, coordinateExtent));
        return Record(0xF004, 15, coordinates, Record(0xF00A, 2, Integers(unchecked((int)id), unchecked((int)flags))),
            Record(0xF00B, (ushort)(((1 + (relativeAngle.HasValue ? 1 : 0) + (filled ? 1 : 0) + (relativeFlips.HasValue ? 1 : 0)) << 4) | 3), new byte[] { 4, 0 }.Concat(BitConverter.GetBytes((int)(angle * 65536)))
                .Concat(relativeAngle.HasValue ? new byte[] { 0xC4, 3 }.Concat(BitConverter.GetBytes((int)(relativeAngle.Value * 65536))) : Array.Empty<byte>())
                .Concat(relativeFlips.HasValue ? new byte[] { 0xFF, 3 }.Concat(BitConverter.GetBytes(relativeFlips.Value)) : Array.Empty<byte>())
                .Concat(filled ? new byte[] { 0x81, 1, 0xFF, 0, 0, 0 } : Array.Empty<byte>()).ToArray()), anchorRecord,
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
