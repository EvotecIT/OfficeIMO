using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;
using OfficeIMO.Drawing.Binary;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.Tests;

// Model-only admission and wrap variants decoded from independent Publisher
// input. These do not qualify native writing, native edits or native appearance.
internal static class PublisherTableEditingFixture {
    internal static PublisherDocument Load(string? variant = null, bool withWrap = false, PublisherReadOptions? options = null) {
        byte[] bytes = File.ReadAllBytes(PublisherNativeTests.Fixture("Sample.pub"));
        Assert.True(OfficeCompoundFileReader.TryRead(bytes, out OfficeCompoundFile? compound, out string? error), error);
        var context = new PublisherParseContext((options ?? new PublisherReadOptions()).Clone(), default);
        var data = new PublisherBinaryData(compound!.Streams["Contents"], "Contents");
        PublisherContents source = new PublisherContentsReader(data, context).Read();
        PublisherTextFrameCatalogReader.Read(data, source, context);
        PublisherQuillText text = new PublisherQuillTextReader(new PublisherBinaryData(compound.Streams["Quill/QuillSub/CONTENTS"], "Quill"), context).Read(source.Palette);
        compound.Streams.TryGetValue("Escher/EscherDelayStm", out byte[]? delay);
        PublisherEscherData artwork = new PublisherEscherReader(new PublisherBinaryData(compound.Streams["Escher/EscherStm"], "Escher"),
            delay == null ? null : new PublisherBinaryData(delay, "EscherDelay"), context).Read();
        PublisherSourceShape native = source.Shapes[299];
        uint storyId = native.Value(0x27)!.Value, pageId = source.ObjectPages[299];
        PublisherEscherShape tableGraphic = artwork.Shapes.Single(shape => shape.Id == 299);
        if (variant is "hidden" or "rotation" or "gradient" or "no-anchor" or "unknown-preset" or "custom-path") {
            var properties = tableGraphic.Properties.ToList();
            if (variant == "hidden") properties.Add(new OfficeArtProperty(properties.Count, 0x03BF, 0x40004000));
            if (variant == "rotation") properties.Add(new OfficeArtProperty(properties.Count, 4, 30 * 65536));
            if (variant == "gradient") {
                properties.Add(new OfficeArtProperty(properties.Count, 0x0180, 4));
                properties.Add(new OfficeArtProperty(properties.Count, 0x01BF, 0x00100010));
                properties.Add(new OfficeArtProperty(properties.Count, 0x0181, 0x000000FF));
                properties.Add(new OfficeArtProperty(properties.Count, 0x0183, 0x00FF0000));
            }
            if (variant == "custom-path") properties.Add(new OfficeArtProperty(properties.Count, 0x0145, 1));
            artwork.Shapes[artwork.Shapes.IndexOf(tableGraphic)] = CopyGraphic(tableGraphic, 299, properties, variant != "no-anchor",
                variant == "unknown-preset" ? 100 : null);
        }
        if (variant is "shared" or "conflict" or "shared-missing-grid" or "shared-unresolved-owner") {
            PublisherSourceTable grid = native.Table!;
            if (variant is "conflict" or "shared-unresolved-owner") {
                grid = new PublisherSourceTable(grid.Columns, grid.Rows);
                grid.Cells.Add(new PublisherSourceCell(0, grid.Rows.Length - 1, 0, grid.Columns.Length - 1));
            }
            source.Shapes.Add(900, new PublisherSourceShape(new PublisherContentChunk(900, 0x10, 0,
                variant == "shared-unresolved-owner" ? null : pageId), native.Fields) {
                Table = variant == "shared-missing-grid" ? null : grid
            });
            if (variant != "shared-unresolved-owner") source.ObjectPages.Add(900, pageId);
            artwork.Shapes.Add(CopyGraphic(tableGraphic, 900, tableGraphic.Properties));
        }
        if (variant == "unresolved") text.TableStories.Remove(storyId);
        if (variant == "partition-conflict") {
            var segments = text.TableStories[storyId].Segments.ToArray();
            segments[0] = new PublisherTextStory(storyId, segments[0].Text.TrimEnd('\n'), segments[0].Paragraphs, false);
            text.TableStories[storyId] = new PublisherTableStory(segments);
        }
        if (variant == "one-id") {
            source.Shapes.Add(uint.MaxValue - 1, new PublisherSourceShape(
                new PublisherContentChunk(uint.MaxValue - 1, 1, 0, pageId), Array.Empty<PublisherBlock>()));
            source.ObjectPages.Add(uint.MaxValue - 1, pageId);
        }
        if (variant is "ordinary-frame" || withWrap) {
            uint frameStory = withWrap ? 901U : storyId;
            if (withWrap) {
                string body = string.Concat(Enumerable.Repeat("Table movement changes the existing text exclusion. ", 25));
                text.Stories.Add(frameStory, new PublisherTextStory(frameStory, body, new[] {
                    new OfficeRichTextParagraph(new[] { new OfficeRichTextRun(body, 10, OfficeColor.Black, fontFamily: "Arial") })
                }));
            }
            source.Shapes.Add(902, new PublisherSourceShape(new PublisherContentChunk(902, 1, 0, pageId), new[] {
                new PublisherBlock(0x27, 0x20, 0, 4, frameStory)
            }) { WrapObjects = withWrap ? new uint[] { 299 } : Array.Empty<uint>() });
            source.ObjectPages.Add(902, pageId);
            artwork.Shapes.Add(new PublisherEscherShape(902, 202, 0, Array.Empty<OfficeArtProperty>()) {
                Bounds = new PublisherNativeRectangle((12 - source.Width / 2) * 12700, (12 - source.Height / 2) * 12700,
                    (400 - source.Width / 2) * 12700, (400 - source.Height / 2) * 12700)
            });
        }
        return new PublisherModelState(source, text, artwork, context, new PublisherNativeSource(bytes, compound)).Project(default);
    }

    private static PublisherEscherShape CopyGraphic(PublisherEscherShape source, uint id, IReadOnlyList<OfficeArtProperty> properties,
        bool anchor = true, int? type = null) =>
        new(id, type ?? source.Type, source.Flags, properties, source.HasParentDrawingSpace) {
            Bounds = anchor ? source.Bounds : null, GroupCoordinates = source.GroupCoordinates, GroupTransform = source.GroupTransform,
            LayoutTransform = source.LayoutTransform, HasPublicationGroupParent = source.HasPublicationGroupParent, HiddenByGroup = source.HiddenByGroup
        };
}
