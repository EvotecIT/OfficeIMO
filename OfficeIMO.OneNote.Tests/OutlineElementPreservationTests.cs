using System.Text;

namespace OfficeIMO.OneNote.Tests;

public sealed class OutlineElementPreservationTests {
    [Fact]
    public void PreservesNonParagraphWrapperMetadataAndNestedContent() {
        var image = new OneNoteImage {
            FileName = "pixel.png",
            Payload = OneNoteBinaryPayload.FromBytes(new byte[] { 1, 2, 3 })
        };
        var nested = new OneNoteParagraph();
        nested.Runs.Add(new OneNoteTextRun { Text = "nested" });
        var page = new OneNotePage { Title = "Wrapped image" };
        page.DirectContent.Add(image);
        page.DirectContent.Add(nested);
        var section = new OneNoteSection { Name = "Wrappers" };
        section.Pages.Add(page);

        OneNoteWriteGraph graph = new OneNoteWriteGraphBuilder().BuildSection(section);
        OneNoteWriteObjectSpace pageSpace = graph.ObjectSpaces[1];
        OneNoteWriteObject pageNode = Assert.Single(pageSpace.Objects, item => item.Jcid == OneNoteSchema.JcidPageNode);
        OneNoteExtendedGuid outlineId = Assert.Single(Property(pageNode, OneNoteSchema.ElementChildNodes).References);
        OneNoteWriteObject outline = Assert.Single(pageSpace.Objects, item => item.Id == outlineId);
        OneNoteExtendedGuid[] directIds = Property(outline, OneNoteSchema.ElementChildNodes).References.ToArray();
        Assert.Equal(2, directIds.Length);
        OneNoteExtendedGuid imageId = directIds[0];
        OneNoteExtendedGuid nestedId = directIds[1];

        var listId = new OneNoteExtendedGuid(Guid.NewGuid(), 1, 17);
        pageSpace.Objects.Add(new OneNoteWriteObject(
            listId,
            OneNoteSchema.JcidNumberListNode,
            new[] {
                new OneNoteWriteProperty(OneNoteSchema.NumberListFormat, data: Encoding.Unicode.GetBytes("\u0002\uFFFD\u0003")),
                new OneNoteWriteProperty(OneNoteSchema.ListRestart, scalar: 3)
            }));

        var authorId = new OneNoteExtendedGuid(Guid.NewGuid(), 1, 17);
        pageSpace.Objects.Add(new OneNoteWriteObject(
            authorId,
            OneNoteSchema.JcidAuthor,
            new[] { new OneNoteWriteProperty(OneNoteSchema.Author, data: Encoding.Unicode.GetBytes("Wrapper author\0")) }));
        var wrapperId = new OneNoteExtendedGuid(Guid.NewGuid(), 1, 17);
        pageSpace.Objects.Add(new OneNoteWriteObject(
            wrapperId,
            OneNoteSchema.JcidOutlineElementNode,
            new[] {
                new OneNoteWriteProperty(OneNoteSchema.OffsetFromParentHorizontal, scalar: FloatBits(12.5F)),
                new OneNoteWriteProperty(OneNoteSchema.OffsetFromParentVertical, scalar: FloatBits(7.25F)),
                new OneNoteWriteProperty(OneNoteSchema.ContentChildNodes, references: new[] { imageId }),
                new OneNoteWriteProperty(OneNoteSchema.ElementChildNodes, references: new[] { nestedId }),
                new OneNoteWriteProperty(OneNoteSchema.ListNodes, references: new[] { listId }),
                new OneNoteWriteProperty(OneNoteSchema.OutlineElementChildLevel, scalar: 2),
                new OneNoteWriteProperty(OneNoteSchema.AuthorMostRecent, references: new[] { authorId })
            }));

        int pageNodeIndex = pageSpace.Objects.IndexOf(pageNode);
        var pageProperties = pageNode.Properties
            .Where(property => (property.RawId & 0x7FFFFFFFU) != OneNoteSchema.ElementChildNodes)
            .Concat(new[] { new OneNoteWriteProperty(OneNoteSchema.ElementChildNodes, references: new[] { wrapperId }) })
            .ToArray();
        pageSpace.Objects[pageNodeIndex] = new OneNoteWriteObject(pageNode.Id, pageNode.Jcid, pageProperties);

        OneNoteSection loaded = OneNoteSectionReader.Read(new MemoryStream(OneNoteRevisionStoreWriter.Write(graph)));
        AssertWrapper(loaded, authorId, canonicalRoot: false);

        byte[] rewritten = OneNoteSectionWriter.Write(loaded, new OneNoteWriterOptions { PreserveUnknownData = false });
        AssertWrapper(OneNoteSectionReader.Read(new MemoryStream(rewritten)), authorId, canonicalRoot: true);
    }

    private static void AssertWrapper(OneNoteSection section, OneNoteExtendedGuid expectedAuthorId, bool canonicalRoot) {
        OneNotePage page = Assert.Single(section.Pages);
        OneNoteOutline wrapper;
        if (canonicalRoot) {
            Assert.Empty(page.DirectContent);
            wrapper = Assert.IsType<OneNoteOutline>(Assert.Single(Assert.Single(page.Outlines).Children));
        } else {
            Assert.Empty(page.Outlines);
            wrapper = Assert.IsType<OneNoteOutline>(Assert.Single(page.DirectContent));
        }
        Assert.True(wrapper.IsOutlineElementWrapper);
        Assert.Equal(12.5, wrapper.Layout!.X);
        Assert.Equal(7.25, wrapper.Layout.Y);
        Assert.Equal("Wrapper author", wrapper.Author!.Name);
        Assert.Equal(expectedAuthorId, wrapper.Author.ObjectId);
        Assert.True(wrapper.WrapperList!.Ordered);
        Assert.Equal(1, wrapper.WrapperList.Level);
        Assert.Collection(wrapper.Children,
            child => Assert.IsType<OneNoteImage>(child),
            child => Assert.Equal("nested", Assert.Single(Assert.IsType<OneNoteParagraph>(child).Runs).Text));
    }

    private static OneNoteWriteProperty Property(OneNoteWriteObject item, uint id) =>
        Assert.Single(item.Properties, property => (property.RawId & 0x7FFFFFFFU) == id);

    private static uint FloatBits(float value) => BitConverter.ToUInt32(BitConverter.GetBytes(value), 0);
}
