using OfficeIMO.Core.Internal;
using OfficeIMO.Publisher.Internal;
using OfficeIMO.TestSupport;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherPageRoleTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Page_names_do_not_change_native_physical_or_master_roles(bool emptyMasterName) {
        byte[] original = File.ReadAllBytes(PublisherNativeTests.Fixture("Sample.pub"));
        PublisherDocument baseline = PublisherDocument.Load(original);
        uint pageId = emptyMasterName ? Assert.Single(baseline.MasterPages).Id : baseline.Pages[0].Id;
        string name = emptyMasterName ? string.Empty : "Named physical page";
        PublisherDocument changed = PublisherDocument.Load(PublisherTableFixture.NamePage(original, pageId, name));

        Assert.Equal(baseline.Pages.Select(page => page.Id), changed.Pages.Select(page => page.Id));
        Assert.Equal(baseline.MasterPages.Select(page => page.Id), changed.MasterPages.Select(page => page.Id));
        PublisherPage renamed = Assert.Single(changed.Pages.Concat(changed.MasterPages), page => page.Id == pageId);
        Assert.Equal(name, renamed.Name);
    }

    [Theory]
    [InlineData(0x2C, 0)]
    [InlineData(0x2C, 7)]
    [InlineData(0x2D, 2)]
    [InlineData(1, 7)]
    public void Invalid_native_page_role_ranges_are_rejected_without_guessing_pages(int fieldId, int value) {
        byte[] bytes = File.ReadAllBytes(PublisherNativeTests.Fixture("Simple.pub"));
        Assert.True(OfficeCompoundFileReader.TryRead(bytes, out OfficeCompoundFile? source, out string? error), error);
        byte[] contents = (byte[])source!.Streams["Contents"].Clone();
        var data = new PublisherBinaryData(contents, "Contents");
        var context = new PublisherParseContext(new PublisherReadOptions(), CancellationToken.None);
        PublisherContents publication = new PublisherContentsReader(data, context).Read();
        PublisherContentChunk document = Assert.Single(publication.Chunks.Values, chunk => chunk.Kind == 0x44);
        var fields = new PublisherBlockReader(data, context).Chunk(document.Offset, document.End);
        PublisherBlock field = Assert.Single(fields, item => item.Id == fieldId);
        byte[] replacement = field.Length == 2 ? BitConverter.GetBytes(checked((ushort)value)) : BitConverter.GetBytes(value);
        Assert.Equal(field.Length, replacement.Length);
        replacement.CopyTo(contents, field.Offset);
        byte[] malformed = OfficeCompoundFileWriter.Rewrite(source,
            new Dictionary<string, byte[]> { ["Contents"] = contents });

        Assert.Throws<InvalidDataException>(() => PublisherDocument.Load(malformed));
    }
}
