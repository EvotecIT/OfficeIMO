using OfficeIMO.Reader;
using OfficeIMO.Reader.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData("")]
    [InlineData("Linked shape")]
    public void Reader_preserves_Pages_text_shape_hyperlinks(string text) {
        const string Target = "https://example.test/pages-shape";
        using MemoryStream package = CreatePagesPackage(includeBody: true, textBox: text,
            includePreview: true, textBoxDrawable: Message(StringField(4, Target)));
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddIWorkHandler().Build();

        OfficeDocumentReadResult result = reader.ReadDocument(package, "linked.pages");

        Assert.Equal(Target, Assert.Single(result.Links).Uri);
        Assert.Same(result.Links[0], Assert.Single(result.Pages[0].Links));
    }

    [Theory]
    [InlineData("")]
    [InlineData("Linked shape")]
    public void Reader_preserves_Keynote_text_shape_hyperlinks(string text) {
        const string Target = "https://example.test/keynote-shape";
        using MemoryStream package = CreateKeynotePackageWithRepeatedSlides(1,
            text: text, textBoxDrawable: Message(StringField(4, Target)));
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddIWorkHandler().Build();

        OfficeDocumentReadResult result = reader.ReadDocument(package, "linked.key");

        Assert.Equal(Target, Assert.Single(result.Links).Uri);
        Assert.Same(result.Links[0], Assert.Single(result.Pages[0].Links));
    }
}
