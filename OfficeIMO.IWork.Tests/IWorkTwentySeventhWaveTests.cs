using OfficeIMO.Excel;
using OfficeIMO.IWork;
using OfficeIMO.IWork.Internal;
using System.Text;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void New_pages_header_footer_text_is_charged_once() {
        using MemoryStream package = CreatePagesPackageWithHeaderFooterVariants();
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, IWorkDocumentKind.Pages,
            new IWorkReadOptions { MaximumProjectedTextItems = 14 });

        IWorkPagesProjection projection = source.ReadPages();

        IWorkPagesSection section = Assert.Single(projection.Sections);
        Assert.Equal(6, section.HeaderContents.Count + section.FooterContents.Count);
    }

    [Theory]
    [InlineData('\u0004')]
    [InlineData('\u0005')]
    [InlineData('\u000C')]
    public void Numbers_text_boxes_normalize_iwork_break_controls(char separator) {
        using MemoryStream package = CreateNumbersPackageWithTextBox("Before" + separator + "After");

        using var result = ExcelIWorkConverter.LoadNumbersWithReport(package);

        Assert.False(result.IsVisualFallback);
        Assert.Equal("Before\nAfter", Assert.Single(Assert.Single(
            result.Projection.Sheets).TextBoxes));
        Assert.Equal("Before\nAfter", Assert.Single(result.Document.Sheets)
            .CellAt(1, 1).GetValue<string>());
        using var saved = new MemoryStream();
        result.Document.Save(saved);
        saved.Position = 0;
        using ExcelDocument reopened = ExcelDocument.Load(saved);
        Assert.Equal("Before\nAfter", Assert.Single(reopened.Sheets)
            .CellAt(1, 1).GetValue<string>());
    }

    [Fact]
    public void Classic_pdf_outer_dictionaries_ignore_nested_terminators() {
        byte[] pdf = CreateOnePageClassicPdf(validKids: true,
            pageDictionaryPrefix:
                "/Resources << /ProcSet [/PDF] >> /Label (ignore >> and \\(nested\\)) % ignore >>\n",
            trailerDictionaryPrefix: "/Info << /Name (ignore >>) >> ");

        Assert.True(IWorkPdfInfo.IsComplete(pdf));
    }

    [Theory]
    [InlineData("/Size 4 /Info << /Root 1 0 R >>")]
    [InlineData("/Size 4 /Info /Root 1 0 R")]
    public void Classic_pdf_required_keys_must_be_outer_dictionary_keys(string replacement) {
        string pdf = Encoding.ASCII.GetString(CreateOnePageClassicPdf(validKids: true));
        pdf = pdf.Replace("/Size 4 /Root 1 0 R", replacement, StringComparison.Ordinal);

        Assert.False(IWorkPdfInfo.IsComplete(Encoding.ASCII.GetBytes(pdf)));
    }

    [Fact]
    public void Non_index_iwa_resources_are_preserved_without_being_decoded() {
        byte[] records = Message(
            ArchiveRecord(1, 10000, Message(ReferenceField(4, 2)), new[] { 2UL }),
            ArchiveRecord(2, 2001, Message(StringField(3, "Body"))));
        using MemoryStream package = CreatePackage(
            ("Index/Document.iwa", FrameIwa(records)),
            ("Data/example.iwa", new byte[] { 0xff, 0x00, 0x01 }));

        IWorkSourceDocument source = IWorkSourceDocument.Open(package, IWorkDocumentKind.Pages);

        Assert.Equal("Body", Assert.Single(source.ReadPages().Paragraphs));
        Assert.Contains(source.Entries, entry => entry.Path == "Data/example.iwa");
        Assert.Equal(2, source.Records.Count);
    }

    [Fact]
    public void Resource_only_iwa_suffixes_do_not_satisfy_the_modern_index_contract() {
        using MemoryStream package = CreatePackage(
            ("Data/example.iwa", new byte[] { 0xff, 0x00, 0x01 }));

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            IWorkSourceDocument.Open(package, IWorkDocumentKind.Pages));

        Assert.Contains("does not contain modern iWork IWA archives", exception.Message,
            StringComparison.Ordinal);
    }

    private static MemoryStream CreateNumbersPackageWithTextBox(string text) {
        const ulong documentId = 1;
        const ulong sheetId = 2;
        const ulong shapeId = 3;
        const ulong storageId = 4;
        byte[] records = Message(
            ArchiveRecord(documentId, 1, Message(ReferenceField(1, sheetId)),
                new[] { sheetId }),
            ArchiveRecord(sheetId, 2,
                Message(StringField(1, "Sheet"), ReferenceField(2, shapeId)),
                new[] { shapeId }),
            ArchiveRecord(shapeId, 2011, Message(ReferenceField(2, storageId)),
                new[] { storageId }),
            ArchiveRecord(storageId, 2001, Message(StringField(3, text))));
        return CreatePackage(("Index/Document.iwa", FrameIwa(records)));
    }
}
