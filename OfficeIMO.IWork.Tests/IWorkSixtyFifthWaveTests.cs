using OfficeIMO.Excel;
using OfficeIMO.IWork;
using OfficeIMO.IWork.Internal;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Numbers_catalogs_are_source_wide_bounded_without_materialized_cells(
        bool formulaCatalog) {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Catalog", 0, 0, 0d,
                textValue: formulaCatalog ? null : string.Empty,
                duplicateString: !formulaCatalog,
                hasFormula: formulaCatalog,
                duplicateFormula: formulaCatalog,
                emptyTile: true)
        });
        IWorkSourceDocument source = IWorkSourceDocument.Open(package,
            IWorkDocumentKind.Numbers,
            new IWorkReadOptions { MaximumTableCatalogEntries = 1 });

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            source.ReadNumbers());

        Assert.Contains(formulaCatalog ? "formula catalog" : "string catalog",
            exception.Message, StringComparison.Ordinal);
        Assert.Contains("table-catalog limit of 1", exception.Message,
            StringComparison.Ordinal);
    }

    [Fact]
    public void Encrypted_pdf_previews_are_rejected() {
        byte[] pdf = CreateOnePageClassicPdf(validKids: true,
            trailerDictionaryPrefix: "/Encrypt 4 0 R ");

        Assert.False(IWorkPdfInfo.IsComplete(pdf));
    }

    [Fact]
    public void Keynote_internal_slide_links_round_trip_for_shapes_runs_and_notes() {
        using MemoryStream package = CreateKeynotePackageWithInternalSlideLinks();

        using var result = PowerPointPresentation.LoadKeynoteWithReport(package);

        Assert.False(result.IsVisualFallback);
        Assert.Equal(2, result.Document.Slides.Count);
        AssertInternalLinks(result.Document);
        Assert.Empty(result.Document.ValidateDocument());

        using var saved = new MemoryStream();
        result.Document.Save(saved);
        saved.Position = 0;
        using PowerPointPresentation reopened = PowerPointPresentation.Load(saved);
        AssertInternalLinks(reopened);
    }

    private static void AssertInternalLinks(PowerPointPresentation presentation) {
        PowerPointSlide first = presentation.Slides[0];
        PowerPointTextBox textBox = Assert.Single(first.TextBoxes);
        Assert.Equal("#slide-2", textBox.Hyperlink!.OriginalString);
        Assert.Equal("#slide-2", Assert.Single(Assert.Single(textBox.Paragraphs).Runs)
            .Hyperlink!.OriginalString);
        Assert.Equal("#slide-2", Assert.Single(Assert.Single(first.Notes.Paragraphs).Runs)
            .Hyperlink!.OriginalString);
    }

    private static MemoryStream CreateKeynotePackageWithInternalSlideLinks() {
        const ulong documentId = 1;
        const ulong showId = 2;
        const ulong firstNodeId = 3;
        const ulong secondNodeId = 4;
        const ulong firstSlideId = 5;
        const ulong secondSlideId = 6;
        const ulong shapeId = 7;
        const ulong storageId = 8;
        const ulong hyperlinkId = 9;
        const ulong noteId = 10;
        const ulong noteStorageId = 11;
        const ulong noteHyperlinkId = 12;
        byte[] slideTree = Message(
            ReferenceField(2, firstNodeId), ReferenceField(2, secondNodeId));
        byte[] hyperlinkTable = Message(BytesField(1,
            Message(VarintField(1, 0), ReferenceField(2, hyperlinkId))));
        byte[] noteHyperlinkTable = Message(BytesField(1,
            Message(VarintField(1, 0), ReferenceField(2, noteHyperlinkId))));
        byte[] drawable = Message(StringField(4, "#slide-2"));
        byte[] shape = Message(BytesField(1, Message(BytesField(1, drawable))),
            ReferenceField(2, storageId));
        byte[] records = Message(
            ArchiveRecord(documentId, 1, Message(ReferenceField(2, showId)),
                new[] { showId }),
            ArchiveRecord(showId, 2, Message(BytesField(3, slideTree)),
                new[] { firstNodeId, secondNodeId }),
            ArchiveRecord(firstNodeId, 4, Message(ReferenceField(2, firstSlideId)),
                new[] { firstSlideId }),
            ArchiveRecord(secondNodeId, 4, Message(ReferenceField(2, secondSlideId)),
                new[] { secondSlideId }),
            ArchiveRecord(firstSlideId, 5,
                Message(ReferenceField(5, shapeId), ReferenceField(27, noteId)),
                new[] { shapeId, noteId }),
            ArchiveRecord(secondSlideId, 5, Message()),
            ArchiveRecord(shapeId, 2011, shape, new[] { storageId }),
            ArchiveRecord(storageId, 2001,
                Message(StringField(3, "Linked title"), BytesField(11, hyperlinkTable)),
                new[] { hyperlinkId }),
            ArchiveRecord(hyperlinkId, 2032, Message(StringField(2, "#slide-2"))),
            ArchiveRecord(noteId, 15, Message(ReferenceField(1, noteStorageId)),
                new[] { noteStorageId }),
            ArchiveRecord(noteStorageId, 2001,
                Message(StringField(3, "Linked note"), BytesField(11, noteHyperlinkTable)),
                new[] { noteHyperlinkId }),
            ArchiveRecord(noteHyperlinkId, 2032, Message(StringField(2, "#slide-2"))));
        return CreatePackage(("Index/Document.iwa", FrameIwa(records)));
    }
}
