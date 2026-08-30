using OfficeIMO.Excel;
using OfficeIMO.IWork;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Numbers_text_boxes_preserve_significant_outer_whitespace() {
        using MemoryStream package = CreateNumbersPackage(Array.Empty<TableSpec>(),
            textBox: "  retained  ");

        IWorkNumbersProjection projection = IWorkSourceDocument.Open(
            package, IWorkDocumentKind.Numbers).ReadNumbers();

        Assert.Equal("  retained  ", Assert.Single(Assert.Single(projection.Sheets).TextBoxes));
    }

    [Fact]
    public void Keynote_table_rotation_is_applied_to_the_native_powerpoint_table() {
        using MemoryStream package = CreateKeynotePackageWithLargeTables(1,
            rotation: 30f, rows: 1, columns: 1);

        using var result = PowerPointPresentation.LoadKeynoteWithReport(package);

        Assert.Equal(30d, Assert.Single(Assert.Single(result.Document.Slides).Tables).Rotation);
        using var bytes = new MemoryStream();
        result.Document.Save(bytes);
        bytes.Position = 0;
        using PowerPointPresentation reopened = PowerPointPresentation.Load(bytes);
        Assert.Equal(30d, Assert.Single(Assert.Single(reopened.Slides).Tables).Rotation);
    }

    [Fact]
    public void Pages_nested_lists_use_native_word_numbering_and_level() {
        using MemoryStream package = CreatePagesPackageWithResolvedNestedList();

        using var result = WordDocument.LoadPagesWithReport(package);

        WordParagraph paragraph = Assert.Single(result.Document.Paragraphs,
            candidate => candidate.Text == "Item");
        Assert.True(paragraph.IsListItem);
        Assert.Equal(1, paragraph.ListItemLevel);
        using var bytes = new MemoryStream();
        result.Document.Save(bytes);
        bytes.Position = 0;
        using WordDocument reopened = WordDocument.Load(bytes);
        WordParagraph persisted = Assert.Single(reopened.Paragraphs,
            candidate => candidate.Text == "Item");
        Assert.True(persisted.IsListItem);
        Assert.Equal(1, persisted.ListItemLevel);
    }

    [Fact]
    public void Rendered_numbers_formulas_are_charged_to_the_source_wide_text_budget() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("T", 1, 1, 42d, hasFormula: true)
        });
        var options = new IWorkReadOptions { MaximumProjectedTextCharacters = 7 };

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            IWorkSourceDocument.Open(package, IWorkDocumentKind.Numbers, options).ReadNumbers());

        Assert.Contains("text character count", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void CrLf_is_one_pages_paragraph_separator() {
        using MemoryStream package = CreatePagesPackageWithStyleChain(depth: 1,
            bodyText: "First\r\nSecond");

        IWorkPagesProjection projection = IWorkSourceDocument.Open(
            package, IWorkDocumentKind.Pages).ReadPages();

        Assert.Collection(projection.Body.Paragraphs,
            paragraph => Assert.Equal("First", Assert.Single(paragraph.Runs).Text),
            paragraph => Assert.Equal("Second", Assert.Single(paragraph.Runs).Text));
    }

    [Fact]
    public void Out_of_range_color_components_disable_editable_pages_reconstruction() {
        using MemoryStream package = CreatePagesPackageWithColor(1.5f);

        IWorkPagesProjection projection = IWorkSourceDocument.Open(
            package, IWorkDocumentKind.Pages).ReadPages();

        Assert.False(projection.HasEditableContent);
        Assert.Contains(projection.Diagnostics,
            diagnostic => diagnostic.Code == "IWORK_PAGES_TEXT_UNSUPPORTED");
    }

    [Fact]
    public void Text_attribute_boundaries_cannot_split_utf16_surrogate_pairs() {
        using MemoryStream package = CreatePagesPackageWithSurrogateSplitBoundary();

        IWorkPagesProjection projection = IWorkSourceDocument.Open(
            package, IWorkDocumentKind.Pages).ReadPages();

        Assert.False(projection.HasEditableContent);
        Assert.Equal("😀", Assert.Single(Assert.Single(projection.Body.Paragraphs).Runs).Text);
    }

    [Fact]
    public void Overlapping_numbers_header_and_footer_regions_disable_editable_reconstruction() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Overlap", 2, 1, 42d, headerRows: 2, footerRows: 1)
        });

        IWorkNumbersProjection projection = IWorkSourceDocument.Open(
            package, IWorkDocumentKind.Numbers).ReadNumbers();

        Assert.False(projection.HasEditableContent);
        Assert.Contains(projection.Diagnostics,
            diagnostic => diagnostic.Code == "IWORK_TABLE_REGIONS_UNSUPPORTED");
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages, 10000u, "IWORK_PAGES_DOCUMENT_DUPLICATE")]
    [InlineData(IWorkDocumentKind.Numbers, 1u, "IWORK_NUMBERS_DOCUMENT_DUPLICATE")]
    [InlineData(IWorkDocumentKind.Keynote, 1u, "IWORK_KEYNOTE_DOCUMENT_DUPLICATE")]
    public void Duplicate_application_roots_are_rejected(IWorkDocumentKind kind,
        uint messageType, string diagnosticCode) {
        byte[] records = Message(
            ArchiveRecord(1, messageType, Message()),
            ArchiveRecord(2, messageType, Message()));
        using MemoryStream package = CreatePackage(
            ("Index/Document.iwa", FrameIwa(records)));

        IWorkDiagnostic[] diagnostics = kind switch {
            IWorkDocumentKind.Pages => IWorkSourceDocument.Open(package, kind).ReadPages().Diagnostics.ToArray(),
            IWorkDocumentKind.Numbers => IWorkSourceDocument.Open(package, kind).ReadNumbers().Diagnostics.ToArray(),
            IWorkDocumentKind.Keynote => IWorkSourceDocument.Open(package, kind).ReadKeynote().Diagnostics.ToArray(),
            _ => throw new InvalidOperationException()
        };

        Assert.Contains(diagnostics, diagnostic => diagnostic.Code == diagnosticCode);
    }

    private static MemoryStream CreatePagesPackageWithColor(float red) {
        const ulong documentId = 1;
        const ulong bodyId = 2;
        const ulong styleId = 3;
        byte[] styleTable = Message(BytesField(1,
            Message(VarintField(1, 0), ReferenceField(2, styleId))));
        byte[] records = Message(
            ArchiveRecord(documentId, 10000, Message(ReferenceField(4, bodyId)), new[] { bodyId }),
            ArchiveRecord(bodyId, 2001,
                Message(StringField(3, "Color"), BytesField(8, styleTable)), new[] { styleId }),
            ArchiveRecord(styleId, 2021,
                Message(BytesField(11, Message(BytesField(7, Message(FloatField(3, red))))))));
        return CreatePackage(("Index/Document.iwa", FrameIwa(records)));
    }

    private static MemoryStream CreatePagesPackageWithSurrogateSplitBoundary() {
        const ulong documentId = 1;
        const ulong bodyId = 2;
        const ulong firstStyleId = 3;
        const ulong secondStyleId = 4;
        byte[] styleTable = Message(
            BytesField(1, Message(VarintField(1, 0), ReferenceField(2, firstStyleId))),
            BytesField(1, Message(VarintField(1, 1), ReferenceField(2, secondStyleId))));
        byte[] records = Message(
            ArchiveRecord(documentId, 10000, Message(ReferenceField(4, bodyId)), new[] { bodyId }),
            ArchiveRecord(bodyId, 2001,
                Message(StringField(3, "😀"), BytesField(8, styleTable)),
                new[] { firstStyleId, secondStyleId }),
            ArchiveRecord(firstStyleId, 2021, Message()),
            ArchiveRecord(secondStyleId, 2021, Message(BytesField(11, Message(VarintField(1, 1))))));
        return CreatePackage(("Index/Document.iwa", FrameIwa(records)));
    }
}
