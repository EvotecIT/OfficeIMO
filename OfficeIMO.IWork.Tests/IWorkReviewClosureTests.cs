using System.IO.Compression;
using OfficeIMO.Excel;
using OfficeIMO.IWork;
using OfficeIMO.IWork.Internal;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Crc_valid_pngs_with_invalid_image_streams_are_not_selected_as_previews() {
        using FileStream input = File.OpenRead(Fixture("nim-iwork/simple.pages"));
        using var fixture = new ZipArchive(input, ZipArchiveMode.Read, leaveOpen: false);
        byte[] validJpeg = ReadEntry(fixture, "preview.jpg");
        byte[] records = Message(ArchiveRecord(1, 10000, Message()));
        using MemoryStream package = CreatePackage(
            ("Index/Document.iwa", FrameIwa(records)),
            ("preview.png", CreateCrcValidPngWithInvalidImageData()),
            ("preview-web.jpg", validJpeg));

        IWorkSourceDocument source = IWorkSourceDocument.Open(package, IWorkDocumentKind.Pages);

        Assert.DoesNotContain(source.Previews, preview => preview.Path == "preview.png");
        Assert.Equal("preview-web.jpg", source.PreferredRasterPreview!.Path);
    }

    [Fact]
    public void Duplicate_numbers_formula_identifiers_retain_only_the_cached_value() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Ambiguous formula", 1, 1, 42d,
                hasFormula: true, duplicateFormula: true)
        });

        using var result = ExcelDocument.LoadNumbersWithReport(package);
        IWorkTableCell cell = Assert.Single(Assert.Single(
            Assert.Single(result.Projection.Sheets).Tables).Cells);

        Assert.False(result.IsVisualFallback);
        Assert.False(cell.FormulaIsComplete);
        Assert.Equal(42d, Assert.IsType<double>(cell.Value), 10);
        Assert.Equal(42d, result.Document.Sheets[0].CellAt(1, 1).GetValue<double>(), 10);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_TABLE_FORMULA_STORAGE_UNSUPPORTED");
    }

    [Fact]
    public void Wrong_wire_numbers_dimensions_disable_editable_reconstruction() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Malformed dimensions", 1, 1, 42d, wrongWireDimensions: true)
        }, includePreview: true);

        using var result = ExcelDocument.LoadNumbersWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_TABLE_DIMENSIONS_UNSUPPORTED");
    }

    [Fact]
    public void Pages_natural_alignment_maps_to_the_word_logical_start_edge() {
        using MemoryStream package = CreatePagesPackageWithStyleChain(
            depth: 1, naturalAlignment: true);

        using var result = WordDocument.LoadPagesWithReport(package);
        WordParagraph paragraph = Assert.Single(result.Document.Paragraphs,
            candidate => candidate.Text == "Styled");

        Assert.Equal(IWorkTextAlignment.Natural,
            Assert.Single(result.Projection.Body.Paragraphs).Style.Alignment);
        Assert.Equal(WordParagraphAlignment.Start, paragraph.ParagraphAlignment);
    }

    [Fact]
    public void Keynote_presenter_note_hyperlinks_round_trip_through_powerpoint_notes() {
        using MemoryStream package = CreateKeynotePackageWithLinkedNotes();
        var target = new Uri("https://example.com/keynote-note");

        using var result = PowerPointPresentation.LoadKeynoteWithReport(package);
        PowerPointTextRun run = Assert.Single(Assert.Single(
            Assert.Single(result.Document.Slides).Notes.Paragraphs).Runs);

        Assert.False(result.IsVisualFallback);
        Assert.Equal("Linked note", run.Text);
        Assert.Equal(target, run.Hyperlink);
        Assert.Empty(result.Document.ValidateDocument());

        using var saved = new MemoryStream();
        result.Document.Save(saved);
        saved.Position = 0;
        using PowerPointPresentation reopened = PowerPointPresentation.Load(saved);
        PowerPointTextRun reopenedRun = Assert.Single(Assert.Single(
            Assert.Single(reopened.Slides).Notes.Paragraphs).Runs);
        Assert.Equal(target, reopenedRun.Hyperlink);
    }

    [Fact]
    public void Formula_renderer_parenthesizes_left_nested_exponentiation() {
        byte[] nodeArray = Message(
            BytesField(1, Message(VarintField(1, 17), DoubleField(4, 2d))),
            BytesField(1, Message(VarintField(1, 17), DoubleField(4, 3d))),
            BytesField(1, Message(VarintField(1, 5))),
            BytesField(1, Message(VarintField(1, 17), DoubleField(4, 4d))),
            BytesField(1, Message(VarintField(1, 5))));
        IWorkWireMessage formula = IWorkProtobuf.Parse(
            Message(BytesField(1, nodeArray)), new IWorkReadOptions());

        IWorkFormulaResult result = IWorkFormulaReader.Render(formula, 0, 0, 32, 128);

        Assert.True(result.IsComplete);
        Assert.Equal("=(2^3)^4", result.Text);
    }

    [Fact]
    public void Incomplete_uncached_numbers_formulas_use_visual_fallback() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Uncached formula", 1, 1, 0d, hasFormula: true,
                formulaWithoutCachedValue: true)
        }, includePreview: true);

        using var result = ExcelDocument.LoadNumbersWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_TABLE_FORMULA_UNSUPPORTED");
    }

    [Fact]
    public void Numbers_owner_isolates_tables_and_applies_default_width_in_constant_space() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Wide", 1, 16_384, 1d, defaultColumnWidth: 70d),
            new TableSpec("Narrow", 1, 1, 2d, defaultColumnWidth: 140d)
        });

        using var result = ExcelDocument.LoadNumbersWithReport(package);

        Assert.False(result.IsVisualFallback);
        Assert.Equal(2, result.Document.Sheets.Count);
        Assert.Equal(1d, result.Document.Sheets[0].CellAt(1, 1).GetValue<double>(), 10);
        Assert.Equal(2d, result.Document.Sheets[1].CellAt(1, 1).GetValue<double>(), 10);
        Assert.Equal(10d, result.Document.Sheets[0].DefaultColumnWidth);
        Assert.Equal(20d, result.Document.Sheets[1].DefaultColumnWidth);
        Assert.Empty(result.Document.Sheets[0].GetColumnDefinitions());
        Assert.Empty(result.Document.Sheets[1].GetColumnDefinitions());
    }

    [Fact]
    public void Numbers_owner_keeps_sheet_text_separate_from_table_coordinates() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Formula coordinates", 1, 1, 42d)
        }, textBox: "Sheet annotation");

        using var result = ExcelDocument.LoadNumbersWithReport(package);

        Assert.False(result.IsVisualFallback);
        Assert.Equal(2, result.Document.Sheets.Count);
        Assert.Equal("Sheet annotation", result.Document.Sheets[0].CellAt(1, 1).GetValue<string>());
        Assert.Equal(42d, result.Document.Sheets[1].CellAt(1, 1).GetValue<double>(), 10);
    }

    [Fact]
    public void Pages_owner_uses_a_default_marker_for_unlabeled_lists() {
        using MemoryStream package = CreatePagesPackageWithUnlabeledList();

        IWorkPagesProjection projection = IWorkSourceDocument.Open(
            package, IWorkDocumentKind.Pages).ReadPages();
        IWorkTextParagraph sourceParagraph = Assert.Single(projection.Body.Paragraphs);
        Assert.Equal(0, sourceParagraph.ListLevel);
        Assert.Null(sourceParagraph.ListLabel);
        package.Position = 0;

        using var result = WordDocument.LoadPagesWithReport(package);

        string[] texts = result.Document.Paragraphs.Select(paragraph => paragraph.Text).ToArray();
        Assert.Contains("\u2022 ", texts);
        Assert.Contains("Item", texts);
    }

    [Fact]
    public void Pages_drawable_hyperlinks_use_visual_fallback_instead_of_silent_loss() {
        using MemoryStream package = CreatePagesPackage(includeBody: true,
            textBox: "Linked shape", includePreview: true,
            textBoxDrawable: Message(StringField(4, "https://example.com/pages-shape")));

        using var result = WordDocument.LoadPagesWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.Equal("https://example.com/pages-shape",
            Assert.Single(result.Projection.TextBoxObjects).Hyperlink);
    }

    [Fact]
    public void PowerPoint_shape_hyperlinks_round_trip_through_the_owner_model() {
        using PowerPointPresentation presentation = PowerPointPresentation.Create();
        PowerPointTextBox textBox = presentation.AddSlide()
            .AddTextBoxPoints("Linked", 10, 10, 100, 30);
        var target = new Uri("https://example.com/keynote-shape");

        textBox.SetHyperlink(target);

        Assert.Equal(target, textBox.Hyperlink);
        Assert.Empty(presentation.ValidateDocument());
        textBox.ClearHyperlink();
        Assert.Null(textBox.Hyperlink);
    }

    [Fact]
    public void Singular_protobuf_fields_use_the_last_wire_value() {
        byte[] payload = Message(VarintField(1, 42));
        using MemoryStream package = CreatePackage(("Index/Document.iwa",
            FrameIwa(ArchiveRecordWithRepeatedSingularFields(1, 1, payload))));

        IWorkSourceDocument source = IWorkSourceDocument.Open(package, IWorkDocumentKind.Numbers);
        IWorkArchiveRecord record = Assert.Single(source.Records);

        Assert.Equal(1ul, record.Identifier);
        Assert.Equal(1u, record.MessageType);
        Assert.Equal(payload, record.GetPayload());
    }

    [Fact]
    public void Character_style_cache_keeps_distinct_inherited_paragraph_defaults() {
        using MemoryStream package = CreatePagesPackageWithSharedCharacterStyle();

        IWorkTextParagraph[] paragraphs = IWorkSourceDocument.Open(package, IWorkDocumentKind.Pages)
            .ReadPages().Body.Paragraphs.ToArray();

        Assert.Equal(2, paragraphs.Length);
        Assert.True(Assert.Single(paragraphs[0].Runs).Style.Bold);
        Assert.True(Assert.Single(paragraphs[1].Runs).Style.Bold);
        Assert.Equal(10d, paragraphs[0].Runs[0].Style.FontSizePoints);
        Assert.Equal(20d, paragraphs[1].Runs[0].Style.FontSizePoints);
    }

    [Fact]
    public void Xml_invalid_scalar_metadata_disables_editable_reconstruction() {
        using MemoryStream package = CreateNumbersPackage(Array.Empty<TableSpec>(),
            includePreview: true, sheetNameBytes: new byte[] { 0x01 });

        using var result = ExcelDocument.LoadNumbersWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics,
            diagnostic => diagnostic.Code == "IWORK_NUMBERS_TEXT_UNSUPPORTED");
    }

    [Fact]
    public void One_sided_modern_numbers_cell_storage_disables_editable_reconstruction() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Incomplete modern row", 1, 1, 42d, omitCurrentOffsets: true)
        }, includePreview: true);

        using var result = ExcelDocument.LoadNumbersWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_TABLE_CELL_STORAGE_UNSUPPORTED");
    }

    [Fact]
    public void Numbers_dates_outside_the_excel_range_use_visual_fallback() {
        double ancientDate = (new DateTime(1, 1, 1, 0, 0, 0, DateTimeKind.Utc)
            - new DateTime(2001, 1, 1, 0, 0, 0, DateTimeKind.Utc)).TotalSeconds;
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Date", 1, 1, ancientDate, date: true),
            new TableSpec("Formula date", 1, 1, ancientDate, hasFormula: true, date: true)
        }, includePreview: true);

        using var result = ExcelDocument.LoadNumbersWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.All(Assert.Single(result.Projection.Sheets).Tables,
            table => Assert.IsType<DateTime>(Assert.Single(table.Cells).Value));
    }

    [Fact]
    public void Wrong_pages_text_box_storage_types_disable_editable_reconstruction() {
        using MemoryStream package = CreatePagesPackage(includeBody: true, textBox: "Not text storage",
            includePreview: true, textBoxStorageType: 9999);

        using var result = WordDocument.LoadPagesWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_PAGES_DRAWABLE_UNSUPPORTED");
    }

    [Fact]
    public void Pages_owner_preserves_blank_paragraphs_and_significant_spaces() {
        using MemoryStream package = CreatePagesPackage(includeBody: true, textBox: null,
            includePreview: false, bodyText: " First \n\nSecond");

        using var result = WordDocument.LoadPagesWithReport(package);

        Assert.Equal(new[] { " First ", string.Empty, "Second" },
            result.Projection.Body.Paragraphs.Select(paragraph => paragraph.Text));
        AssertBlankParagraphBetween(result.Document.Paragraphs.Select(paragraph => paragraph.Text));
        using var bytes = new MemoryStream();
        result.Document.Save(bytes);
        bytes.Position = 0;
        using WordDocument reopened = WordDocument.Load(bytes);
        AssertBlankParagraphBetween(reopened.Paragraphs.Select(paragraph => paragraph.Text));
    }

    [Fact]
    public void Pages_owner_breaks_header_inheritance_for_an_explicitly_empty_section() {
        using MemoryStream package = CreatePagesPackageWithTwoSections(emptySecondSection: true);

        using var result = WordDocument.LoadPagesWithReport(package);

        Assert.Contains(result.Document.Sections[0].Header.Default!.Paragraphs,
            paragraph => paragraph.Text == "First header");
        Assert.NotNull(result.Document.Sections[1].Header.Default);
        Assert.DoesNotContain(result.Document.Sections[1].Header.Default!.Paragraphs,
            paragraph => paragraph.Text == "First header");
        using var bytes = new MemoryStream();
        result.Document.Save(bytes);
        bytes.Position = 0;
        using WordDocument reopened = WordDocument.Load(bytes);
        Assert.NotNull(reopened.Sections[1].Header.Default);
        Assert.DoesNotContain(reopened.Sections[1].Header.Default!.Paragraphs,
            paragraph => paragraph.Text == "First header");
    }

    private static void AssertBlankParagraphBetween(IEnumerable<string> paragraphTexts) {
        string[] texts = paragraphTexts.ToArray();
        int first = Array.IndexOf(texts, " First ");
        int second = Array.IndexOf(texts, "Second");
        Assert.True(first >= 0 && second > first);
        Assert.Contains(string.Empty, texts.Skip(first + 1).Take(second - first - 1));
    }

    private static MemoryStream CreatePagesPackageWithUnlabeledList() {
        const ulong documentId = 1;
        const ulong bodyId = 2;
        const ulong listStyleId = 3;
        byte[] listTable = Message(BytesField(1,
            Message(VarintField(1, 0), ReferenceField(2, listStyleId))));
        byte[] records = Message(
            ArchiveRecord(documentId, 10000, Message(ReferenceField(4, bodyId)), new[] { bodyId }),
            ArchiveRecord(bodyId, 2001,
                Message(StringField(3, "Item"), BytesField(7, listTable)),
                new[] { listStyleId }),
            ArchiveRecord(listStyleId, 2023, Message(VarintField(11, 1))));
        return CreatePackage(("Index/Document.iwa", FrameIwa(records)));
    }

    private static MemoryStream CreateKeynotePackageWithLinkedNotes() {
        const ulong documentId = 1;
        const ulong showId = 2;
        const ulong nodeId = 3;
        const ulong slideId = 4;
        const ulong noteId = 5;
        const ulong storageId = 6;
        const ulong hyperlinkId = 7;
        byte[] slideTree = Message(ReferenceField(2, nodeId));
        byte[] hyperlinkTable = Message(BytesField(1,
            Message(VarintField(1, 0), ReferenceField(2, hyperlinkId))));
        byte[] records = Message(
            ArchiveRecord(documentId, 1, Message(ReferenceField(2, showId)), new[] { showId }),
            ArchiveRecord(showId, 2, Message(BytesField(3, slideTree)), new[] { nodeId }),
            ArchiveRecord(nodeId, 4, Message(ReferenceField(2, slideId)), new[] { slideId }),
            ArchiveRecord(slideId, 5, Message(ReferenceField(27, noteId)), new[] { noteId }),
            ArchiveRecord(noteId, 100, Message(ReferenceField(1, storageId)), new[] { storageId }),
            ArchiveRecord(storageId, 2001,
                Message(StringField(3, "Linked note"), BytesField(11, hyperlinkTable)),
                new[] { hyperlinkId }),
            ArchiveRecord(hyperlinkId, 2032,
                Message(StringField(2, "https://example.com/keynote-note"))));
        return CreatePackage(("Index/Document.iwa", FrameIwa(records)));
    }
}
