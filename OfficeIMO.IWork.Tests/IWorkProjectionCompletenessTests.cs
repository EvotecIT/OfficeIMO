using OfficeIMO.Excel;
using OfficeIMO.IWork;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Cached_formula_durations_keep_their_excel_day_fraction_scale() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Formula duration", 1, 1, 3600d, hasFormula: true, duration: true)
        });

        using var result = ExcelDocument.LoadNumbersWithReport(package);
        IWorkTableCell cell = Assert.Single(Assert.Single(Assert.Single(result.Projection.Sheets).Tables).Cells);

        Assert.Equal(IWorkCellKind.Formula, cell.Kind);
        Assert.Equal(IWorkCellKind.Duration, cell.ValueKind);
        Assert.Equal(1d / 24d, result.Document.Sheets[0].CellAt(1, 1).GetValue<double>(), 10);
    }

    [Fact]
    public void Decodes_the_decimal128_high_coefficient_bit_at_weight_two_to_112() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Decimal128", 1, 1, 0d, decimal128HighBit: true)
        });

        IWorkTableCell cell = Assert.Single(Assert.Single(Assert.Single(
            IWorkSourceDocument.Open(package, IWorkDocumentKind.Numbers).ReadNumbers().Sheets).Tables).Cells);

        Assert.Equal(Math.Pow(2d, 112), Assert.IsType<double>(cell.Value), 12);
    }

    [Fact]
    public void Enforces_the_projected_sheet_budget_before_owner_materialization() {
        using MemoryStream package = CreateNumbersPackage(Array.Empty<TableSpec>(),
            sheetReferenceCount: 2);
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, IWorkDocumentKind.Numbers,
            new IWorkReadOptions { MaximumProjectedSheets = 1 });

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() => source.ReadNumbers());

        Assert.Contains("sheet count", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Enforces_the_projected_table_budget_before_owner_materialization() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("First", 1, 1, 1d),
            new TableSpec("Second", 1, 1, 2d)
        });
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, IWorkDocumentKind.Numbers,
            new IWorkReadOptions { MaximumProjectedTables = 1 });

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() => source.ReadNumbers());

        Assert.Contains("table count", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Enforces_the_projected_text_item_budget_before_owner_materialization() {
        using MemoryStream package = CreatePagesPackage(includeBody: true, textBox: null, includePreview: false,
            bodyText: "First\nSecond");
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, IWorkDocumentKind.Pages,
            new IWorkReadOptions { MaximumProjectedTextItems = 1 });

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() => source.ReadPages());

        Assert.Contains("text item count", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Duplicate_numbers_sheet_references_disable_editable_reconstruction() {
        using MemoryStream package = CreateNumbersPackage(Array.Empty<TableSpec>(), includePreview: true,
            sheetReferenceCount: 2);

        using var result = ExcelDocument.LoadNumbersWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_NUMBERS_DUPLICATE_SHEET");
    }

    [Fact]
    public void Duplicate_numbers_drawable_references_disable_editable_reconstruction() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Table", 1, 1, 1d)
        }, includePreview: true, duplicateFirstDrawable: true);

        using var result = ExcelDocument.LoadNumbersWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_NUMBERS_DUPLICATE_DRAWABLE");
    }

    [Fact]
    public void Enforces_the_projected_slide_budget_before_owner_materialization() {
        using MemoryStream package = CreateKeynotePackageWithRepeatedSlides(2);
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, IWorkDocumentKind.Keynote,
            new IWorkReadOptions { MaximumProjectedSlides = 1 });

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() => source.ReadKeynote());

        Assert.Contains("slide count", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Duplicate_keynote_slide_references_disable_editable_reconstruction() {
        using MemoryStream package = CreateKeynotePackageWithRepeatedSlides(2);

        using var result = PowerPointPresentation.LoadKeynoteWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_KEYNOTE_DUPLICATE_SLIDE");
    }

    [Fact]
    public void Malformed_declared_pages_section_tables_disable_editable_reconstruction() {
        using MemoryStream package = CreatePagesPackageWithMalformedSectionTable();

        using var result = WordDocument.LoadPagesWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_PAGES_SECTION_UNSUPPORTED");
    }

    [Fact]
    public void Invalid_pages_text_runs_disable_editable_reconstruction() {
        using MemoryStream package = CreatePagesPackage(includeBody: true, textBox: null, includePreview: true,
            bodyBytes: new byte[] { 0xc3, 0x28 });

        using var result = WordDocument.LoadPagesWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_PAGES_TEXT_UNSUPPORTED");
    }

    [Fact]
    public void Invalid_numbers_text_runs_disable_editable_reconstruction() {
        using MemoryStream package = CreateNumbersPackage(Array.Empty<TableSpec>(), includePreview: true,
            textBoxBytes: new byte[] { 0xc3, 0x28 });

        using var result = ExcelDocument.LoadNumbersWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_NUMBERS_TEXT_STORAGE_UNSUPPORTED");
    }

    [Fact]
    public void Invalid_keynote_text_runs_disable_editable_reconstruction() {
        using MemoryStream package = CreateKeynotePackageWithInvalidText();

        using var result = PowerPointPresentation.LoadKeynoteWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_KEYNOTE_TEXT_STORAGE_UNSUPPORTED");
    }

    [Fact]
    public void Wrong_pages_header_storage_types_disable_editable_reconstruction() {
        using MemoryStream package = CreatePagesPackageWithWrongHeaderStorage();

        using var result = WordDocument.LoadPagesWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_PAGES_HEADER_FOOTER_UNSUPPORTED");
    }

    [Fact]
    public void Pdf_previews_without_xref_entries_trailer_and_catalog_are_rejected() {
        byte[] malformed = System.Text.Encoding.ASCII.GetBytes(
            "%PDF-1.4\nxref\nstartxref\n9\n%%EOF\n");
        using MemoryStream package = CreatePagesPackage(includeBody: true, textBox: null, includePreview: true,
            pdfPreviewBytes: malformed);

        IWorkSourceDocument source = IWorkSourceDocument.Open(package, IWorkDocumentKind.Pages);

        Assert.DoesNotContain(source.Previews, preview => preview.MediaType == "application/pdf");
        Assert.Equal("preview.png", source.PreferredRasterPreview!.Path);
    }

    [Fact]
    public void Pdf_previews_without_a_catalog_pages_tree_are_rejected() {
        string malformed = System.Text.Encoding.ASCII.GetString(CreateValidPdf())
            .Replace("/Pages 2 0 R", "/Pagez 2 0 R");
        using MemoryStream package = CreatePagesPackage(includeBody: true, textBox: null, includePreview: true,
            pdfPreviewBytes: System.Text.Encoding.ASCII.GetBytes(malformed));

        IWorkSourceDocument source = IWorkSourceDocument.Open(package, IWorkDocumentKind.Pages);

        Assert.DoesNotContain(source.Previews, preview => preview.MediaType == "application/pdf");
        Assert.Equal("preview.png", source.PreferredRasterPreview!.Path);
    }

    [Fact]
    public void Pages_visual_fallback_fits_inside_the_word_section_content_area() {
        using MemoryStream package = CreatePagesPackage(includeBody: false, textBox: null, includePreview: true);
        using var result = WordDocument.LoadPagesWithReport(package);

        WordSection section = result.Document.Sections[0];
        double contentWidth = ((long)(section.PageSettings.Width ?? WordPageSizes.Letter.WidthTwips)
            - section.Margins.Left - section.Margins.Right) / 20d;
        double contentHeight = ((long)(section.PageSettings.Height ?? WordPageSizes.Letter.HeightTwips)
            - section.Margins.Top.GetValueOrDefault() - section.Margins.Bottom.GetValueOrDefault()) / 20d;
        WordImage image = Assert.Single(result.Document.Images);

        Assert.InRange(image.Width.GetValueOrDefault(), 1d, contentWidth);
        Assert.InRange(image.Height.GetValueOrDefault(), 1d, contentHeight);
    }

    [Fact]
    public void Keynote_rotation_outside_the_pptx_range_uses_visual_fallback() {
        using MemoryStream package = CreateKeynotePackageWithRepeatedSlides(1, rotation: float.MaxValue);

        using var result = PowerPointPresentation.LoadKeynoteWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.True(result.Projection.HasEditableContent);
    }

    [Fact]
    public void Keynote_font_size_outside_the_pptx_range_uses_visual_fallback() {
        using MemoryStream package = CreateKeynotePackageWithRepeatedSlides(1, fontSize: 5000f);

        using var result = PowerPointPresentation.LoadKeynoteWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.True(result.Projection.HasEditableContent);
    }

    [Fact]
    public void Wrong_wire_keynote_slide_size_disables_editable_reconstruction() {
        using MemoryStream package = CreateKeynotePackageWithRepeatedSlides(1, wrongWireSlideSize: true);

        using var result = PowerPointPresentation.LoadKeynoteWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_KEYNOTE_SLIDE_SIZE_UNSUPPORTED");
    }

    private static MemoryStream CreateKeynotePackageWithRepeatedSlides(int referenceCount,
        float? rotation = null, float? fontSize = null, bool wrongWireSlideSize = false) {
        const ulong documentId = 1;
        const ulong showId = 2;
        const ulong nodeId = 3;
        const ulong slideId = 4;
        const ulong shapeId = 5;
        const ulong storageId = 6;
        byte[] slideTree = Message(Enumerable.Range(0, referenceCount)
            .Select(_ => ReferenceField(2, nodeId))
            .ToArray());
        var shapeFields = new List<byte[]> { ReferenceField(2, storageId) };
        if (rotation.HasValue) {
            byte[] geometry = Message(FloatField(4, rotation.Value));
            byte[] drawable = Message(BytesField(1, geometry));
            byte[] shape = Message(BytesField(1, drawable));
            shapeFields.Insert(0, BytesField(1, shape));
        }
        var storageFields = new List<byte[]> { StringField(3, "Title") };
        var extraRecords = new List<byte[]>();
        if (fontSize.HasValue) {
            const ulong characterStyleId = 7;
            byte[] styleEntry = Message(VarintField(1, 0), ReferenceField(2, characterStyleId));
            storageFields.Add(BytesField(8, Message(BytesField(1, styleEntry))));
            extraRecords.Add(ArchiveRecord(characterStyleId, 2021,
                Message(BytesField(11, Message(FloatField(3, fontSize.Value))))));
        }
        byte[] showPayload = wrongWireSlideSize
            ? Message(BytesField(3, slideTree), VarintField(4, 1))
            : Message(BytesField(3, slideTree));
        var records = new List<byte[]> {
            ArchiveRecord(documentId, 1, Message(ReferenceField(2, showId))),
            ArchiveRecord(showId, 7000, showPayload),
            ArchiveRecord(nodeId, 7001, Message(ReferenceField(2, slideId))),
            ArchiveRecord(slideId, 7002, Message(ReferenceField(5, shapeId))),
            ArchiveRecord(shapeId, 2011, Message(shapeFields.ToArray())),
            ArchiveRecord(storageId, 2001, Message(storageFields.ToArray()))
        };
        records.AddRange(extraRecords);
        return CreatePackage(
            ("Index/Slide.iwa", FrameIwa(Message(records.ToArray()))),
            ("preview.png", ValidPreviewPng()));
    }

    private static MemoryStream CreateKeynotePackageWithInvalidText() {
        const ulong documentId = 1;
        const ulong showId = 2;
        const ulong nodeId = 3;
        const ulong slideId = 4;
        const ulong shapeId = 5;
        const ulong storageId = 6;
        byte[] slideTree = Message(ReferenceField(2, nodeId));
        byte[] records = Message(
            ArchiveRecord(documentId, 1, Message(ReferenceField(2, showId))),
            ArchiveRecord(showId, 7000, Message(BytesField(3, slideTree))),
            ArchiveRecord(nodeId, 7001, Message(ReferenceField(2, slideId))),
            ArchiveRecord(slideId, 7002, Message(ReferenceField(5, shapeId))),
            ArchiveRecord(shapeId, 2011, Message(ReferenceField(2, storageId))),
            ArchiveRecord(storageId, 2001, Message(BytesField(3, new byte[] { 0xc3, 0x28 }))));
        return CreatePackage(
            ("Index/Slide.iwa", FrameIwa(records)),
            ("preview.png", ValidPreviewPng()));
    }

    private static MemoryStream CreatePagesPackageWithMalformedSectionTable() {
        const ulong documentId = 1;
        const ulong bodyId = 2;
        byte[] body = Message(StringField(3, "Body"), BytesField(17, new byte[] { 0x08, 0x80 }));
        byte[] records = Message(
            ArchiveRecord(documentId, 10000, Message(ReferenceField(4, bodyId)), new[] { bodyId }),
            ArchiveRecord(bodyId, 2001, body));
        return CreatePackage(
            ("Index/Document.iwa", FrameIwa(records)),
            ("preview.png", ValidPreviewPng()));
    }

    private static MemoryStream CreatePagesPackageWithWrongHeaderStorage() {
        const ulong documentId = 1;
        const ulong bodyId = 2;
        const ulong sectionId = 3;
        const ulong headerFooterId = 4;
        const ulong wrongStorageId = 5;
        byte[] sectionTable = Message(BytesField(1, Message(ReferenceField(2, sectionId))));
        byte[] body = Message(StringField(3, "Body"), BytesField(17, sectionTable));
        byte[] records = Message(
            ArchiveRecord(documentId, 10000, Message(ReferenceField(4, bodyId)), new[] { bodyId }),
            ArchiveRecord(bodyId, 2001, body, new[] { sectionId }),
            ArchiveRecord(sectionId, 10011, Message(ReferenceField(23, headerFooterId)), new[] { headerFooterId }),
            ArchiveRecord(headerFooterId, 10143, Message(ReferenceField(1, wrongStorageId)),
                new[] { wrongStorageId }),
            ArchiveRecord(wrongStorageId, 9999, Message(StringField(3, "Not a text storage"))));
        return CreatePackage(
            ("Index/Document.iwa", FrameIwa(records)),
            ("preview.png", ValidPreviewPng()));
    }
}
