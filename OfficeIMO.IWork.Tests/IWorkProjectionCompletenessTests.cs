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
    public void High_precision_decimal128_values_use_visual_fallback_instead_of_rounding() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Decimal128", 1, 1, 0d, decimal128HighBit: true)
        }, includePreview: true);

        using var result = ExcelDocument.LoadNumbersWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics,
            diagnostic => diagnostic.Code == "IWORK_TABLE_CELL_DECODE");
    }

    [Fact]
    public void Enforces_the_projected_sheet_budget_before_owner_materialization() {
        using MemoryStream package = CreateNumbersPackage(Array.Empty<TableSpec>(),
            sheetReferenceCount: 2);
        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            IWorkSourceDocument.Open(package, IWorkDocumentKind.Numbers,
                new IWorkReadOptions { MaximumProjectedSheets = 1 }));

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

    [Fact]
    public void Wrong_keynote_required_record_types_disable_editable_reconstruction() {
        foreach (MemoryStream package in new[] {
                     CreateKeynotePackageWithRepeatedSlides(1, showType: 9999),
                     CreateKeynotePackageWithRepeatedSlides(1, nodeType: 9999),
                     CreateKeynotePackageWithRepeatedSlides(1, slideType: 9999)
                 }) {
            using (package)
            using (var result = PowerPointPresentation.LoadKeynoteWithReport(package)) {
                Assert.True(result.IsVisualFallback);
                Assert.False(result.Projection.HasEditableContent);
            }
        }
    }

    [Fact]
    public void Xml_invalid_text_runs_disable_all_editable_owner_projections() {
        using MemoryStream pagesPackage = CreatePagesPackage(includeBody: true, textBox: null,
            includePreview: true, bodyBytes: new byte[] { 0x01 });
        using MemoryStream numbersPackage = CreateNumbersPackage(Array.Empty<TableSpec>(),
            includePreview: true, textBoxBytes: new byte[] { 0x01 });
        using MemoryStream keynotePackage = CreateKeynotePackageWithInvalidText(new byte[] { 0x01 });

        using var pages = WordDocument.LoadPagesWithReport(pagesPackage);
        using var numbers = ExcelDocument.LoadNumbersWithReport(numbersPackage);
        using var keynote = PowerPointPresentation.LoadKeynoteWithReport(keynotePackage);

        Assert.True(pages.IsVisualFallback);
        Assert.True(numbers.IsVisualFallback);
        Assert.True(keynote.IsVisualFallback);
    }

    [Fact]
    public void Keynote_owner_preserves_explicit_zero_sized_text_boxes() {
        using MemoryStream package = CreateKeynotePackageWithRepeatedSlides(1, rotation: 0f);

        using var result = PowerPointPresentation.LoadKeynoteWithReport(package);

        PowerPointTextBox textBox = Assert.Single(Assert.Single(result.Document.Slides).TextBoxes);
        Assert.Equal(0d, textBox.WidthPoints);
        Assert.Equal(0d, textBox.HeightPoints);
        using var bytes = new MemoryStream();
        result.Document.Save(bytes);
        bytes.Position = 0;
        using PowerPointPresentation reopened = PowerPointPresentation.Load(bytes);
        PowerPointTextBox persisted = Assert.Single(Assert.Single(reopened.Slides).TextBoxes);
        Assert.Equal(0d, persisted.WidthPoints);
        Assert.Equal(0d, persisted.HeightPoints);
    }

    private static MemoryStream CreateKeynotePackageWithRepeatedSlides(int referenceCount,
        float? rotation = null, float? fontSize = null, bool wrongWireSlideSize = false,
        uint showType = 2, uint nodeType = 4, uint slideType = 5,
        string text = "Title", string? slideName = null, string? listLabel = null,
        bool wrongWireSkippedFlag = false, byte[]? textBoxDrawable = null,
        float? slideWidth = null, float? slideHeight = null,
        bool naturalAlignment = false, bool duplicateDrawableInField = false,
        bool aliasDrawableAcrossFields = false, int? drawableReferenceCount = null,
        int unexpectedSlideTreeFieldCount = 0, float? spaceBefore = null,
        float? spaceAfter = null) {
        const ulong documentId = 1;
        const ulong showId = 2;
        const ulong nodeId = 3;
        const ulong slideId = 4;
        const ulong shapeId = 5;
        const ulong storageId = 6;
        byte[] slideTree = Message(Enumerable.Range(0, referenceCount)
            .Select(_ => ReferenceField(2, nodeId))
            .Concat(Enumerable.Range(0, unexpectedSlideTreeFieldCount)
                .Select(value => VarintField(1, checked((ulong)value))))
            .ToArray());
        var shapeFields = new List<byte[]> { ReferenceField(2, storageId) };
        if (rotation.HasValue) {
            byte[] geometry = Message(FloatField(4, rotation.Value));
            byte[] drawable = Message(BytesField(1, geometry));
            byte[] shape = Message(BytesField(1, drawable));
            shapeFields.Insert(0, BytesField(1, shape));
        }
        if (textBoxDrawable != null) {
            shapeFields.Insert(0, BytesField(1,
                Message(BytesField(1, textBoxDrawable))));
        }
        var storageFields = new List<byte[]> { StringField(3, text) };
        var extraRecords = new List<byte[]>();
        if (listLabel != null) {
            const ulong listStyleId = 8;
            byte[] listEntry = Message(VarintField(1, 0), ReferenceField(2, listStyleId));
            storageFields.Add(BytesField(7, Message(BytesField(1, listEntry))));
            extraRecords.Add(ArchiveRecord(listStyleId, 2023,
                Message(VarintField(11, 1), StringField(16, listLabel))));
        }
        if (fontSize.HasValue) {
            const ulong characterStyleId = 7;
            byte[] styleEntry = Message(VarintField(1, 0), ReferenceField(2, characterStyleId));
            storageFields.Add(BytesField(8, Message(BytesField(1, styleEntry))));
            extraRecords.Add(ArchiveRecord(characterStyleId, 2021,
                Message(BytesField(11, Message(FloatField(3, fontSize.Value))))));
        }
        if (naturalAlignment || spaceBefore.HasValue || spaceAfter.HasValue) {
            const ulong paragraphStyleId = 9;
            byte[] styleEntry = Message(VarintField(1, 0),
                ReferenceField(2, paragraphStyleId));
            storageFields.Add(BytesField(5, Message(BytesField(1, styleEntry))));
            var paragraphFields = new List<byte[]>();
            if (naturalAlignment) paragraphFields.Add(VarintField(1, 4));
            if (spaceBefore.HasValue) paragraphFields.Add(FloatField(21, spaceBefore.Value));
            if (spaceAfter.HasValue) paragraphFields.Add(FloatField(20, spaceAfter.Value));
            extraRecords.Add(ArchiveRecord(paragraphStyleId, 2022,
                Message(BytesField(12, Message(paragraphFields.ToArray())))));
        }
        byte[] slideSize = slideWidth.HasValue && slideHeight.HasValue
            ? BytesField(4, Message(FloatField(1, slideWidth.Value), FloatField(2, slideHeight.Value)))
            : Array.Empty<byte>();
        byte[] showPayload = wrongWireSlideSize
            ? Message(BytesField(3, slideTree), VarintField(4, 1))
            : Message(BytesField(3, slideTree), slideSize);
        var slideFields = drawableReferenceCount.HasValue
            ? Enumerable.Range(0, drawableReferenceCount.Value)
                .Select(_ => ReferenceField(5, shapeId)).ToList()
            : new List<byte[]> { ReferenceField(5, shapeId) };
        if (duplicateDrawableInField) slideFields.Add(ReferenceField(5, shapeId));
        if (aliasDrawableAcrossFields) slideFields.Add(ReferenceField(6, shapeId));
        if (slideName != null) slideFields.Add(StringField(10, slideName));
        var records = new List<byte[]> {
            ArchiveRecord(documentId, 1, Message(ReferenceField(2, showId))),
            ArchiveRecord(showId, showType, showPayload),
            ArchiveRecord(nodeId, nodeType, Message(ReferenceField(2, slideId),
                wrongWireSkippedFlag ? BytesField(4, new byte[] { 1 }) : Array.Empty<byte>())),
            ArchiveRecord(slideId, slideType, Message(slideFields.ToArray())),
            ArchiveRecord(shapeId, 2011, Message(shapeFields.ToArray())),
            ArchiveRecord(storageId, 2001, Message(storageFields.ToArray()))
        };
        records.AddRange(extraRecords);
        return CreatePackage(
            ("Index/Slide.iwa", FrameIwa(Message(records.ToArray()))),
            ("preview.png", ValidPreviewPng()));
    }

    private static MemoryStream CreateKeynotePackageWithInvalidText(byte[]? textBytes = null) {
        const ulong documentId = 1;
        const ulong showId = 2;
        const ulong nodeId = 3;
        const ulong slideId = 4;
        const ulong shapeId = 5;
        const ulong storageId = 6;
        byte[] slideTree = Message(ReferenceField(2, nodeId));
        byte[] records = Message(
            ArchiveRecord(documentId, 1, Message(ReferenceField(2, showId))),
            ArchiveRecord(showId, 2, Message(BytesField(3, slideTree))),
            ArchiveRecord(nodeId, 4, Message(ReferenceField(2, slideId))),
            ArchiveRecord(slideId, 5, Message(ReferenceField(5, shapeId))),
            ArchiveRecord(shapeId, 2011, Message(ReferenceField(2, storageId))),
            ArchiveRecord(storageId, 2001, Message(BytesField(3, textBytes ?? new byte[] { 0xc3, 0x28 }))));
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
            ArchiveRecord(sectionId, 10011, Message(ReferenceField(25, headerFooterId)), new[] { headerFooterId }),
            ArchiveRecord(headerFooterId, 10143, Message(ReferenceField(1, wrongStorageId)),
                new[] { wrongStorageId }),
            ArchiveRecord(wrongStorageId, 9999, Message(StringField(3, "Not a text storage"))));
        return CreatePackage(
            ("Index/Document.iwa", FrameIwa(records)),
            ("preview.png", ValidPreviewPng()));
    }
}
