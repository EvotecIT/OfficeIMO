using OfficeIMO.Excel;
using OfficeIMO.IWork;
using OfficeIMO.Internal;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Enforces_the_combined_decompressed_iwa_budget_across_entries() {
        byte[] first = ArchiveRecord(1, 1, new byte[48]);
        byte[] second = ArchiveRecord(2, 6000, new byte[48]);
        using MemoryStream package = CreatePackage(
            ("Index/Document.iwa", FrameIwa(first)),
            ("Index/Table.iwa", FrameIwa(second)));
        int perEntryLimit = Math.Max(first.Length, second.Length);
        var options = new IWorkReadOptions {
            MaximumDecompressedIwaBytes = perEntryLimit,
            MaximumSnappyChunkBytes = perEntryLimit,
            MaximumTotalDecompressedIwaBytes = first.Length + second.Length - 1
        };

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            IWorkSourceDocument.Open(package, IWorkDocumentKind.Numbers, options));

        Assert.Contains("source-wide limit", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Decodes_numbers_rows_that_use_wide_offsets() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Wide", 1, 1, 42d, wideOffsets: true)
        });

        IWorkTable table = Assert.Single(Assert.Single(
            IWorkSourceDocument.Open(package, IWorkDocumentKind.Numbers).ReadNumbers().Sheets).Tables);

        Assert.Equal(42d, Assert.IsType<double>(table.GetCell(1, 1)!.Value), 10);
    }

    [Fact]
    public void Preserves_cached_numbers_values_when_a_formula_marker_is_present() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Formula", 1, 1, 42d, hasFormula: true)
        });

        IWorkTableCell cell = Assert.Single(Assert.Single(Assert.Single(
            IWorkSourceDocument.Open(package, IWorkDocumentKind.Numbers).ReadNumbers().Sheets).Tables).Cells);

        Assert.Equal(IWorkCellKind.Formula, cell.Kind);
        Assert.Equal("=?", cell.Formula);
        Assert.Equal(42d, Assert.IsType<double>(cell.Value), 10);

        using MemoryStream ownerPackage = CreateNumbersPackage(new[] {
            new TableSpec("Formula", 1, 1, 42d, hasFormula: true)
        });
        using var result = ExcelDocument.LoadNumbersWithReport(ownerPackage);
        Assert.Equal(42d, result.Document.Sheets[0].CellAt(1, 1).GetValue<double>(), 10);
    }

    [Fact]
    public void Enforces_the_materialized_cell_budget_across_all_tables() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("First", 1, 1, 1d),
            new TableSpec("Second", 1, 1, 2d)
        });
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, IWorkDocumentKind.Numbers,
            new IWorkReadOptions { MaximumMaterializedCells = 1 });

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() => source.ReadNumbers());

        Assert.Contains("source-wide limit", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Pre_bnc_numbers_storage_is_preserved_but_not_claimed_as_editable() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Legacy", 1, 1, 0d, legacyStorage: true)
        }, includePreview: true);

        IWorkSourceDocument source = IWorkSourceDocument.Open(package, IWorkDocumentKind.Numbers);
        IWorkNumbersProjection projection = source.ReadNumbers();
        IWorkImportReport report = projection.CreateImportReport(
            IWorkProjectionKind.VisualFallback, source.PreferredRasterPreview);

        Assert.False(projection.HasEditableContent);
        Assert.Contains(projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_TABLE_LEGACY_CELL_STORAGE");
        Assert.Contains(report.UnsupportedRecords, record => record.MessageType == 6002);
    }

    [Fact]
    public void Unsupported_numbers_table_models_are_preserved_and_disable_editable_claims() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Unsupported", 1, 1, 0d, missingModel: true)
        }, includePreview: true);

        IWorkSourceDocument source = IWorkSourceDocument.Open(package, IWorkDocumentKind.Numbers);
        IWorkNumbersProjection projection = source.ReadNumbers();
        IWorkImportReport report = projection.CreateImportReport(
            IWorkProjectionKind.VisualFallback, source.PreferredRasterPreview);

        Assert.False(projection.HasEditableContent);
        Assert.Contains(projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_TABLE_MODEL_UNSUPPORTED");
        Assert.Contains(report.UnsupportedRecords, record => record.MessageType == 6000);
    }

    [Fact]
    public void Missing_numbers_tiles_are_preserved_and_disable_editable_claims() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Missing tile", 1, 1, 0d, missingTile: true)
        }, includePreview: true);

        IWorkSourceDocument source = IWorkSourceDocument.Open(package, IWorkDocumentKind.Numbers);
        IWorkNumbersProjection projection = source.ReadNumbers();
        IWorkImportReport report = projection.CreateImportReport(
            IWorkProjectionKind.VisualFallback, source.PreferredRasterPreview);

        Assert.False(projection.HasEditableContent);
        Assert.Contains(projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_TABLE_TILE_UNSUPPORTED");
        Assert.Contains(report.UnsupportedRecords, record => record.MessageType == 6001);
    }

    [Fact]
    public void Missing_keynote_slides_are_preserved_and_use_visual_fallback() {
        using MemoryStream package = CreateKeynotePackageWithMissingSlide();

        using var result = PowerPointPresentation.LoadKeynoteWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.False(result.Projection.HasEditableContent);
        Assert.Empty(result.Projection.Slides);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_KEYNOTE_SLIDE_MISSING");
        Assert.Contains(result.ImportReport.UnsupportedRecords, record => record.MessageType == 4);
    }

    [Fact]
    public void Missing_keynote_presenter_notes_disable_editable_reconstruction() {
        using MemoryStream package = CreateKeynotePackageWithMissingNotes();

        using var result = PowerPointPresentation.LoadKeynoteWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.False(result.Projection.HasEditableContent);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_KEYNOTE_NOTES_UNSUPPORTED");
        Assert.Contains(result.ImportReport.UnsupportedRecords, record => record.Identifier == 4);
    }

    [Fact]
    public void Visual_only_numbers_import_bypasses_semantic_table_limits() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Large", 100, 1, 42d)
        }, includePreview: true);
        var options = new IWorkReadOptions {
            ImportMode = IWorkImportMode.VisualOnly,
            MaximumTableRows = 1
        };

        using var result = ExcelDocument.LoadNumbersWithReport(package, options);

        Assert.True(result.IsVisualFallback);
        Assert.Empty(result.Projection.Sheets);
        Assert.Equal(0, result.ImportReport.ReconstructedItemCount);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_SEMANTIC_PROJECTION_SKIPPED");
    }

    [Fact]
    public void Counts_zip_directory_nodes_against_the_package_entry_limit() {
        using var package = new MemoryStream();
        using (var archive = new ZipArchive(package, ZipArchiveMode.Create, leaveOpen: true)) {
            archive.CreateEntry("Index/");
            archive.CreateEntry("Index/Subdirectory/");
            WriteEntry(archive, "Index/Document.iwa", FrameIwa(ArchiveRecord(1, 1, Array.Empty<byte>())));
        }
        package.Position = 0;

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            IWorkSourceDocument.Open(package, IWorkDocumentKind.Numbers,
                new IWorkReadOptions { MaximumEntryCount = 2 }));

        Assert.Contains("entry count", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Counts_directory_bundle_nodes_against_the_package_entry_limit() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-iwork-entry-limit-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(Path.Combine(directory, "Index", "Subdirectory"));
        try {
            File.WriteAllBytes(Path.Combine(directory, "Index", "Document.iwa"),
                FrameIwa(ArchiveRecord(1, 1, Array.Empty<byte>())));

            InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
                IWorkSourceDocument.Open(directory, IWorkDocumentKind.Numbers,
                    new IWorkReadOptions { MaximumEntryCount = 2 }));

            Assert.Contains("entry count", exception.Message, StringComparison.OrdinalIgnoreCase);
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void Directory_bundles_reject_fifo_entries_without_blocking() {
        if (OperatingSystem.IsWindows()) return;

        string directory = Path.Combine(Path.GetTempPath(), "officeimo-iwork-fifo-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(Path.Combine(directory, "Index"));
        try {
            File.WriteAllBytes(Path.Combine(directory, "Index", "Document.iwa"),
                FrameIwa(ArchiveRecord(1, 1, Array.Empty<byte>())));
            string fifo = Path.Combine(directory, "blocking.fifo");
            Assert.Equal(0, CreateFifo(fifo, 0x180));

            InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
                IWorkSourceDocument.Open(directory, IWorkDocumentKind.Numbers));

            Assert.Contains("regular file", exception.Message, StringComparison.OrdinalIgnoreCase);
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void Regular_file_handles_must_resolve_within_the_captured_bundle_root() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-iwork-physical-root-" + Guid.NewGuid().ToString("N"));
        string root = Path.Combine(directory, "root");
        string outside = Path.Combine(directory, "outside");
        Directory.CreateDirectory(root);
        Directory.CreateDirectory(outside);
        try {
            string outsideFile = Path.Combine(outside, "Document.iwa");
            File.WriteAllBytes(outsideFile, FrameIwa(ArchiveRecord(1, 1, Array.Empty<byte>())));

            InvalidDataException exception = Assert.Throws<InvalidDataException>(() => {
                using FileStream _ = OfficePathIdentity.OpenRegularFileForRead(
                    outsideFile, OfficePathIdentity.ResolvePhysicalPath(root), 81920);
            });

            Assert.Contains("outside the source directory", exception.Message, StringComparison.OrdinalIgnoreCase);
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void Splits_the_first_numbers_table_when_leading_text_would_overflow_xlsx_rows() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Full Height", 1_048_576, 1, 42d)
        }, textBox: "Leading text");
        using var result = ExcelDocument.LoadNumbersWithReport(package);

        Assert.Equal(2, result.Document.Sheets.Count);
        Assert.Equal("Leading text", result.Document.Sheets[0].CellAt(1, 1).GetValue<string>());
        Assert.Equal(42d, result.Document.Sheets[1].CellAt(1, 1).GetValue<double>(), 10);
        using var bytes = new MemoryStream();
        result.Document.Save(bytes);
        bytes.Position = 0;
        using ExcelDocument reopened = ExcelDocument.Load(bytes);
        Assert.Equal(2, reopened.Sheets.Count);
    }

    [Fact]
    public void Malformed_numbers_references_disable_editable_reconstruction() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Table", 1, 1, 42d)
        }, includePreview: true, includeMalformedDrawableReference: true);

        using var result = ExcelDocument.LoadNumbersWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.False(result.Projection.HasEditableContent);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_NUMBERS_DRAWABLE_UNSUPPORTED");
    }

    [Fact]
    public void Missing_pages_body_disables_editable_reconstruction_even_when_a_text_box_exists() {
        using MemoryStream package = CreatePagesPackage(includeBody: false, textBox: "Floating text", includePreview: true);

        using var result = WordDocument.LoadPagesWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.False(result.Projection.HasEditableContent);
        Assert.Equal("Floating text", Assert.Single(result.Projection.TextBoxes));
        Assert.Contains(result.Projection.Diagnostics, diagnostic => diagnostic.Code == "IWORK_PAGES_BODY_MISSING");
    }

    [Fact]
    public void Pages_owner_projects_text_boxes_with_source_geometry() {
        using MemoryStream package = CreatePagesPackage(includeBody: true, textBox: "Floating text",
            includePreview: false, textBoxDrawable: GeometryDrawable(36f, 72f, 216f, 108f));

        using var result = WordDocument.LoadPagesWithReport(package);

        Assert.False(result.IsVisualFallback);
        WordTextBox textBox = Assert.Single(result.Document.TextBoxes);
        Assert.Equal("Floating text", Assert.Single(textBox.Paragraphs).Text);
        Assert.Equal(36 * 12700, textBox.HorizontalPositionOffset!.Value);
        Assert.Equal(72 * 12700, textBox.VerticalPositionOffset!.Value);
        Assert.Equal(216L * 12700L, textBox.Width);
        Assert.Equal(108L * 12700L, textBox.Height);
        using var bytes = new MemoryStream();
        result.Document.Save(bytes);
        bytes.Position = 0;
        using WordDocument reopened = WordDocument.Load(bytes);
        WordTextBox persisted = Assert.Single(reopened.TextBoxes);
        Assert.Equal("Floating text", Assert.Single(persisted.Paragraphs).Text);
        Assert.Equal(36 * 12700, persisted.HorizontalPositionOffset!.Value);
        Assert.Equal(216L * 12700L, persisted.Width);
    }

    [Fact]
    public void Empty_pages_body_is_valid_editable_reconstruction() {
        using MemoryStream package = CreatePagesPackage(includeBody: true, textBox: null, includePreview: false,
            bodyText: string.Empty);

        using var result = WordDocument.LoadPagesWithReport(package);

        Assert.False(result.IsVisualFallback);
        Assert.True(result.Projection.HasEditableContent);
        Assert.Empty(result.Projection.Paragraphs);
        Assert.Equal(IWorkProjectionKind.EditableReconstruction, result.ImportReport.ProjectionKind);
        Assert.Equal(0, result.ImportReport.ReconstructedItemCount);
    }

    [Fact]
    public void Enforces_source_wide_projected_text_character_bounds_before_materialization() {
        using MemoryStream package = CreatePagesPackage(includeBody: true, textBox: null,
            includePreview: false, bodyText: "12345");
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, IWorkDocumentKind.Pages,
            new IWorkReadOptions { MaximumProjectedTextCharacters = 4 });

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() => source.ReadPages());

        Assert.Contains("character count", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Enforces_source_wide_text_character_bounds_for_numbers_text_shapes() {
        using MemoryStream package = CreateNumbersPackage(Array.Empty<TableSpec>(), textBox: "12345");
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, IWorkDocumentKind.Numbers,
            new IWorkReadOptions { MaximumProjectedTextCharacters = 4 });

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() => source.ReadNumbers());

        Assert.Contains("character count", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Enforces_bounded_iterative_text_style_inheritance() {
        using MemoryStream package = CreatePagesPackageWithStyleChain(depth: 3);
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, IWorkDocumentKind.Pages,
            new IWorkReadOptions { MaximumTextStyleInheritanceDepth = 2 });

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() => source.ReadPages());

        Assert.Contains("style inheritance", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Enforces_text_attribute_bounds_before_nested_entry_projection() {
        using MemoryStream package = CreatePagesPackageWithTwoCharacterStyles();
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, IWorkDocumentKind.Pages,
            new IWorkReadOptions { MaximumProjectedTextItems = 1 });

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() => source.ReadPages());

        Assert.Contains("attribute count", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Invalid_utf8_in_rich_style_fields_disables_editable_reconstruction() {
        using MemoryStream package = CreatePagesPackageWithStyleChain(depth: 1,
            invalidFontName: true, includePreview: true);

        using var result = WordDocument.LoadPagesWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.False(result.Projection.Body.IsComplete);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_PAGES_TEXT_UNSUPPORTED");
    }

    [Fact]
    public void Malformed_rich_text_colors_disable_editable_reconstruction() {
        using MemoryStream package = CreatePagesPackageWithStyleChain(depth: 1,
            malformedColor: true, includePreview: true);

        using var result = WordDocument.LoadPagesWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.False(result.Projection.Body.IsComplete);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_PAGES_TEXT_UNSUPPORTED");
    }

    [Fact]
    public void Wrong_wire_rich_text_flags_disable_editable_reconstruction() {
        using MemoryStream package = CreatePagesPackageWithStyleChain(depth: 1,
            wrongWireBold: true, includePreview: true);

        using var result = WordDocument.LoadPagesWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.False(result.Projection.Body.IsComplete);
    }

    [Fact]
    public void Unknown_paragraph_alignment_disables_editable_reconstruction() {
        using MemoryStream package = CreatePagesPackageWithStyleChain(depth: 1,
            invalidAlignment: true, includePreview: true);

        using var result = WordDocument.LoadPagesWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.False(result.Projection.Body.IsComplete);
    }

    [Fact]
    public void Preserves_each_adjacent_rich_text_style_boundary() {
        using MemoryStream package = CreatePagesPackageWithTwoCharacterStyles();

        IWorkPagesProjection projection = IWorkSourceDocument.Open(package, IWorkDocumentKind.Pages).ReadPages();
        IWorkTextParagraph paragraph = Assert.Single(projection.Body.Paragraphs);

        Assert.Collection(paragraph.Runs,
            run => {
                Assert.Equal("A", run.Text);
                Assert.False(run.Style.Bold);
            },
            run => {
                Assert.Equal("B", run.Text);
                Assert.True(run.Style.Bold);
            });
    }

    [Fact]
    public void Malformed_nested_drawable_geometry_disables_editable_reconstruction() {
        using MemoryStream package = CreatePagesPackage(includeBody: true, textBox: "Floating",
            includePreview: true, textBoxDrawable: Message(BytesField(1, new byte[] { 0x08, 0x80 })));

        using var result = WordDocument.LoadPagesWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_PAGES_DRAWABLE_UNSUPPORTED");
    }

    [Fact]
    public void Wrong_wire_drawable_geometry_disables_editable_reconstruction() {
        using MemoryStream package = CreatePagesPackage(includeBody: true, textBox: "Floating",
            includePreview: true, textBoxDrawable: Message(VarintField(1, 1)));

        using var result = WordDocument.LoadPagesWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_PAGES_DRAWABLE_UNSUPPORTED");
    }

    [Fact]
    public void Pages_owner_falls_back_before_oversized_destination_measurements() {
        using MemoryStream package = CreatePagesPackage(includeBody: true, textBox: null,
            includePreview: true, documentLayoutFields: PageLayoutFields(float.MaxValue));

        using var result = WordDocument.LoadPagesWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.Equal(IWorkProjectionKind.VisualFallback, result.ImportReport.ProjectionKind);
    }

    [Fact]
    public void Missing_pages_section_references_disable_editable_reconstruction() {
        using MemoryStream package = CreatePagesPackageWithMissingSection();

        using var result = WordDocument.LoadPagesWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.False(result.Projection.HasEditableContent);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_PAGES_SECTION_UNSUPPORTED");
    }

    [Fact]
    public void Pages_preserves_distinct_header_storages_in_section_order() {
        using MemoryStream package = CreatePagesPackageWithDuplicateHeaders();

        IWorkPagesProjection projection = IWorkSourceDocument.Open(package, IWorkDocumentKind.Pages).ReadPages();

        Assert.True(projection.HasEditableContent);
        IWorkPagesSection section = Assert.Single(projection.Sections);
        Assert.Equal(2, section.HeaderContents.Count);
        Assert.Equal(new[] { "Header", "Header" }, projection.Headers);
    }

    [Fact]
    public void Pages_owner_preserves_header_footer_association_across_sections() {
        using MemoryStream package = CreatePagesPackageWithTwoSections();

        using var result = WordDocument.LoadPagesWithReport(package);

        Assert.False(result.IsVisualFallback);
        Assert.Equal(2, result.Projection.Sections.Count);
        Assert.Equal(2, result.Document.Sections.Count);
        Assert.Contains(result.Document.Sections[0].Header.Default!.Paragraphs,
            paragraph => paragraph.Text == "First header");
        Assert.Contains(result.Document.Sections[1].Header.Default!.Paragraphs,
            paragraph => paragraph.Text == "Second header");
        Assert.Contains(result.Document.Sections[1].Footer.Default!.Paragraphs,
            paragraph => paragraph.Text == "Second footer");
        using var bytes = new MemoryStream();
        result.Document.Save(bytes);
        bytes.Position = 0;
        using WordDocument reopened = WordDocument.Load(bytes);
        Assert.Equal(2, reopened.Sections.Count);
        Assert.Contains(reopened.Sections[0].Header.Default!.Paragraphs,
            paragraph => paragraph.Text == "First header");
        Assert.Contains(reopened.Sections[1].Header.Default!.Paragraphs,
            paragraph => paragraph.Text == "Second header");
    }

    [Fact]
    public void Pages_layout_breaks_do_not_shift_later_section_headers_and_footers() {
        using MemoryStream package = CreatePagesPackageWithTwoSections(includeLayoutBreak: true);

        using var result = WordDocument.LoadPagesWithReport(package);

        Assert.False(result.IsVisualFallback);
        Assert.Equal(2, result.Projection.Sections.Count);
        Assert.Equal(3, result.Document.Sections.Count);
        Assert.Contains(result.Document.Sections[0].Header.Default!.Paragraphs,
            paragraph => paragraph.Text == "First header");
        Assert.Null(result.Document.Sections[1].Header.Default);
        Assert.Contains(result.Document.Sections[2].Header.Default!.Paragraphs,
            paragraph => paragraph.Text == "Second header");
        Assert.Contains(result.Document.Sections[2].Footer.Default!.Paragraphs,
            paragraph => paragraph.Text == "Second footer");
        using var bytes = new MemoryStream();
        result.Document.Save(bytes);
        bytes.Position = 0;
        using WordDocument reopened = WordDocument.Load(bytes);
        Assert.Equal(3, reopened.Sections.Count);
        Assert.Contains(reopened.Sections[2].Header.Default!.Paragraphs,
            paragraph => paragraph.Text == "Second header");
    }

    [Fact]
    public void Orphaned_pages_shapes_are_preserved_but_not_inserted_as_text_boxes() {
        const ulong documentId = 1;
        const ulong bodyId = 2;
        const ulong orphanShapeId = 3;
        const ulong orphanStorageId = 4;
        byte[] records = Message(
            ArchiveRecord(documentId, 10000, Message(ReferenceField(4, bodyId)), new[] { bodyId }),
            ArchiveRecord(bodyId, 2001, Message(StringField(3, "Body"))),
            ArchiveRecord(orphanShapeId, 2011, Message(ReferenceField(2, orphanStorageId)), new[] { orphanStorageId }),
            ArchiveRecord(orphanStorageId, 2001, Message(StringField(3, "Orphan"))));
        using MemoryStream package = CreatePackage(("Index/Document.iwa", FrameIwa(records)));

        IWorkSourceDocument source = IWorkSourceDocument.Open(package, IWorkDocumentKind.Pages);
        IWorkPagesProjection projection = source.ReadPages();
        IWorkImportReport report = projection.CreateImportReport(IWorkProjectionKind.EditableReconstruction);

        Assert.Empty(projection.TextBoxes);
        Assert.Contains(report.UnsupportedRecords, record => record.Identifier == orphanShapeId);
        Assert.Contains(report.UnsupportedRecords, record => record.Identifier == orphanStorageId);
    }

    [Fact]
    public void Wrong_numbers_sheet_record_types_disable_editable_reconstruction() {
        byte[] records = Message(
            ArchiveRecord(1, 1, Message(ReferenceField(1, 2))),
            ArchiveRecord(2, 6000, Message()));
        using MemoryStream package = CreatePackage(
            ("Index/Document.iwa", FrameIwa(records)),
            ("preview.png", ValidPreviewPng()));

        using var result = ExcelDocument.LoadNumbersWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.False(result.Projection.HasEditableContent);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_NUMBERS_SHEET_TYPE_UNSUPPORTED");
        Assert.Contains(result.ImportReport.UnsupportedRecords, record => record.Identifier == 2);
    }

    [Fact]
    public void Duplicate_primary_identifiers_are_rejected_before_projection() {
        byte[] records = Message(
            ArchiveRecord(1, 1, Message(ReferenceField(1, 2))),
            ArchiveRecord(2, 2, Message(StringField(1, "Sheet"))),
            ArchiveRecord(2, 9999, Message()));
        using MemoryStream package = CreatePackage(("Index/Document.iwa", FrameIwa(records)));

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            IWorkSourceDocument.Open(package, IWorkDocumentKind.Numbers));

        Assert.Contains("primary IWA record", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Every_partially_consumed_iwa_record_remains_in_the_loss_report() {
        using MemoryStream package = CreatePagesPackage(includeBody: true, textBox: "Reachable", includePreview: false);
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, IWorkDocumentKind.Pages);

        IWorkImportReport report = source.ReadPages().CreateImportReport(IWorkProjectionKind.EditableReconstruction);

        Assert.Equal(report.TotalRecordCount, report.UnsupportedRecordCount);
        Assert.Equal(source.Records.Count, report.UnsupportedRecords.Count);
        Assert.True(report.HasConversionLoss);
    }

    [Fact]
    public void Duplicate_numbers_tile_indexes_disable_editable_reconstruction_before_replay() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Duplicate", 1, 1, 1d, duplicateCell: true)
        }, includePreview: true);

        using var result = ExcelDocument.LoadNumbersWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.False(result.Projection.HasEditableContent);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_TABLE_DUPLICATE_TILE");
    }

    [Fact]
    public void Duplicate_numbers_physical_tiles_disable_editable_reconstruction_before_replay() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Duplicate physical tile", 257, 1, 1d, duplicateTileIdentity: true)
        }, includePreview: true);

        using var result = ExcelDocument.LoadNumbersWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.False(result.Projection.HasEditableContent);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_TABLE_DUPLICATE_TILE");
    }

    [Fact]
    public void Duplicate_numbers_rows_within_a_tile_disable_editable_reconstruction() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Duplicate row", 1, 1, 1d, duplicateTileRow: true)
        }, includePreview: true);

        using var result = ExcelDocument.LoadNumbersWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.False(result.Projection.HasEditableContent);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_TABLE_TILE_ROW_UNSUPPORTED");
    }

    [Fact]
    public void Duplicate_numbers_string_keys_disable_editable_reconstruction() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Duplicate strings", 1, 1, 0d, textValue: "Value", duplicateString: true)
        }, includePreview: true);

        using var result = ExcelDocument.LoadNumbersWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.False(result.Projection.HasEditableContent);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_TABLE_STRING_STORAGE_UNSUPPORTED");
    }

    [Fact]
    public void Public_source_and_projection_collections_are_immutable() {
        using MemoryStream package = CreatePagesPackage(includeBody: true, textBox: null, includePreview: false);
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, IWorkDocumentKind.Pages);
        IWorkPagesProjection projection = source.ReadPages();

        var records = Assert.IsAssignableFrom<IList<IWorkArchiveRecord>>(source.Records);
        var entries = Assert.IsAssignableFrom<IList<IWorkPackageEntry>>(source.Entries);
        var paragraphs = Assert.IsAssignableFrom<IList<string>>(projection.Paragraphs);

        Assert.True(records.IsReadOnly);
        Assert.True(entries.IsReadOnly);
        Assert.True(paragraphs.IsReadOnly);
        Assert.Throws<NotSupportedException>(() => records.Clear());
        Assert.Throws<NotSupportedException>(() => entries.Clear());
        Assert.Throws<NotSupportedException>(() => paragraphs.Clear());
    }

    [Fact]
    public void Resource_names_that_begin_with_slide_do_not_override_pages_detection() {
        using MemoryStream package = CreatePagesPackage(includeBody: true, textBox: null, includePreview: false,
            archivePath: "Index/SlideshowResource.iwa");

        IWorkSourceDocument source = IWorkSourceDocument.Open(package, IWorkDocumentKind.Pages);

        Assert.Equal(IWorkDocumentKind.Pages, source.Kind);
        Assert.Equal("Body", Assert.Single(source.ReadPages().Paragraphs));
    }

    [Fact]
    public void Numbers_durations_are_written_as_excel_day_fractions() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Duration", 1, 1, 3600d, duration: true)
        });

        using var result = ExcelDocument.LoadNumbersWithReport(package);

        Assert.Equal(1d / 24d, result.Document.Sheets[0].CellAt(1, 1).GetValue<double>(), 10);
    }

    [Fact]
    public void Numbers_text_beyond_the_xlsx_cell_limit_uses_visual_fallback_without_losing_source_text() {
        string longText = new('x', 32_768);
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Long text", 1, 1, 0d, textValue: longText)
        }, includePreview: true);

        using var result = ExcelDocument.LoadNumbersWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.True(result.Projection.HasEditableContent);
        Assert.Equal(longText, Assert.Single(Assert.Single(result.Projection.Sheets).Tables).GetCell(1, 1)!.Value);

        using MemoryStream editableOnlyPackage = CreateNumbersPackage(new[] {
            new TableSpec("Long text", 1, 1, 0d, textValue: longText)
        }, includePreview: true);
        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            ExcelDocument.LoadNumbersWithReport(editableOnlyPackage,
                new IWorkReadOptions { ImportMode = IWorkImportMode.EditableOnly }));
        Assert.Contains("32,767", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Numbers_visual_fallback_preserves_preview_aspect_ratio() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Table", 1, 1, 42d)
        }, includePreview: true, previewBytes: CreateSizedPreviewPng(2400, 1200));

        using var result = ExcelDocument.LoadNumbersWithReport(package,
            new IWorkReadOptions { ImportMode = IWorkImportMode.VisualOnly });
        ExcelImage image = Assert.Single(result.Document.Sheets[0].Images);

        Assert.Equal(1600, image.WidthPixels);
        Assert.Equal(800, image.HeightPixels);
    }

    [Fact]
    public void Structurally_complete_pdf_previews_are_preserved_as_full_document_assets() {
        using MemoryStream package = CreatePagesPackage(includeBody: true, textBox: null, includePreview: false,
            pdfPreviewBytes: CreateValidPdf());

        IWorkPreviewAsset preview = Assert.Single(
            IWorkSourceDocument.Open(package, IWorkDocumentKind.Pages).Previews);

        Assert.Equal("application/pdf", preview.MediaType);
        Assert.Equal(IWorkVisualCoverage.FullDocument, preview.Coverage);
    }

    [Fact]
    public void Ignores_resource_names_and_malformed_files_that_only_look_like_previews() {
        using FileStream input = File.OpenRead(Fixture("nim-iwork/simple.pages"));
        using var source = new ZipArchive(input, ZipArchiveMode.Read, leaveOpen: false);
        using var package = new MemoryStream();
        using (var target = new ZipArchive(package, ZipArchiveMode.Create, leaveOpen: true)) {
            foreach (ZipArchiveEntry entry in source.Entries.Where(candidate =>
                         !string.Equals(candidate.FullName, "preview.jpg", StringComparison.OrdinalIgnoreCase)
                         && !string.Equals(candidate.FullName, "preview-web.jpg", StringComparison.OrdinalIgnoreCase)
                         && !string.Equals(candidate.FullName, "preview.png", StringComparison.OrdinalIgnoreCase)
                         && !string.Equals(candidate.FullName, "preview.pdf", StringComparison.OrdinalIgnoreCase))) {
                ZipArchiveEntry copy = target.CreateEntry(entry.FullName);
                using Stream sourceStream = entry.Open();
                using Stream targetStream = copy.Open();
                sourceStream.CopyTo(targetStream);
            }
            byte[] validPreview = ReadEntry(source, "preview.jpg");
            WriteEntry(target, "Data/unrelated-preview.jpg", validPreview);
            WriteEntry(target, "preview.jpg", new byte[] { 0xff, 0xd8, 0xff });
            WriteEntry(target, "preview.png", CreateIncompletePngHeader(1200, 900));
            WriteEntry(target, "preview.pdf", System.Text.Encoding.ASCII.GetBytes("%PDF-"));
            WriteEntry(target, "preview-web.jpg", validPreview);
        }
        package.Position = 0;

        IWorkSourceDocument document = IWorkSourceDocument.Open(package, IWorkDocumentKind.Pages);

        Assert.DoesNotContain(document.Previews, preview => preview.Path.StartsWith("Data/", StringComparison.Ordinal));
        Assert.DoesNotContain(document.Previews, preview => preview.Path == "preview.jpg");
        Assert.DoesNotContain(document.Previews, preview => preview.Path == "preview.png");
        Assert.DoesNotContain(document.Previews, preview => preview.Path == "preview.pdf");
        Assert.Equal("preview-web.jpg", document.PreferredRasterPreview!.Path);
    }

    private static MemoryStream CreateKeynotePackageWithMissingSlide() {
        const ulong documentId = 1;
        const ulong showId = 2;
        const ulong nodeId = 3;
        const ulong missingSlideId = 4;
        byte[] slideTree = Message(ReferenceField(2, nodeId));
        byte[] records = Message(
            ArchiveRecord(documentId, 1, Message(ReferenceField(2, showId))),
            ArchiveRecord(showId, 2, Message(BytesField(3, slideTree))),
            ArchiveRecord(nodeId, 4, Message(ReferenceField(2, missingSlideId))));
        return CreatePackage(
            ("Index/Document.iwa", FrameIwa(records)),
            ("preview.png", ValidPreviewPng()));
    }

    private static MemoryStream CreateKeynotePackageWithMissingNotes() {
        const ulong documentId = 1;
        const ulong showId = 2;
        const ulong nodeId = 3;
        const ulong slideId = 4;
        const ulong missingNoteId = 5;
        byte[] slideTree = Message(ReferenceField(2, nodeId));
        byte[] records = Message(
            ArchiveRecord(documentId, 1, Message(ReferenceField(2, showId))),
            ArchiveRecord(showId, 2, Message(BytesField(3, slideTree))),
            ArchiveRecord(nodeId, 4, Message(ReferenceField(2, slideId))),
            ArchiveRecord(slideId, 5, Message(ReferenceField(27, missingNoteId))));
        return CreatePackage(
            ("Index/Slide.iwa", FrameIwa(records)),
            ("preview.png", ValidPreviewPng()));
    }

    private static MemoryStream CreatePagesPackageWithMissingSection() {
        const ulong documentId = 1;
        const ulong bodyId = 2;
        const ulong missingSectionId = 3;
        byte[] sectionTable = Message(BytesField(1, Message(ReferenceField(2, missingSectionId))));
        byte[] body = Message(StringField(3, "Body"), BytesField(17, sectionTable));
        byte[] records = Message(
            ArchiveRecord(documentId, 10000, Message(ReferenceField(4, bodyId)), new[] { bodyId }),
            ArchiveRecord(bodyId, 2001, body));
        return CreatePackage(
            ("Index/Document.iwa", FrameIwa(records)),
            ("preview.png", ValidPreviewPng()));
    }

    private static MemoryStream CreatePagesPackageWithDuplicateHeaders() {
        const ulong documentId = 1;
        const ulong bodyId = 2;
        const ulong sectionId = 3;
        const ulong headerFooterId = 4;
        const ulong firstHeaderId = 5;
        const ulong secondHeaderId = 6;
        byte[] sectionTable = Message(BytesField(1, Message(ReferenceField(2, sectionId))));
        byte[] body = Message(StringField(3, "Body"), BytesField(17, sectionTable));
        byte[] records = Message(
            ArchiveRecord(documentId, 10000, Message(ReferenceField(4, bodyId)), new[] { bodyId }),
            ArchiveRecord(bodyId, 2001, body, new[] { sectionId }),
            ArchiveRecord(sectionId, 10011, Message(ReferenceField(23, headerFooterId)), new[] { headerFooterId }),
            ArchiveRecord(headerFooterId, 10143,
                Message(ReferenceField(1, firstHeaderId), ReferenceField(1, secondHeaderId)),
                new[] { firstHeaderId, secondHeaderId }),
            ArchiveRecord(firstHeaderId, 2001, Message(StringField(3, "Header"))),
            ArchiveRecord(secondHeaderId, 2001, Message(StringField(3, "Header"))));
        return CreatePackage(("Index/Document.iwa", FrameIwa(records)));
    }

    private static MemoryStream CreatePagesPackageWithTwoSections(bool includeLayoutBreak = false,
        bool emptySecondSection = false) {
        const ulong documentId = 1;
        const ulong bodyId = 2;
        const ulong firstSectionId = 3;
        const ulong secondSectionId = 4;
        const ulong firstHeaderFooterId = 5;
        const ulong secondHeaderFooterId = 6;
        const ulong firstHeaderId = 7;
        const ulong secondHeaderId = 8;
        const ulong secondFooterId = 9;
        byte[] sectionTable = Message(
            BytesField(1, Message(ReferenceField(2, firstSectionId))),
            BytesField(1, Message(ReferenceField(2, secondSectionId))));
        string bodyText = includeLayoutBreak ? "First\u0005Layout\u0004Second" : "First\u0004Second";
        byte[] body = Message(StringField(3, bodyText), BytesField(17, sectionTable));
        byte[] records = Message(
            ArchiveRecord(documentId, 10000, Message(ReferenceField(4, bodyId)), new[] { bodyId }),
            ArchiveRecord(bodyId, 2001, body, new[] { firstSectionId, secondSectionId }),
            ArchiveRecord(firstSectionId, 10011, Message(ReferenceField(23, firstHeaderFooterId)),
                new[] { firstHeaderFooterId }),
            ArchiveRecord(secondSectionId, 10011, Message(ReferenceField(23, secondHeaderFooterId)),
                new[] { secondHeaderFooterId }),
            ArchiveRecord(firstHeaderFooterId, 10143, Message(ReferenceField(1, firstHeaderId)),
                new[] { firstHeaderId }),
            ArchiveRecord(secondHeaderFooterId, 10143, emptySecondSection
                    ? Message()
                    : Message(ReferenceField(1, secondHeaderId), ReferenceField(2, secondFooterId)),
                emptySecondSection ? Array.Empty<ulong>() : new[] { secondHeaderId, secondFooterId }),
            ArchiveRecord(firstHeaderId, 2001, Message(StringField(3, "First header"))),
            ArchiveRecord(secondHeaderId, 2001, Message(StringField(3, "Second header"))),
            ArchiveRecord(secondFooterId, 2001, Message(StringField(3, "Second footer"))));
        return CreatePackage(("Index/Document.iwa", FrameIwa(records)));
    }

    private static MemoryStream CreatePagesPackageWithSharedCharacterStyle() {
        const ulong documentId = 1;
        const ulong bodyId = 2;
        const ulong firstParagraphStyleId = 10;
        const ulong secondParagraphStyleId = 11;
        const ulong characterStyleId = 12;
        byte[] firstParagraph = Message(VarintField(1, 0), ReferenceField(2, firstParagraphStyleId));
        byte[] secondParagraph = Message(VarintField(1, 2), ReferenceField(2, secondParagraphStyleId));
        byte[] character = Message(VarintField(1, 0), ReferenceField(2, characterStyleId));
        byte[] body = Message(StringField(3, "A\nB"),
            BytesField(5, Message(BytesField(1, firstParagraph), BytesField(1, secondParagraph))),
            BytesField(8, Message(BytesField(1, character))));
        byte[] records = Message(
            ArchiveRecord(documentId, 10000, Message(ReferenceField(4, bodyId)), new[] { bodyId }),
            ArchiveRecord(bodyId, 2001, body,
                new[] { firstParagraphStyleId, secondParagraphStyleId, characterStyleId }),
            ArchiveRecord(firstParagraphStyleId, 2022,
                Message(BytesField(11, Message(FloatField(3, 10f))))),
            ArchiveRecord(secondParagraphStyleId, 2022,
                Message(BytesField(11, Message(FloatField(3, 20f))))),
            ArchiveRecord(characterStyleId, 2021,
                Message(BytesField(11, Message(VarintField(1, 1))))));
        return CreatePackage(("Index/Document.iwa", FrameIwa(records)));
    }

    private static MemoryStream CreatePagesPackageWithStyleChain(int depth,
        bool invalidFontName = false, bool malformedColor = false, bool wrongWireBold = false,
        bool invalidAlignment = false, bool includePreview = false,
        bool naturalAlignment = false, string bodyText = "Styled", bool bold = false) {
        const ulong documentId = 1;
        const ulong bodyId = 2;
        const ulong firstStyleId = 10;
        byte[] styleEntry = Message(VarintField(1, 0), ReferenceField(2, firstStyleId));
        byte[] body = Message(StringField(3, bodyText),
            BytesField(5, Message(BytesField(1, styleEntry))));
        var records = new List<byte[]> {
            ArchiveRecord(documentId, 10000, Message(ReferenceField(4, bodyId)), new[] { bodyId }),
            ArchiveRecord(bodyId, 2001, body, new[] { firstStyleId })
        };
        for (int index = 0; index < depth; index++) {
            ulong identifier = firstStyleId + (ulong)index;
            ulong? parent = index + 1 < depth ? identifier + 1 : null;
            var fields = new List<byte[]>();
            if (parent.HasValue) fields.Add(BytesField(1, Message(ReferenceField(3, parent.Value))));
            if (index == 0 && invalidFontName) {
                fields.Add(BytesField(11, Message(BytesField(5, new byte[] { 0xc3, 0x28 }))));
            }
            if (index == 0 && malformedColor) {
                fields.Add(BytesField(11, Message(BytesField(7, new byte[] { 0x08, 0x80 }))));
            }
            if (index == 0 && wrongWireBold) {
                fields.Add(BytesField(11, Message(FloatField(1, 1f))));
            } else if (index == 0 && bold) {
                fields.Add(BytesField(11, Message(VarintField(1, 1))));
            }
            if (index == 0 && invalidAlignment) {
                fields.Add(BytesField(12, Message(VarintField(1, 99))));
            } else if (index == 0 && naturalAlignment) {
                fields.Add(BytesField(12, Message(VarintField(1, 4))));
            }
            records.Add(ArchiveRecord(identifier, 2022, Message(fields.ToArray()),
                parent.HasValue ? new[] { parent.Value } : Array.Empty<ulong>()));
        }
        byte[] iwaStream = Message(records.ToArray());
        return includePreview
            ? CreatePackage(("Index/Document.iwa", FrameIwa(iwaStream)), ("preview.png", ValidPreviewPng()))
            : CreatePackage(("Index/Document.iwa", FrameIwa(iwaStream)));
    }

    private static MemoryStream CreatePagesPackageWithTwoCharacterStyles() {
        const ulong documentId = 1;
        const ulong bodyId = 2;
        const ulong firstStyleId = 10;
        const ulong secondStyleId = 11;
        byte[] characterTable = Message(
            BytesField(1, Message(VarintField(1, 0), ReferenceField(2, firstStyleId))),
            BytesField(1, Message(VarintField(1, 1), ReferenceField(2, secondStyleId))));
        byte[] records = Message(
            ArchiveRecord(documentId, 10000, Message(ReferenceField(4, bodyId)), new[] { bodyId }),
            ArchiveRecord(bodyId, 2001,
                Message(StringField(3, "AB"), BytesField(8, characterTable)),
                new[] { firstStyleId, secondStyleId }),
            ArchiveRecord(firstStyleId, 2021,
                Message(BytesField(11, Message(VarintField(1, 0))))),
            ArchiveRecord(secondStyleId, 2021,
                Message(BytesField(11, Message(VarintField(1, 1))))));
        return CreatePackage(("Index/Document.iwa", FrameIwa(records)));
    }

    private static MemoryStream CreateNumbersPackage(IReadOnlyList<TableSpec> tables, string? textBox = null,
        bool includePreview = false, bool includeMalformedDrawableReference = false, byte[]? previewBytes = null,
        byte[]? textBoxBytes = null, int sheetReferenceCount = 1, bool duplicateFirstDrawable = false,
        byte[]? sheetNameBytes = null, bool includeWrongWireDrawableReference = false) {
        const ulong documentId = 1;
        const ulong sheetId = 2;
        var records = new List<byte[]>();
        var sheetFields = new List<byte[]> {
            sheetNameBytes == null ? StringField(1, "Sheet") : BytesField(1, sheetNameBytes)
        };
        byte[][] documentFields = Enumerable.Range(0, sheetReferenceCount)
            .Select(_ => ReferenceField(1, sheetId))
            .ToArray();
        records.Add(ArchiveRecord(documentId, 1, Message(documentFields)));

        for (int index = 0; index < tables.Count; index++) {
            TableSpec table = tables[index];
            ulong tableInfoId = checked((ulong)(10 + index * 4));
            ulong modelId = tableInfoId + 1;
            ulong tileId = tableInfoId + 2;
            ulong stringListId = tableInfoId + 3;
            sheetFields.Add(ReferenceField(2, tableInfoId));
            records.Add(ArchiveRecord(tableInfoId, 6000,
                table.MissingModel ? Message() : Message(ReferenceField(2, modelId))));
            if (table.MissingModel) continue;

            byte[] rowInfo = table.LegacyStorage
                ? Message(VarintField(1, 0), BytesField(3, new byte[] { 1 }), BytesField(4, new byte[] { 1 }))
                : CreateBncRow(table);
            byte[] tilePayload = table.DuplicateTileRow
                ? Message(BytesField(5, rowInfo), BytesField(5, rowInfo))
                : Message(BytesField(5, rowInfo));
            if (!table.MissingTile) records.Add(ArchiveRecord(tileId, 6002, tilePayload));

            byte[] tileEntry = Message(VarintField(1, 0), ReferenceField(2, tileId));
            byte[] tileStorage;
            if (table.DuplicateCell) {
                ulong duplicateTileId = tileId + 100_000;
                records.Add(ArchiveRecord(duplicateTileId, 6002, tilePayload));
                byte[] duplicateEntry = Message(VarintField(1, 0), ReferenceField(2, duplicateTileId));
                tileStorage = Message(BytesField(1, tileEntry), BytesField(1, duplicateEntry));
            } else if (table.DuplicateTileIdentity) {
                byte[] duplicateEntry = Message(VarintField(1, 1), ReferenceField(2, tileId));
                tileStorage = Message(BytesField(1, tileEntry), BytesField(1, duplicateEntry));
            } else {
                tileStorage = Message(BytesField(1, tileEntry));
            }
            ulong formulaListId = tableInfoId + 200_000;
            var storeFields = new List<byte[]> { BytesField(3, tileStorage) };
            if (table.TextValue != null) storeFields.Add(ReferenceField(4, stringListId));
            if (table.DuplicateFormula) storeFields.Add(ReferenceField(6, formulaListId));
            byte[] store = Message(storeFields.ToArray());
            var modelFields = new List<byte[]> {
                BytesField(4, store),
                table.WrongWireDimensions
                    ? BytesField(6, new byte[] { checked((byte)table.Rows) })
                    : VarintField(6, checked((ulong)table.Rows)),
                VarintField(7, checked((ulong)table.Columns)),
                StringField(8, table.Name)
            };
            if (table.DefaultColumnWidth.HasValue) {
                modelFields.Add(DoubleField(17, table.DefaultColumnWidth.Value));
            }
            if (table.DefaultRowHeight.HasValue) {
                modelFields.Add(DoubleField(16, table.DefaultRowHeight.Value));
            }
            if (table.HeaderRows > 0) modelFields.Add(VarintField(9, checked((ulong)table.HeaderRows)));
            if (table.FooterRows > 0) modelFields.Add(VarintField(11, checked((ulong)table.FooterRows)));
            byte[] model = Message(modelFields.ToArray());
            records.Add(ArchiveRecord(modelId, 6001, model));
            if (table.TextValue != null) {
                byte[] stringEntry = Message(VarintField(1, 1), StringField(3, table.TextValue));
                byte[] stringPayload = table.DuplicateString
                    ? Message(BytesField(3, stringEntry), BytesField(3, stringEntry))
                    : Message(BytesField(3, stringEntry));
                records.Add(ArchiveRecord(stringListId, 6200, stringPayload));
            }
            if (table.DuplicateFormula) {
                byte[] firstFormula = FormulaConstant(1d);
                byte[] secondFormula = FormulaConstant(2d);
                byte[] firstEntry = Message(VarintField(1, 0), BytesField(5, firstFormula));
                byte[] secondEntry = Message(VarintField(1, 0), BytesField(5, secondFormula));
                records.Add(ArchiveRecord(formulaListId, 6201,
                    Message(BytesField(3, firstEntry), BytesField(3, secondEntry))));
            }
        }

        if (textBox != null || textBoxBytes != null) {
            const ulong shapeId = 1000;
            const ulong storageId = 1001;
            sheetFields.Add(ReferenceField(2, shapeId));
            records.Add(ArchiveRecord(shapeId, 2011, Message(ReferenceField(2, storageId))));
            records.Add(ArchiveRecord(storageId, 2001, Message(
                textBoxBytes != null ? BytesField(3, textBoxBytes) : StringField(3, textBox!))));
        }
        if (includeMalformedDrawableReference) {
            sheetFields.Add(BytesField(2, new byte[] { 0x08, 0x80 }));
        }
        if (includeWrongWireDrawableReference) sheetFields.Add(VarintField(2, 10));
        if (duplicateFirstDrawable && tables.Count > 0) sheetFields.Add(ReferenceField(2, 10));

        records.Insert(1, ArchiveRecord(sheetId, 2, Message(sheetFields.ToArray())));
        byte[] iwaStream = Message(records.ToArray());
        return includePreview
            ? CreatePackage(
                ("Index/Document.iwa", FrameIwa(iwaStream)),
                ("preview.png", previewBytes ?? ValidPreviewPng()))
            : CreatePackage(("Index/Document.iwa", FrameIwa(iwaStream)));
    }

    private static byte[] CreateBncRow(TableSpec table) {
        int cellOffset = table.WideOffsets ? 4 : 0;
        int valueBytes = table.FormulaWithoutCachedValue ? 0 : table.Decimal128HighBit ? 16 : 8;
        var buffer = new byte[cellOffset + 12 + valueBytes + (table.HasFormula ? 4 : 0)];
        buffer[cellOffset] = 5;
        buffer[cellOffset + 1] = table.FormulaWithoutCachedValue ? (byte)9
            : table.TextValue != null ? (byte)3
            : table.Date ? (byte)5
            : table.Duration ? (byte)7
            : (byte)2;
        uint valueFlag = table.FormulaWithoutCachedValue ? 0
            : table.TextValue != null ? 1u << 3
            : table.Decimal128HighBit ? 1u
            : table.Date ? 1u << 2
            : 1u << 1;
        WriteUInt32(buffer, cellOffset + 8, valueFlag | (table.HasFormula ? 1u << 9 : 0));
        if (!table.FormulaWithoutCachedValue) {
            if (table.TextValue != null) WriteUInt32(buffer, cellOffset + 12, 1);
            else if (table.Decimal128HighBit) {
                buffer[cellOffset + 26] = 0x41;
                buffer[cellOffset + 27] = 0x30;
            } else {
                Buffer.BlockCopy(BitConverter.GetBytes(table.Value), 0, buffer, cellOffset + 12, 8);
            }
        }
        ushort encodedOffset = checked((ushort)(table.WideOffsets ? cellOffset / 4 : cellOffset));
        byte[] offsets = table.OddCurrentOffsets
            ? new[] { (byte)encodedOffset }
            : new[] { (byte)encodedOffset, (byte)(encodedOffset >> 8) };
        var fields = new List<byte[]> { VarintField(1, 0), BytesField(6, buffer) };
        if (!table.OmitCurrentOffsets) fields.Add(BytesField(7, offsets));
        if (table.WideOffsets) fields.Add(VarintField(8, 1));
        return Message(fields.ToArray());
    }

    private static byte[] FormulaConstant(double value) => Message(BytesField(1,
        Message(BytesField(1, Message(VarintField(1, 17), DoubleField(4, value))))));

    private static MemoryStream CreatePagesPackage(bool includeBody, string? textBox, bool includePreview,
        string archivePath = "Index/Document.iwa", byte[]? pdfPreviewBytes = null, string bodyText = "Body",
        byte[]? bodyBytes = null, byte[]? textBoxDrawable = null, byte[]? documentLayoutFields = null,
        uint textBoxStorageType = 2001) {
        const ulong documentId = 1;
        const ulong bodyId = 2;
        const ulong shapeId = 3;
        const ulong shapeStorageId = 4;
        var documentReferences = new List<ulong>();
        if (includeBody) documentReferences.Add(bodyId);
        if (textBox != null) documentReferences.Add(shapeId);
        var records = new List<byte[]> {
            ArchiveRecord(documentId, 10000,
                Message(includeBody ? ReferenceField(4, bodyId) : Array.Empty<byte>(),
                    documentLayoutFields ?? Array.Empty<byte>()), documentReferences)
        };
        if (includeBody) records.Add(ArchiveRecord(bodyId, 2001, Message(
            bodyBytes != null ? BytesField(3, bodyBytes) : StringField(3, bodyText))));
        if (textBox != null) {
            byte[] shape = textBoxDrawable == null
                ? Message(ReferenceField(2, shapeStorageId))
                : Message(BytesField(1, Message(BytesField(1, textBoxDrawable))),
                    ReferenceField(2, shapeStorageId));
            records.Add(ArchiveRecord(shapeId, 2011, shape,
                new[] { shapeStorageId }));
            records.Add(ArchiveRecord(shapeStorageId, textBoxStorageType, Message(StringField(3, textBox))));
        }
        byte[] iwaStream = Message(records.ToArray());
        var entries = new List<(string Path, byte[] Bytes)> { (archivePath, FrameIwa(iwaStream)) };
        if (includePreview) entries.Add(("preview.png", ValidPreviewPng()));
        if (pdfPreviewBytes != null) entries.Add(("preview.pdf", pdfPreviewBytes));
        return CreatePackage(entries.ToArray());
    }

    private static byte[] ArchiveRecord(ulong identifier, uint type, byte[] payload,
        IReadOnlyList<ulong>? objectReferences = null) {
        byte[][] referenceFields = (objectReferences ?? Array.Empty<ulong>())
            .Select(reference => VarintField(5, reference))
            .ToArray();
        byte[] messageInfo = Message(new[] {
            VarintField(1, type),
            VarintField(3, checked((ulong)payload.Length))
        }.Concat(referenceFields).ToArray());
        byte[] archiveInfo = Message(VarintField(1, identifier), BytesField(2, messageInfo));
        return Message(Varint(checked((ulong)archiveInfo.Length)), archiveInfo, payload);
    }

    private static byte[] ArchiveRecordWithRepeatedSingularFields(ulong identifier, uint type,
        byte[] payload) {
        byte[] messageInfo = Message(
            VarintField(1, 9999), VarintField(1, type),
            VarintField(3, 0), VarintField(3, checked((ulong)payload.Length)));
        byte[] archiveInfo = Message(
            VarintField(1, 999), VarintField(1, identifier), BytesField(2, messageInfo));
        return Message(Varint(checked((ulong)archiveInfo.Length)), archiveInfo, payload);
    }

    private static byte[] FrameIwa(byte[] uncompressed) {
        var block = new List<byte>(uncompressed.Length + 10);
        block.AddRange(Varint(checked((ulong)uncompressed.Length)));
        int encodedLength = checked(uncompressed.Length - 1);
        if (uncompressed.Length <= 60) {
            block.Add(checked((byte)(encodedLength << 2)));
        } else {
            int lengthBytes = encodedLength <= byte.MaxValue ? 1
                : encodedLength <= ushort.MaxValue ? 2
                : encodedLength <= 0x00ff_ffff ? 3 : 4;
            block.Add(checked((byte)((59 + lengthBytes) << 2)));
            for (int index = 0; index < lengthBytes; index++) block.Add((byte)(encodedLength >> (index * 8)));
        }
        block.AddRange(uncompressed);
        int chunkLength = block.Count;
        var result = new List<byte>(chunkLength + 4) {
            0,
            (byte)chunkLength,
            (byte)(chunkLength >> 8),
            (byte)(chunkLength >> 16)
        };
        result.AddRange(block);
        return result.ToArray();
    }

    private static byte[] ReferenceField(int field, ulong identifier) =>
        BytesField(field, Message(VarintField(1, identifier)));

    private static byte[] StringField(int field, string value) =>
        BytesField(field, System.Text.Encoding.UTF8.GetBytes(value));

    private static byte[] VarintField(int field, ulong value) =>
        Message(Varint(checked((ulong)(field << 3))), Varint(value));

    private static byte[] GeometryDrawable(float left, float top, float width, float height) =>
        Message(BytesField(1, Message(
            BytesField(1, Message(FloatField(1, left), FloatField(2, top))),
            BytesField(2, Message(FloatField(1, width), FloatField(2, height))))));

    private static byte[] PageLayoutFields(float topMargin) => Message(
        FloatField(30, 612f), FloatField(31, 792f), FloatField(32, 72f),
        FloatField(33, 72f), FloatField(34, topMargin), FloatField(35, 72f),
        FloatField(36, 36f), FloatField(37, 36f));

    private static byte[] FloatField(int field, float value) =>
        Message(Varint(checked((ulong)((field << 3) | 5))), BitConverter.GetBytes(value));

    private static byte[] DoubleField(int field, double value) =>
        Message(Varint(checked((ulong)((field << 3) | 1))), BitConverter.GetBytes(value));

    private static byte[] BytesField(int field, byte[] value) =>
        Message(Varint(checked((ulong)((field << 3) | 2))), Varint(checked((ulong)value.Length)), value);

    private static byte[] Varint(ulong value) {
        var result = new List<byte>(10);
        do {
            byte next = (byte)(value & 0x7f);
            value >>= 7;
            if (value != 0) next |= 0x80;
            result.Add(next);
        } while (value != 0);
        return result.ToArray();
    }

    private static byte[] Message(params byte[][] parts) {
        int length = parts.Sum(part => part.Length);
        var result = new byte[length];
        int offset = 0;
        foreach (byte[] part in parts) {
            Buffer.BlockCopy(part, 0, result, offset, part.Length);
            offset += part.Length;
        }
        return result;
    }

    private static MemoryStream CreatePackage(params (string Path, byte[] Bytes)[] entries) {
        var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true)) {
            foreach ((string path, byte[] bytes) in entries) WriteEntry(archive, path, bytes);
        }
        stream.Position = 0;
        return stream;
    }

    private static void WriteEntry(ZipArchive archive, string path, byte[] bytes) {
        ZipArchiveEntry entry = archive.CreateEntry(path);
        using Stream target = entry.Open();
        target.Write(bytes, 0, bytes.Length);
    }

    private static byte[] ReadEntry(ZipArchive archive, string path) {
        ZipArchiveEntry entry = archive.GetEntry(path)!;
        using Stream source = entry.Open();
        using var bytes = new MemoryStream();
        source.CopyTo(bytes);
        return bytes.ToArray();
    }

    private static void WriteUInt32(byte[] bytes, int offset, uint value) {
        bytes[offset] = (byte)value;
        bytes[offset + 1] = (byte)(value >> 8);
        bytes[offset + 2] = (byte)(value >> 16);
        bytes[offset + 3] = (byte)(value >> 24);
    }

    private static byte[] CreateIncompletePngHeader(int width, int height) {
        var bytes = new byte[24] {
            0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a,
            0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
            0, 0, 0, 0, 0, 0, 0, 0
        };
        WriteBigEndian32(bytes, 16, width);
        WriteBigEndian32(bytes, 20, height);
        return bytes;
    }

    private static void WriteBigEndian32(byte[] bytes, int offset, int value) {
        bytes[offset] = (byte)(value >> 24);
        bytes[offset + 1] = (byte)(value >> 16);
        bytes[offset + 2] = (byte)(value >> 8);
        bytes[offset + 3] = (byte)value;
    }

    private static byte[] ValidPreviewPng() => Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");

    private static byte[] CreateSizedPreviewPng(int width, int height) {
        var header = new byte[13];
        WriteBigEndian32(header, 0, width);
        WriteBigEndian32(header, 4, height);
        header[8] = 8;
        header[9] = 0;
        using var imageData = new MemoryStream();
        imageData.WriteByte(0x78);
        imageData.WriteByte(0x9c);
        using (var deflate = new DeflateStream(imageData, CompressionMode.Compress, leaveOpen: true)) {
            var row = new byte[checked(width + 1)];
            for (int index = 0; index < height; index++) deflate.Write(row, 0, row.Length);
        }
        long decodedLength = checked((long)(width + 1) * height);
        uint adler = (uint)(decodedLength % 65521) << 16 | 1u;
        var checksum = new byte[4];
        WriteBigEndian32(checksum, 0, unchecked((int)adler));
        imageData.Write(checksum, 0, checksum.Length);
        byte[] signature = { 0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a };
        return Message(signature, CreatePngChunk("IHDR", header),
            CreatePngChunk("IDAT", imageData.ToArray()),
            CreatePngChunk("IEND", Array.Empty<byte>()));
    }

    private static byte[] CreateCrcValidPngWithInvalidImageData() {
        byte[] valid = ValidPreviewPng();
        byte[] signatureAndHeader = valid.Take(33).ToArray();
        return Message(signatureAndHeader, CreatePngChunk("IDAT", new byte[] { 0 }),
            CreatePngChunk("IEND", Array.Empty<byte>()));
    }

    private static byte[] CreatePngChunk(string type, byte[] data) {
        var chunk = new byte[12 + data.Length];
        WriteBigEndian32(chunk, 0, data.Length);
        byte[] typeBytes = System.Text.Encoding.ASCII.GetBytes(type);
        Buffer.BlockCopy(typeBytes, 0, chunk, 4, 4);
        Buffer.BlockCopy(data, 0, chunk, 8, data.Length);
        WriteBigEndian32(chunk, 8 + data.Length,
            unchecked((int)CalculatePngCrc(chunk, 4, 4 + data.Length)));
        return chunk;
    }

    private static uint CalculatePngCrc(byte[] bytes, int offset, int length) {
        uint crc = uint.MaxValue;
        for (int index = offset; index < offset + length; index++) {
            crc ^= bytes[index];
            for (int bit = 0; bit < 8; bit++) {
                crc = (crc & 1) != 0 ? 0xedb88320U ^ crc >> 1 : crc >> 1;
            }
        }
        return crc ^ uint.MaxValue;
    }

    private static byte[] CreateValidPdf() {
        const string header = "%PDF-1.4\n";
        const string catalog = "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n";
        const string pages = "2 0 obj\n<< /Type /Pages /Count 0 /Kids [] >>\nendobj\n";
        int catalogOffset = System.Text.Encoding.ASCII.GetByteCount(header);
        int pagesOffset = System.Text.Encoding.ASCII.GetByteCount(header + catalog);
        string prefix = header + catalog + pages;
        int xrefOffset = System.Text.Encoding.ASCII.GetByteCount(prefix);
        string suffix = "xref\n0 3\n0000000000 65535 f \n"
            + catalogOffset.ToString("D10", System.Globalization.CultureInfo.InvariantCulture) + " 00000 n \n"
            + pagesOffset.ToString("D10", System.Globalization.CultureInfo.InvariantCulture) + " 00000 n \n"
            + "trailer\n<< /Size 3 /Root 1 0 R >>\nstartxref\n"
            + xrefOffset.ToString(System.Globalization.CultureInfo.InvariantCulture)
            + "\n%%EOF\n";
        return System.Text.Encoding.ASCII.GetBytes(prefix + suffix);
    }

    private static string Fixture(string relativePath) =>
        Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus",
            relativePath.Replace('/', Path.DirectorySeparatorChar));

    [System.Runtime.InteropServices.DllImport("libc", EntryPoint = "mkfifo", SetLastError = true,
        CharSet = System.Runtime.InteropServices.CharSet.Ansi)]
    private static extern int CreateFifo(string path, uint mode);

    private sealed class TableSpec {
        internal TableSpec(string name, int rows, int columns, double value,
            bool wideOffsets = false, bool legacyStorage = false, bool hasFormula = false,
            bool missingModel = false, bool missingTile = false, string? textValue = null,
            bool duration = false, bool duplicateCell = false, bool duplicateString = false,
            bool decimal128HighBit = false, bool duplicateTileIdentity = false,
            bool duplicateTileRow = false, bool omitCurrentOffsets = false, bool date = false,
            bool formulaWithoutCachedValue = false, double? defaultColumnWidth = null,
            bool duplicateFormula = false, bool wrongWireDimensions = false,
            bool oddCurrentOffsets = false, double? defaultRowHeight = null,
            int headerRows = 0, int footerRows = 0) {
            Name = name;
            Rows = rows;
            Columns = columns;
            Value = value;
            WideOffsets = wideOffsets;
            LegacyStorage = legacyStorage;
            HasFormula = hasFormula;
            MissingModel = missingModel;
            MissingTile = missingTile;
            TextValue = textValue;
            Duration = duration;
            DuplicateCell = duplicateCell;
            DuplicateString = duplicateString;
            Decimal128HighBit = decimal128HighBit;
            DuplicateTileIdentity = duplicateTileIdentity;
            DuplicateTileRow = duplicateTileRow;
            OmitCurrentOffsets = omitCurrentOffsets;
            Date = date;
            FormulaWithoutCachedValue = formulaWithoutCachedValue;
            DefaultColumnWidth = defaultColumnWidth;
            DuplicateFormula = duplicateFormula;
            WrongWireDimensions = wrongWireDimensions;
            OddCurrentOffsets = oddCurrentOffsets;
            DefaultRowHeight = defaultRowHeight;
            HeaderRows = headerRows;
            FooterRows = footerRows;
        }

        internal string Name { get; }
        internal int Rows { get; }
        internal int Columns { get; }
        internal double Value { get; }
        internal bool WideOffsets { get; }
        internal bool LegacyStorage { get; }
        internal bool HasFormula { get; }
        internal bool MissingModel { get; }
        internal bool MissingTile { get; }
        internal string? TextValue { get; }
        internal bool Duration { get; }
        internal bool DuplicateCell { get; }
        internal bool DuplicateString { get; }
        internal bool Decimal128HighBit { get; }
        internal bool DuplicateTileIdentity { get; }
        internal bool DuplicateTileRow { get; }
        internal bool OmitCurrentOffsets { get; }
        internal bool Date { get; }
        internal bool FormulaWithoutCachedValue { get; }
        internal double? DefaultColumnWidth { get; }
        internal bool DuplicateFormula { get; }
        internal bool WrongWireDimensions { get; }
        internal bool OddCurrentOffsets { get; }
        internal double? DefaultRowHeight { get; }
        internal int HeaderRows { get; }
        internal int FooterRows { get; }
    }
}
