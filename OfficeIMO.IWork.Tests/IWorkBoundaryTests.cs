using OfficeIMO.Excel;
using OfficeIMO.IWork;
using OfficeIMO.Internal;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

namespace OfficeIMO.IWork.Tests;

public sealed class IWorkBoundaryTests {
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

        IWorkNumbersTable table = Assert.Single(Assert.Single(
            IWorkSourceDocument.Open(package, IWorkDocumentKind.Numbers).ReadNumbers().Sheets).Tables);

        Assert.Equal(42d, Assert.IsType<double>(table.GetCell(1, 1)!.Value), 10);
    }

    [Fact]
    public void Preserves_cached_numbers_values_when_a_formula_marker_is_present() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Formula", 1, 1, 42d, hasFormula: true)
        });

        IWorkNumbersCell cell = Assert.Single(Assert.Single(Assert.Single(
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
            diagnostic.Code == "IWORK_NUMBERS_LEGACY_CELL_STORAGE");
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
            diagnostic.Code == "IWORK_NUMBERS_TABLE_MODEL_UNSUPPORTED");
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
            diagnostic.Code == "IWORK_NUMBERS_TILE_UNSUPPORTED");
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
        Assert.Contains(result.ImportReport.UnsupportedRecords, record => record.MessageType == 7001);
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
    public void Missing_pages_section_references_disable_editable_reconstruction() {
        using MemoryStream package = CreatePagesPackageWithMissingSection();

        using var result = WordDocument.LoadPagesWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.False(result.Projection.HasEditableContent);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_PAGES_SECTION_UNSUPPORTED");
    }

    [Fact]
    public void Pages_headers_are_deduplicated_in_source_order() {
        using MemoryStream package = CreatePagesPackageWithDuplicateHeaders();

        IWorkPagesProjection projection = IWorkSourceDocument.Open(package, IWorkDocumentKind.Pages).ReadPages();

        Assert.True(projection.HasEditableContent);
        Assert.Equal("Header", Assert.Single(projection.Headers));
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
    public void Duplicate_primary_identifiers_remain_in_the_loss_report() {
        byte[] records = Message(
            ArchiveRecord(1, 1, Message(ReferenceField(1, 2))),
            ArchiveRecord(2, 2, Message(StringField(1, "Sheet"))),
            ArchiveRecord(2, 9999, Message()));
        using MemoryStream package = CreatePackage(("Index/Document.iwa", FrameIwa(records)));
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, IWorkDocumentKind.Numbers);
        IWorkNumbersProjection projection = source.ReadNumbers();

        IWorkImportReport report = projection.CreateImportReport(IWorkProjectionKind.EditableReconstruction);

        Assert.True(projection.HasEditableContent);
        Assert.Equal(3, report.TotalRecordCount);
        Assert.Contains(report.UnsupportedRecords, record =>
            record.Identifier == 2 && record.MessageType == 9999);
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
    public void Duplicate_numbers_cell_coordinates_disable_editable_reconstruction() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Duplicate", 1, 1, 1d, duplicateCell: true)
        }, includePreview: true);

        using var result = ExcelDocument.LoadNumbersWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.False(result.Projection.HasEditableContent);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_NUMBERS_DUPLICATE_CELL");
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
            diagnostic.Code == "IWORK_NUMBERS_STRING_STORAGE_UNSUPPORTED");
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
            ArchiveRecord(showId, 7000, Message(BytesField(3, slideTree))),
            ArchiveRecord(nodeId, 7001, Message(ReferenceField(2, missingSlideId))));
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
            ArchiveRecord(showId, 7000, Message(BytesField(3, slideTree))),
            ArchiveRecord(nodeId, 7001, Message(ReferenceField(2, slideId))),
            ArchiveRecord(slideId, 7002, Message(ReferenceField(27, missingNoteId))));
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

    private static MemoryStream CreateNumbersPackage(IReadOnlyList<TableSpec> tables, string? textBox = null,
        bool includePreview = false, bool includeMalformedDrawableReference = false, byte[]? previewBytes = null) {
        const ulong documentId = 1;
        const ulong sheetId = 2;
        var records = new List<byte[]>();
        var sheetFields = new List<byte[]> { StringField(1, "Sheet") };
        var documentFields = new[] { ReferenceField(1, sheetId) };
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
            byte[] tilePayload = Message(BytesField(5, rowInfo));
            if (!table.MissingTile) records.Add(ArchiveRecord(tileId, 6002, tilePayload));

            byte[] tileEntry = Message(VarintField(1, 0), ReferenceField(2, tileId));
            byte[] tileStorage;
            if (table.DuplicateCell) {
                ulong duplicateTileId = tileId + 100_000;
                records.Add(ArchiveRecord(duplicateTileId, 6002, tilePayload));
                byte[] duplicateEntry = Message(VarintField(1, 0), ReferenceField(2, duplicateTileId));
                tileStorage = Message(BytesField(1, tileEntry), BytesField(1, duplicateEntry));
            } else {
                tileStorage = Message(BytesField(1, tileEntry));
            }
            byte[] store = table.TextValue == null
                ? Message(BytesField(3, tileStorage))
                : Message(BytesField(3, tileStorage), ReferenceField(4, stringListId));
            byte[] model = Message(
                BytesField(4, store),
                VarintField(6, checked((ulong)table.Rows)),
                VarintField(7, checked((ulong)table.Columns)),
                StringField(8, table.Name));
            records.Add(ArchiveRecord(modelId, 6001, model));
            if (table.TextValue != null) {
                byte[] stringEntry = Message(VarintField(1, 1), StringField(3, table.TextValue));
                byte[] stringPayload = table.DuplicateString
                    ? Message(BytesField(3, stringEntry), BytesField(3, stringEntry))
                    : Message(BytesField(3, stringEntry));
                records.Add(ArchiveRecord(stringListId, 6200, stringPayload));
            }
        }

        if (textBox != null) {
            const ulong shapeId = 1000;
            const ulong storageId = 1001;
            sheetFields.Add(ReferenceField(2, shapeId));
            records.Add(ArchiveRecord(shapeId, 2011, Message(ReferenceField(2, storageId))));
            records.Add(ArchiveRecord(storageId, 2001, Message(StringField(3, textBox))));
        }
        if (includeMalformedDrawableReference) {
            sheetFields.Add(BytesField(2, new byte[] { 0x08, 0x80 }));
        }

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
        var buffer = new byte[cellOffset + (table.HasFormula ? 24 : 20)];
        buffer[cellOffset] = 5;
        buffer[cellOffset + 1] = table.TextValue != null ? (byte)3 : table.Duration ? (byte)7 : (byte)2;
        uint valueFlag = table.TextValue != null ? 1u << 3 : 1u << 1;
        WriteUInt32(buffer, cellOffset + 8, valueFlag | (table.HasFormula ? 1u << 9 : 0));
        if (table.TextValue != null) WriteUInt32(buffer, cellOffset + 12, 1);
        else Buffer.BlockCopy(BitConverter.GetBytes(table.Value), 0, buffer, cellOffset + 12, 8);
        ushort encodedOffset = checked((ushort)(table.WideOffsets ? cellOffset / 4 : cellOffset));
        byte[] offsets = { (byte)encodedOffset, (byte)(encodedOffset >> 8) };
        var fields = new List<byte[]> {
            VarintField(1, 0),
            BytesField(6, buffer),
            BytesField(7, offsets)
        };
        if (table.WideOffsets) fields.Add(VarintField(8, 1));
        return Message(fields.ToArray());
    }

    private static MemoryStream CreatePagesPackage(bool includeBody, string? textBox, bool includePreview,
        string archivePath = "Index/Document.iwa", byte[]? pdfPreviewBytes = null, string bodyText = "Body") {
        const ulong documentId = 1;
        const ulong bodyId = 2;
        const ulong shapeId = 3;
        const ulong shapeStorageId = 4;
        var documentReferences = new List<ulong>();
        if (includeBody) documentReferences.Add(bodyId);
        if (textBox != null) documentReferences.Add(shapeId);
        var records = new List<byte[]> {
            ArchiveRecord(documentId, 10000,
                includeBody ? Message(ReferenceField(4, bodyId)) : Message(), documentReferences)
        };
        if (includeBody) records.Add(ArchiveRecord(bodyId, 2001, Message(StringField(3, bodyText))));
        if (textBox != null) {
            records.Add(ArchiveRecord(shapeId, 2011, Message(ReferenceField(2, shapeStorageId)),
                new[] { shapeStorageId }));
            records.Add(ArchiveRecord(shapeStorageId, 2001, Message(StringField(3, textBox))));
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
        byte[] bytes = ValidPreviewPng();
        WriteBigEndian32(bytes, 16, width);
        WriteBigEndian32(bytes, 20, height);
        WriteBigEndian32(bytes, 29, unchecked((int)CalculatePngCrc(bytes, 12, 17)));
        return bytes;
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
        const string prefix = "%PDF-1.4\n1 0 obj\n<< /Type /Catalog >>\nendobj\n";
        int xrefOffset = System.Text.Encoding.ASCII.GetByteCount(prefix);
        string suffix = "xref\n0 2\n0000000000 65535 f \n0000000009 00000 n \n"
            + "trailer\n<< /Size 2 /Root 1 0 R >>\nstartxref\n"
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
            bool duration = false, bool duplicateCell = false, bool duplicateString = false) {
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
    }
}
