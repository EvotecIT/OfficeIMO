using OfficeIMO.Excel;
using OfficeIMO.IWork;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

namespace OfficeIMO.IWork.Tests;

public sealed class IWorkCorpusTests {
    public static IEnumerable<object[]> Corpus() {
        yield return new object[] { "nim-iwork/simple.pages", IWorkDocumentKind.Pages };
        yield return new object[] { "nim-iwork/simple.numbers", IWorkDocumentKind.Numbers };
        yield return new object[] { "nim-iwork/simple.key", IWorkDocumentKind.Keynote };
        yield return new object[] { "iwork-converter/a.pages", IWorkDocumentKind.Pages };
        yield return new object[] { "iwork-converter/a.numbers", IWorkDocumentKind.Numbers };
        yield return new object[] { "iwork-converter/a.key", IWorkDocumentKind.Keynote };
        yield return new object[] { "numbers-parser/issue-102-v15.1.numbers", IWorkDocumentKind.Numbers };
        yield return new object[] { "numbers-parser/test-9-merges.numbers", IWorkDocumentKind.Numbers };
        yield return new object[] { "numbers-parser/test-10-formulas.numbers", IWorkDocumentKind.Numbers };
        yield return new object[] { "picodocs/sample-v14.4.pages", IWorkDocumentKind.Pages };
        yield return new object[] { "keynotekit/tabledeck-v15.2.1.key", IWorkDocumentKind.Keynote };
        yield return new object[] { "keynotekit/imagedeck-v15.2.1.key", IWorkDocumentKind.Keynote };
    }

    [Theory]
    [MemberData(nameof(Corpus))]
    public void Reads_independently_produced_versions_with_path_stream_parity(string relativePath,
        IWorkDocumentKind expectedKind) {
        string path = Fixture(relativePath);
        IWorkSourceDocument fromPath = IWorkSourceDocument.Open(path);
        using FileStream stream = File.OpenRead(path);
        IWorkSourceDocument fromStream = IWorkSourceDocument.Open(stream, expectedKind);

        Assert.Equal(expectedKind, fromPath.Kind);
        Assert.Equal(expectedKind, fromStream.Kind);
        Assert.NotEmpty(fromPath.Entries);
        Assert.NotEmpty(fromPath.Records);
        Assert.Equal(fromPath.Records.Count, fromStream.Records.Count);
        Assert.Equal(fromPath.Entries.Select(entry => entry.Path), fromStream.Entries.Select(entry => entry.Path));
        Assert.NotEmpty(fromPath.BuildVersions);
    }

    [Fact]
    public void Reads_current_pages_text_and_preserves_unrecognized_records() {
        IWorkSourceDocument source = IWorkSourceDocument.Open(Fixture("nim-iwork/simple.pages"));
        IWorkPagesProjection pages = source.ReadPages();
        IWorkImportReport report = pages.CreateImportReport(IWorkProjectionKind.EditableReconstruction);

        Assert.Equal("hello pages", pages.Paragraphs[0]);
        Assert.Contains(pages.Paragraphs, paragraph => paragraph.Contains(
            "second paragraph with some words", StringComparison.Ordinal));
        Assert.True(report.TotalRecordCount >= report.UnsupportedRecords.Count);
        Assert.NotEmpty(report.UnsupportedRecords);
        Assert.True(report.HasConversionLoss);
    }

    [Fact]
    public void Reads_current_numbers_as_sparse_typed_cells() {
        IWorkNumbersProjection numbers = IWorkSourceDocument.Open(Fixture("nim-iwork/simple.numbers")).ReadNumbers();
        IWorkTable table = Assert.Single(Assert.Single(numbers.Sheets).Tables);

        Assert.Equal(3, table.RowCount);
        Assert.Equal(3, table.ColumnCount);
        Assert.Equal("a", table.GetCell(1, 1)!.Value);
        Assert.Equal("C", table.GetCell(1, 3)!.Value);
        Assert.Equal(2d, Assert.IsType<double>(table.GetCell(2, 2)!.Value), 10);
        Assert.Equal("Z", table.GetCell(3, 3)!.Value);
        Assert.Throws<ArgumentOutOfRangeException>(() => table.GetCell(2, 4));
    }

    [Fact]
    public void Numbers_recovers_complete_formulas_cached_values_and_table_metadata() {
        IWorkNumbersProjection numbers = IWorkSourceDocument.Open(
            Fixture("numbers-parser/test-10-formulas.numbers")).ReadNumbers();
        IWorkTable table = numbers.Sheets[0].Tables[0];

        Assert.True(numbers.HasEditableContent);
        Assert.InRange(table.DefaultRowHeight!.Value, 19.92d, 19.94d);
        Assert.Equal(98d, table.DefaultColumnWidth);
        IWorkTableCell arithmetic = table.GetCell(2, 2)!;
        Assert.Equal(IWorkCellKind.Formula, arithmetic.Kind);
        Assert.Equal("=A1+A2", arithmetic.Formula);
        Assert.True(arithmetic.FormulaIsComplete);
        Assert.Equal(3d, Assert.IsType<double>(arithmetic.Value), 10);
        Assert.Equal("=SUM(A1:A2)", table.GetCell(6, 2)!.Formula);
        Assert.Equal("=IF(A6>6,TRUE,FALSE)", table.GetCell(6, 3)!.Formula);
        Assert.Equal("=LEFT(A3,1)", numbers.Sheets[1].Tables[0].GetCell(3, 2)!.Formula);
    }

    [Fact]
    public void Numbers_recovers_merges_and_excel_projects_them_as_editable_ranges() {
        const string RelativePath = "numbers-parser/test-9-merges.numbers";
        IWorkNumbersProjection numbers = IWorkSourceDocument.Open(Fixture(RelativePath)).ReadNumbers();
        IWorkTable first = numbers.Sheets[0].Tables[0];

        Assert.True(numbers.HasEditableContent);
        Assert.Equal(5, first.MergedRanges.Count);
        Assert.Contains(first.MergedRanges, merge => merge.FirstRow == 2 && merge.FirstColumn == 1
            && merge.LastRow == 2 && merge.LastColumn == 2);
        Assert.Contains(first.MergedRanges, merge => merge.FirstRow == 7 && merge.FirstColumn == 4
            && merge.LastRow == 8 && merge.LastColumn == 5);

        using var result = ExcelDocument.LoadNumbersWithReport(Fixture(RelativePath));
        Assert.False(result.IsVisualFallback);
        Assert.Equal(5, result.Document.Sheets[0].GetMergedRanges().Count);
        Assert.Contains(result.Document.Sheets[0].GetMergedRanges(), merge => merge.A1Range == "A2:B2");
    }

    [Fact]
    public void Numbers_enforces_the_configured_merged_range_bound() {
        IWorkSourceDocument source = IWorkSourceDocument.Open(
            Fixture("numbers-parser/test-9-merges.numbers"),
            new IWorkReadOptions { MaximumTableMergedRanges = 1 });

        Assert.Throws<InvalidDataException>(() => source.ReadNumbers());
    }

    [Fact]
    public void Numbers_owner_projects_source_formulas_with_cached_values() {
        using var result = ExcelDocument.LoadNumbersWithReport(
            Fixture("numbers-parser/test-10-formulas.numbers"));

        Assert.False(result.IsVisualFallback);
        ExcelSheet first = result.Document.Sheets[0];
        Assert.Equal("A1+A2", first.GetFormulaText(2, 2));
        Assert.Equal("SUM(A1:A2)", first.GetFormulaText(6, 2));
        Assert.Equal(3d, first.CellAt(2, 2).GetValue<double>(), 10);
    }

    [Fact]
    public void Numbers_formula_projection_honors_node_and_character_bounds() {
        const string RelativePath = "numbers-parser/test-10-formulas.numbers";
        IWorkSourceDocument nodeBounded = IWorkSourceDocument.Open(Fixture(RelativePath),
            new IWorkReadOptions { MaximumFormulaNodes = 1 });
        Assert.Throws<InvalidDataException>(() => nodeBounded.ReadNumbers());

        IWorkNumbersProjection characterBounded = IWorkSourceDocument.Open(Fixture(RelativePath),
            new IWorkReadOptions { MaximumFormulaCharacters = 4 }).ReadNumbers();
        IWorkTableCell formula = characterBounded.Sheets[0].Tables[0].GetCell(2, 2)!;
        Assert.False(formula.FormulaIsComplete);
        Assert.Equal("=#FORMULA!", formula.Formula);
        Assert.Equal(3d, Assert.IsType<double>(formula.Value), 10);
    }

    [Fact]
    public void Reads_current_keynote_slides_and_notes() {
        IWorkKeynoteProjection keynote = IWorkSourceDocument.Open(Fixture("nim-iwork/simple.key")).ReadKeynote();

        Assert.Equal(2, keynote.Slides.Count);
        Assert.Equal("hello keynote", keynote.Slides[0].Title);
        Assert.Contains("first bullet", keynote.Slides[0].Body);
        Assert.Equal("second slide", keynote.Slides[1].Title);
        Assert.Contains("note text here", keynote.Slides[1].PresenterNotes, StringComparison.Ordinal);
    }

    [Fact]
    public void Older_and_newer_versions_recover_editable_structure() {
        IWorkPagesProjection pages = IWorkSourceDocument.Open(Fixture("iwork-converter/a.pages")).ReadPages();
        IWorkNumbersProjection olderNumbers = IWorkSourceDocument.Open(Fixture("iwork-converter/a.numbers")).ReadNumbers();
        IWorkKeynoteProjection keynote = IWorkSourceDocument.Open(Fixture("iwork-converter/a.key")).ReadKeynote();
        IWorkNumbersProjection newerNumbers = IWorkSourceDocument.Open(Fixture("numbers-parser/issue-102-v15.1.numbers")).ReadNumbers();

        Assert.True(pages.Paragraphs.Count >= 30);
        Assert.Contains(pages.Paragraphs, paragraph => paragraph.Contains("合同", StringComparison.Ordinal));
        Assert.Equal(2, Assert.Single(olderNumbers.Sheets).Tables.Count);
        Assert.Equal(253, olderNumbers.Sheets[0].Tables.Sum(table => table.Cells.Count));
        Assert.Single(keynote.Slides);
        Assert.NotEmpty(keynote.Slides[0].Body);
        IWorkTable newerTable = Assert.Single(Assert.Single(newerNumbers.Sheets).Tables);
        Assert.Equal(7, newerTable.RowCount);
        Assert.Equal(11, newerTable.ColumnCount);
        Assert.Equal("Cats", newerTable.GetCell(1, 3)!.Value);
    }

    [Fact]
    public void Pages_rich_text_extracts_typography_but_marks_inline_objects_incomplete() {
        IWorkPagesProjection pages = IWorkSourceDocument.Open(Fixture("iwork-converter/a.pages")).ReadPages();

        Assert.False(pages.Body.IsComplete);
        Assert.Equal(45, pages.Body.Paragraphs.Count);
        IWorkTextParagraph title = pages.Body.Paragraphs[0];
        IWorkTextRun titleRun = Assert.Single(title.Runs);
        Assert.Equal("购 销 合 同", title.Text);
        Assert.Equal(IWorkTextAlignment.Center, title.Style.Alignment);
        Assert.True(titleRun.Style.Bold);
        Assert.False(titleRun.Style.Italic);
        Assert.Equal(26d, titleRun.Style.FontSizePoints);
        Assert.Equal("SimSun", titleRun.Style.FontName);
        Assert.Equal("000000", titleRun.Style.Color!.RgbHex);
        Assert.Empty(pages.Body.Paragraphs[1].Runs);
    }

    [Fact]
    public void Pages_owner_uses_visual_fallback_for_unresolved_inline_objects() {
        using var result = WordDocument.LoadPagesWithReport(Fixture("iwork-converter/a.pages"));

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics,
            diagnostic => diagnostic.Code == "IWORK_PAGES_TEXT_UNSUPPORTED");
    }

    [Fact]
    public void Pages_recovers_embedded_image_and_shared_editable_tables() {
        IWorkPagesProjection pages = IWorkSourceDocument.Open(Fixture("picodocs/sample-v14.4.pages")).ReadPages();

        Assert.False(pages.HasEditableContent);
        Assert.Contains(pages.Diagnostics,
            diagnostic => diagnostic.Code == "IWORK_PAGES_TEXT_UNSUPPORTED");
        Assert.Equal(3, pages.Tables.Count);
        Assert.Equal((5, 4), (pages.Tables[0].RowCount, pages.Tables[0].ColumnCount));
        Assert.Equal(18, pages.Tables[0].Cells.Count);
        Assert.Equal(4, pages.Tables[2].Cells.Count(cell => cell.Kind == IWorkCellKind.Formula));
        IWorkImageAsset image = Assert.Single(pages.Images);
        Assert.Equal("image/png", image.MediaType);
        Assert.Equal(1000, image.PixelWidth);
        Assert.Equal(520, image.PixelHeight);
    }

    [Fact]
    public void Pages_owner_projects_tables_and_embedded_image_into_word() {
        using var result = WordDocument.LoadPagesWithReport(Fixture("picodocs/sample-v14.4.pages"));

        Assert.True(result.IsVisualFallback);
        Assert.Empty(result.Document.Tables);
        Assert.Single(result.Document.Images);
    }

    [Fact]
    public void Keynote_recovers_slide_size_positioned_rich_text_and_typography() {
        IWorkKeynoteProjection keynote = IWorkSourceDocument.Open(Fixture("iwork-converter/a.key")).ReadKeynote();

        Assert.Equal(1920d, keynote.SlideSize!.WidthPoints);
        Assert.Equal(1080d, keynote.SlideSize.HeightPoints);
        IWorkKeynoteSlide slide = Assert.Single(keynote.Slides);
        Assert.Equal(22, slide.TextBoxes.Count);
        IWorkTextBox box = slide.TextBoxes.First(item => item.Content.PlainText == "账户passport");
        Assert.InRange(box.Geometry!.LeftPoints, 371.74d, 371.75d);
        Assert.InRange(box.Geometry.TopPoints, 537.95d, 537.96d);
        IWorkTextRun run = Assert.Single(Assert.Single(box.Content.Paragraphs).Runs);
        Assert.Equal(IWorkTextAlignment.Center, box.Content.Paragraphs[0].Style.Alignment);
        Assert.Equal(28d, run.Style.FontSizePoints);
        Assert.Equal("HelveticaNeue-Medium", run.Style.FontName);
    }

    [Fact]
    public void Keynote_owner_projects_source_canvas_geometry_and_rich_text() {
        using var result = PowerPointPresentation.LoadKeynoteWithReport(Fixture("iwork-converter/a.key"));

        Assert.Equal(1920d, result.Document.SlideSize.WidthPoints, 3);
        Assert.Equal(1080d, result.Document.SlideSize.HeightPoints, 3);
        PowerPointSlide slide = Assert.Single(result.Document.Slides);
        PowerPointTextBox box = slide.TextBoxes.First(item => item.Text == "账户passport");
        Assert.InRange(box.LeftPoints, 371.74d, 371.75d);
        Assert.InRange(box.TopPoints, 537.95d, 537.96d);
        PowerPointTextRun run = Assert.Single(Assert.Single(box.Paragraphs).Runs);
        Assert.Equal(28d, run.FontSizePoints);
        Assert.Equal("HelveticaNeue-Medium", run.FontName);
    }

    [Fact]
    public void Keynote_recovers_independently_produced_editable_table_and_image() {
        IWorkKeynoteProjection tableDeck = IWorkSourceDocument.Open(
            Fixture("keynotekit/tabledeck-v15.2.1.key")).ReadKeynote();
        IWorkTable table = Assert.Single(Assert.Single(tableDeck.Slides).Tables);

        Assert.True(tableDeck.HasEditableContent);
        Assert.Equal((3, 3), (table.RowCount, table.ColumnCount));
        Assert.Equal("Product", table.GetCell(1, 1)!.Value);
        Assert.Equal(24_000d, Assert.IsType<double>(table.GetCell(2, 3)!.Value), 10);
        Assert.InRange(table.Geometry!.LeftPoints, 94.99d, 95.01d);
        Assert.InRange(table.Geometry.WidthPoints, 1729.99d, 1730.01d);

        IWorkKeynoteProjection imageDeck = IWorkSourceDocument.Open(
            Fixture("keynotekit/imagedeck-v15.2.1.key")).ReadKeynote();
        IWorkImageAsset image = Assert.Single(Assert.Single(imageDeck.Slides).Images);
        Assert.Equal("red.png", image.FileName);
        Assert.Equal("image/png", image.MediaType);
        Assert.Equal((400, 300), (image.PixelWidth, image.PixelHeight));
        byte[] firstRead = image.GetBytes();
        byte original = firstRead[0];
        firstRead[0] ^= 0xff;
        Assert.Equal(original, image.GetBytes()[0]);
    }

    [Fact]
    public void Keynote_owner_projects_table_and_image_as_editable_powerpoint_shapes() {
        using var tableResult = PowerPointPresentation.LoadKeynoteWithReport(
            Fixture("keynotekit/tabledeck-v15.2.1.key"));
        PowerPointTable table = Assert.Single(Assert.Single(tableResult.Document.Slides).Tables);
        IWorkTable sourceTable = Assert.Single(Assert.Single(
            tableResult.Projection.Slides).Tables);

        Assert.False(tableResult.IsVisualFallback);
        Assert.Equal((3, 3), (table.Rows, table.Columns));
        Assert.Equal("Product", table.GetCell(0, 0).Text);
        Assert.Equal("24000", table.GetCell(1, 2).Text);
        Assert.InRange(table.LeftPoints, 94.99d, 95.01d);
        double expectedWidth = sourceTable.DefaultColumnWidth!.Value * sourceTable.ColumnCount;
        double expectedHeight = sourceTable.DefaultRowHeight!.Value * sourceTable.RowCount;
        Assert.InRange(table.WidthPoints, expectedWidth - 0.001d, expectedWidth + 0.001d);
        Assert.InRange(table.HeightPoints, expectedHeight - 0.001d, expectedHeight + 0.001d);

        using var imageResult = PowerPointPresentation.LoadKeynoteWithReport(
            Fixture("keynotekit/imagedeck-v15.2.1.key"));
        PowerPointPicture picture = Assert.Single(Assert.Single(imageResult.Document.Slides).Pictures);
        Assert.Equal("image/png", picture.ContentType);
        Assert.True(picture.GetImageBytes().Length > 100);
    }

    [Fact]
    public void Semantic_owner_adapters_save_and_reopen_editable_outputs() {
        using var pages = WordDocument.LoadPagesWithReport(Fixture("nim-iwork/simple.pages"));
        using var numbers = ExcelDocument.LoadNumbersWithReport(Fixture("nim-iwork/simple.numbers"));
        using var keynote = PowerPointPresentation.LoadKeynoteWithReport(Fixture("nim-iwork/simple.key"));

        Assert.False(pages.IsVisualFallback);
        Assert.False(numbers.IsVisualFallback);
        Assert.False(keynote.IsVisualFallback);
        using var wordBytes = new MemoryStream();
        using var excelBytes = new MemoryStream();
        using var powerPointBytes = new MemoryStream();
        pages.Document.Save(wordBytes);
        numbers.Document.Save(excelBytes);
        keynote.Document.Save(powerPointBytes);

        wordBytes.Position = 0;
        excelBytes.Position = 0;
        powerPointBytes.Position = 0;
        using WordDocument word = WordDocument.Load(wordBytes);
        using ExcelDocument excel = ExcelDocument.Load(excelBytes);
        using PowerPointPresentation powerPoint = PowerPointPresentation.Load(powerPointBytes);
        Assert.Contains(word.Paragraphs, paragraph => paragraph.Text == "hello pages");
        Assert.Single(excel.Sheets);
        Assert.Equal("a", excel.Sheets[0].CellAt(1, 1).GetValue<string>());
        Assert.Equal(2, powerPoint.Slides.Count);
    }

    [Theory]
    [InlineData("nim-iwork/simple.pages", IWorkDocumentKind.Pages)]
    [InlineData("nim-iwork/simple.numbers", IWorkDocumentKind.Numbers)]
    [InlineData("nim-iwork/simple.key", IWorkDocumentKind.Keynote)]
    public void Explicit_visual_mode_is_reported_as_preview_fallback(string relativePath, IWorkDocumentKind kind) {
        var options = new IWorkReadOptions { ImportMode = IWorkImportMode.VisualOnly };
        IWorkImportReport report = kind switch {
            IWorkDocumentKind.Pages => ReadPagesVisual(relativePath, options),
            IWorkDocumentKind.Numbers => ReadNumbersVisual(relativePath, options),
            IWorkDocumentKind.Keynote => ReadKeynoteVisual(relativePath, options),
            _ => throw new ArgumentOutOfRangeException(nameof(kind))
        };

        Assert.Equal(IWorkProjectionKind.VisualFallback, report.ProjectionKind);
        Assert.NotNull(report.VisualPreview);
        Assert.Equal(0, report.ReconstructedItemCount);
        Assert.True(report.HasConversionLoss);
    }

    [Fact]
    public void Keynote_visual_fallback_letterboxes_non_widescreen_previews() {
        using var result = PowerPointPresentation.LoadKeynoteWithReport(
            Fixture("nim-iwork/simple.key"),
            new IWorkReadOptions { ImportMode = IWorkImportMode.VisualOnly });

        PowerPointPicture picture = Assert.Single(Assert.Single(result.Document.Slides).Pictures);
        Assert.Equal(10d, picture.WidthInches, 3);
        Assert.Equal(7.5d, picture.HeightInches, 3);
        Assert.InRange(picture.LeftInches, 1.666d, 1.667d);
        Assert.Equal(0d, picture.TopInches, 3);
    }

    [Fact]
    public void Rejects_configured_record_and_package_bounds() {
        Assert.Throws<InvalidDataException>(() => IWorkSourceDocument.Open(
            Fixture("nim-iwork/simple.pages"), new IWorkReadOptions { MaximumRecordCount = 1 }));
        Assert.Throws<InvalidDataException>(() => IWorkSourceDocument.Open(
            Fixture("nim-iwork/simple.pages"), new IWorkReadOptions { MaximumPackageBytes = 1024 }));
    }

    [Fact]
    public void Does_not_trust_a_wrong_expected_application_kind() {
        using FileStream stream = File.OpenRead(Fixture("nim-iwork/simple.numbers"));
        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            IWorkSourceDocument.Open(stream, IWorkDocumentKind.Pages));
        Assert.Contains("Numbers", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Can_disable_unsupported_record_reporting_without_discarding_the_source_records() {
        IWorkSourceDocument source = IWorkSourceDocument.Open(Fixture("nim-iwork/simple.pages"),
            new IWorkReadOptions { PreserveUnsupportedRecords = false });
        IWorkImportReport report = source.ReadPages().CreateImportReport(IWorkProjectionKind.EditableReconstruction);

        Assert.NotEmpty(source.Records);
        Assert.Empty(report.UnsupportedRecords);
        Assert.True(report.UnsupportedRecordCount > 0);
        Assert.True(report.HasConversionLoss);
        IWorkArchiveRecord populated = source.Records.First(record => record.PayloadLength > 0);
        byte original = populated.GetPayload()[0];
        byte[] copy = populated.GetPayload();
        copy[0] ^= 0xff;
        Assert.Equal(original, populated.GetPayload()[0]);
    }

    [Fact]
    public void Reads_an_extracted_directory_bundle_with_the_same_record_count() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-iwork-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            ZipFile.ExtractToDirectory(Fixture("nim-iwork/simple.pages"), directory);
            IWorkSourceDocument zipped = IWorkSourceDocument.Open(Fixture("nim-iwork/simple.pages"));
            IWorkSourceDocument bundle = IWorkSourceDocument.Open(directory, IWorkDocumentKind.Pages);

            Assert.Equal(IWorkContainerKind.DirectoryBundle, bundle.ContainerKind);
            Assert.Equal(zipped.Records.Count, bundle.Records.Count);
            Assert.Equal(zipped.ReadPages().Paragraphs, bundle.ReadPages().Paragraphs);
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void Reads_a_package_with_a_nested_index_zip() {
        using FileStream input = File.OpenRead(Fixture("nim-iwork/simple.pages"));
        using var sourceArchive = new ZipArchive(input, ZipArchiveMode.Read, leaveOpen: false);
        using var nestedBytes = new MemoryStream();
        using var packageBytes = new MemoryStream();
        using (var nested = new ZipArchive(nestedBytes, ZipArchiveMode.Create, leaveOpen: true)) {
            using var package = new ZipArchive(packageBytes, ZipArchiveMode.Create, leaveOpen: true);
            foreach (ZipArchiveEntry sourceEntry in sourceArchive.Entries.Where(entry => !string.IsNullOrEmpty(entry.Name))) {
                bool indexEntry = sourceEntry.FullName.StartsWith("Index/", StringComparison.Ordinal);
                string targetPath = indexEntry ? sourceEntry.FullName.Substring("Index/".Length) : sourceEntry.FullName;
                ZipArchiveEntry targetEntry = (indexEntry ? nested : package).CreateEntry(targetPath);
                using Stream source = sourceEntry.Open();
                using Stream target = targetEntry.Open();
                source.CopyTo(target);
            }
        }
        packageBytes.Position = 0;
        using (var package = new ZipArchive(packageBytes, ZipArchiveMode.Update, leaveOpen: true)) {
            ZipArchiveEntry nestedEntry = package.CreateEntry("Index.zip");
            using Stream target = nestedEntry.Open();
            nestedBytes.Position = 0;
            nestedBytes.CopyTo(target);
        }
        packageBytes.Position = 0;

        IWorkSourceDocument sourceDocument = IWorkSourceDocument.Open(packageBytes, IWorkDocumentKind.Pages);
        IWorkSourceDocument directDocument = IWorkSourceDocument.Open(Fixture("nim-iwork/simple.pages"));
        Assert.Equal(IWorkContainerKind.ZipPackageWithNestedIndex, sourceDocument.ContainerKind);
        Assert.Equal(directDocument.Records.Count, sourceDocument.Records.Count);
        Assert.Equal(directDocument.ReadPages().Paragraphs, sourceDocument.ReadPages().Paragraphs);
    }

    [Fact]
    public void Rejects_unsafe_package_entry_paths() {
        using var source = new MemoryStream();
        using (var archive = new ZipArchive(source, ZipArchiveMode.Create, leaveOpen: true)) {
            ZipArchiveEntry entry = archive.CreateEntry("../Index/Document.iwa");
            using Stream target = entry.Open();
            target.WriteByte(0);
        }
        source.Position = 0;

        Assert.Throws<InvalidDataException>(() => IWorkSourceDocument.Open(source, IWorkDocumentKind.Pages));
    }

    private static IWorkImportReport ReadPagesVisual(string relativePath, IWorkReadOptions options) {
        using var result = WordDocument.LoadPagesWithReport(Fixture(relativePath), options);
        return result.ImportReport;
    }

    private static IWorkImportReport ReadNumbersVisual(string relativePath, IWorkReadOptions options) {
        using var result = ExcelDocument.LoadNumbersWithReport(Fixture(relativePath), options);
        return result.ImportReport;
    }

    private static IWorkImportReport ReadKeynoteVisual(string relativePath, IWorkReadOptions options) {
        using var result = PowerPointPresentation.LoadKeynoteWithReport(Fixture(relativePath), options);
        return result.ImportReport;
    }

    private static string Fixture(string relativePath) =>
        Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus",
            relativePath.Replace('/', Path.DirectorySeparatorChar));
}
