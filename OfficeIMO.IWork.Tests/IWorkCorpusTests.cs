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
        Assert.Contains("second paragraph with some words", pages.Paragraphs[1], StringComparison.Ordinal);
        Assert.True(report.TotalRecordCount >= report.UnsupportedRecords.Count);
        Assert.NotEmpty(report.UnsupportedRecords);
        Assert.True(report.HasConversionLoss);
    }

    [Fact]
    public void Reads_current_numbers_as_sparse_typed_cells() {
        IWorkNumbersProjection numbers = IWorkSourceDocument.Open(Fixture("nim-iwork/simple.numbers")).ReadNumbers();
        IWorkNumbersTable table = Assert.Single(Assert.Single(numbers.Sheets).Tables);

        Assert.Equal(3, table.RowCount);
        Assert.Equal(3, table.ColumnCount);
        Assert.Equal("a", table.GetCell(1, 1)!.Value);
        Assert.Equal("C", table.GetCell(1, 3)!.Value);
        Assert.Equal(2d, Assert.IsType<double>(table.GetCell(2, 2)!.Value), 10);
        Assert.Equal("Z", table.GetCell(3, 3)!.Value);
        Assert.Throws<ArgumentOutOfRangeException>(() => table.GetCell(2, 4));
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
        IWorkNumbersTable newerTable = Assert.Single(Assert.Single(newerNumbers.Sheets).Tables);
        Assert.Equal(7, newerTable.RowCount);
        Assert.Equal(11, newerTable.ColumnCount);
        Assert.Equal("Cats", newerTable.GetCell(1, 3)!.Value);
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
        Assert.True(report.HasConversionLoss);
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
