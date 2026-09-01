using OfficeIMO.Excel;
using OfficeIMO.IWork;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;
using System.Threading.Tasks;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Numbers_table_accessibility_descriptions_use_visual_fallback() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Accessible", 1, 1, 42d)
        }, includePreview: true, tableDrawable: Message(StringField(8, "Source table")));

        using var result = ExcelDocument.LoadNumbersWithReport(package);
        IWorkTable sourceTable = Assert.Single(Assert.Single(result.Projection.Sheets).Tables);

        Assert.True(result.IsVisualFallback);
        Assert.Equal("Source table", sourceTable.AccessibilityDescription);
        Assert.Contains(result.Projection.Diagnostics,
            diagnostic => diagnostic.Code == "IWORK_NUMBERS_TABLE_ACCESSIBILITY_UNSUPPORTED");
    }

    [Fact]
    public void Pages_table_accessibility_descriptions_are_preserved_in_word() {
        using MemoryStream package = CreatePagesPackageWithTableGeometry(
            0f, 0f, 0f, 0f, 0f, includePreview: false,
            accessibilityDescription: "Source table");

        using var result = WordDocument.LoadPagesWithReport(package);
        using var saved = new MemoryStream();
        result.Document.Save(saved);
        saved.Position = 0;
        using WordDocument reopened = WordDocument.Load(saved);

        Assert.False(result.IsVisualFallback);
        Assert.Equal("Source table", Assert.Single(result.Projection.Tables).AccessibilityDescription);
        Assert.Equal("Source table", Assert.Single(reopened.Tables).Description);
    }

    [Fact]
    public void Keynote_table_accessibility_descriptions_are_preserved_in_powerpoint() {
        using MemoryStream package = CreateKeynotePackageWithTableDefaults(
            rows: 1, columns: 1, defaultRowHeight: 20d, defaultColumnWidth: 40d,
            accessibilityDescription: "Source table");

        using var result = PowerPointPresentation.LoadKeynoteWithReport(package);
        using var saved = new MemoryStream();
        result.Document.Save(saved);
        saved.Position = 0;
        using PowerPointPresentation reopened = PowerPointPresentation.Load(saved);

        Assert.False(result.IsVisualFallback);
        Assert.Equal("Source table", Assert.Single(Assert.Single(
            result.Projection.Slides).Tables).AccessibilityDescription);
        Assert.Equal("Source table", Assert.Single(Assert.Single(reopened.Slides).Tables).AltText);
    }

    [Fact]
    public void Malformed_table_accessibility_metadata_disables_editable_reconstruction() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Malformed", 1, 1, 42d)
        }, includePreview: true, tableDrawable: Message(VarintField(8, 1)));

        using var result = ExcelDocument.LoadNumbersWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics,
            diagnostic => diagnostic.Code == "IWORK_TABLE_DRAWABLE_UNSUPPORTED");
    }

    [Fact]
    public async Task Top_level_fifo_paths_and_links_are_rejected_without_blocking() {
        if (OperatingSystem.IsWindows()) return;

        string directory = Path.Combine(Path.GetTempPath(),
            "officeimo-iwork-top-level-fifo-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        string fifo = Path.Combine(directory, "blocking.pages");
        string link = Path.Combine(directory, "linked.pages");
        try {
            Assert.Equal(0, CreateFifo(fifo, 0x180));
            File.CreateSymbolicLink(link, fifo);

            foreach (string sourcePath in new[] { fifo, link }) {
                Task<Exception?> open = Task.Run<Exception?>(() => Record.Exception(() =>
                    IWorkSourceDocument.Open(sourcePath, IWorkDocumentKind.Pages)));
                bool completedWithoutWriter = ReferenceEquals(
                    await Task.WhenAny(open, Task.Delay(TimeSpan.FromSeconds(2))), open);
                if (!completedWithoutWriter) {
                    await Task.Run(() => {
                        using var writer = new FileStream(fifo, FileMode.Open, FileAccess.Write,
                            FileShare.ReadWrite);
                    });
                    await open.WaitAsync(TimeSpan.FromSeconds(2));
                }

                Assert.True(completedWithoutWriter,
                    "Opening a top-level FIFO source blocked while waiting for a writer.");
                Assert.IsType<InvalidDataException>(await open);
            }
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void Directory_bundles_with_nested_indexes_retain_directory_identity() {
        string directory = Path.Combine(Path.GetTempPath(),
            "officeimo-iwork-nested-directory-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            using (ZipArchive archive = ZipFile.Open(Path.Combine(directory, "Index.zip"),
                       ZipArchiveMode.Create)) {
                WriteEntry(archive, "Document.iwa", FrameIwa(Message(
                    ArchiveRecord(1, 1, Message()))));
            }

            IWorkSourceDocument source = IWorkSourceDocument.Open(
                directory, IWorkDocumentKind.Numbers);

            Assert.Equal(IWorkContainerKind.DirectoryBundle, source.ContainerKind);
            Assert.Contains(source.Entries, entry => entry.Path == "Index.zip");
            Assert.Contains(source.Entries, entry => entry.Path == "Index/Document.iwa");
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void Keynote_import_reports_count_metadata_only_title_and_text_shapes() {
        using MemoryStream package = CreateKeynotePackageWithMetadataOnlyTextBoxes();

        using var result = PowerPointPresentation.LoadKeynoteWithReport(package);

        Assert.False(result.IsVisualFallback);
        Assert.Equal(3, result.ImportReport.ReconstructedItemCount);
    }
}
