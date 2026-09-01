using System.IO.Compression;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Word;

namespace OfficeIMO.Workflows.Tests;

public sealed class OfficeOutputWorkflowTests {
    [Fact]
    public async Task PageImageExportPreservesCallerSelectionAndPublishesValidatedFolder() {
        using var scope = new TestDirectory();
        string input = CreatePdf(scope.Path, "source.pdf", "One", "Two", "Three");
        string output = Path.Combine(scope.Path, "pages");

        PdfPageImageExportResult result = await new OfficeWorkflowRunner().ExportPdfPagesAsync(
            new PdfPageImageExportRequest {
                InputPath = input,
                OutputDirectory = output,
                Pages = "3,1",
                Format = OfficeImageExportFormat.Png,
                ConflictPolicy = OfficeWorkflowConflictPolicy.Fail
            });

        Assert.True(result.Succeeded, result.Summary);
        Assert.Equal(output, result.OutputDirectory);
        Assert.Equal([3, 1], result.Files.Select(static file => file.PageNumber));
        Assert.All(result.Files, file => {
            Assert.True(File.Exists(file.Path));
            Assert.True(OfficeImageReader.TryValidateContent(File.ReadAllBytes(file.Path), file.Path, out OfficeImageInfo info));
            Assert.Equal(file.Width, info.Width);
            Assert.Equal(file.Height, info.Height);
        });
        Assert.Contains(result.Diagnostics, static diagnostic => diagnostic.Code == "PageImagesReopened");
        Assert.Empty(Directory.GetDirectories(scope.Path, ".*.tmp"));
    }

    [Fact]
    public async Task PageImageExportRestoresAndReplacesOneInterruptedRecoveryDirectory() {
        using var scope = new TestDirectory();
        string input = CreatePdf(scope.Path, "source.pdf", "Replacement");
        string output = Path.Combine(scope.Path, "pages");
        string recovery = output + ".officeimo-recovery-interrupted";
        Directory.CreateDirectory(recovery);
        await File.WriteAllTextAsync(Path.Combine(recovery, "previous.txt"), "previous output");

        PdfPageImageExportResult result = await new OfficeWorkflowRunner().ExportPdfPagesAsync(
            new PdfPageImageExportRequest {
                InputPath = input,
                OutputDirectory = output,
                ConflictPolicy = OfficeWorkflowConflictPolicy.Replace
            });

        Assert.True(result.Succeeded, result.Summary);
        Assert.Equal(OfficeWorkflowFailureKind.None, result.FailureKind);
        Assert.True(Directory.Exists(output));
        Assert.Single(Directory.GetFiles(output, "*.png"));
        Assert.False(File.Exists(Path.Combine(output, "previous.txt")));
        Assert.Empty(Directory.GetDirectories(scope.Path, "pages.officeimo-recovery-*"));
    }

    [Fact]
    public async Task PageImageExportRefusesAmbiguousInterruptedReplacementWithRecoveryDetails() {
        using var scope = new TestDirectory();
        string input = CreatePdf(scope.Path, "source.pdf", "Replacement");
        string output = Path.Combine(scope.Path, "pages");
        string recovery = output + ".officeimo-recovery-interrupted";
        Directory.CreateDirectory(output);
        Directory.CreateDirectory(recovery);
        await File.WriteAllTextAsync(Path.Combine(output, "current.txt"), "current output");
        await File.WriteAllTextAsync(Path.Combine(recovery, "previous.txt"), "previous output");

        PdfPageImageExportResult result = await new OfficeWorkflowRunner().ExportPdfPagesAsync(
            new PdfPageImageExportRequest {
                InputPath = input,
                OutputDirectory = output,
                ConflictPolicy = OfficeWorkflowConflictPolicy.Replace
            });

        Assert.Equal(OfficeWorkflowStatus.Failed, result.Status);
        Assert.Equal(OfficeWorkflowFailureKind.OutputFailed, result.FailureKind);
        OfficeWorkflowDiagnostic diagnostic = Assert.Single(result.Diagnostics, static item => item.Code == "PageImageExportFailed");
        Assert.Equal(output, diagnostic.Details["destination"]);
        Assert.Contains(recovery, diagnostic.Details["recoveryPaths"], StringComparison.Ordinal);
        Assert.True(File.Exists(Path.Combine(output, "current.txt")));
        Assert.True(File.Exists(Path.Combine(recovery, "previous.txt")));
        Assert.Empty(Directory.GetDirectories(scope.Path, ".pages.*.tmp"));
    }

    [Fact]
    public async Task ConcurrentPageImageReplacementsSerializeWithoutLeavingRecoveryDirectories() {
        using var scope = new TestDirectory();
        string firstInput = CreatePdf(scope.Path, "first.pdf", "First");
        string secondInput = CreatePdf(scope.Path, "second.pdf", "Second");
        string output = Path.Combine(scope.Path, "pages");
        Directory.CreateDirectory(output);
        await File.WriteAllTextAsync(Path.Combine(output, "previous.txt"), "previous output");
        var runner = new OfficeWorkflowRunner();

        Task<PdfPageImageExportResult> first = runner.ExportPdfPagesAsync(new PdfPageImageExportRequest {
            InputPath = firstInput,
            OutputDirectory = output,
            ConflictPolicy = OfficeWorkflowConflictPolicy.Replace
        });
        Task<PdfPageImageExportResult> second = runner.ExportPdfPagesAsync(new PdfPageImageExportRequest {
            InputPath = secondInput,
            OutputDirectory = output,
            ConflictPolicy = OfficeWorkflowConflictPolicy.Replace
        });
        PdfPageImageExportResult[] results = await Task.WhenAll(first, second);

        Assert.All(results, result => Assert.True(result.Succeeded, result.Summary));
        Assert.Single(Directory.GetFiles(output, "*.png"));
        Assert.Empty(Directory.GetDirectories(scope.Path, "pages.officeimo-recovery-*"));
    }

    [Fact]
    public async Task AssemblyNormalizesPdfImageAndWordThroughFirstPartyOwners() {
        using var scope = new TestDirectory();
        string pdf = CreatePdf(scope.Path, "first.pdf", "PDF source");
        string image = Path.Combine(scope.Path, "image.png");
        await File.WriteAllBytesAsync(image, CreatePng("Image source"));
        string word = Path.Combine(scope.Path, "third.docx");
        using (WordDocument document = WordDocument.Create(word)) {
            document.AddParagraph("Word source");
            document.Save();
        }
        string output = Path.Combine(scope.Path, "assembled.pdf");

        PdfAssemblyResult result = await new OfficeWorkflowRunner().AssemblePdfAsync(new PdfAssemblyRequest {
            Sources = [pdf, image, word],
            OutputPath = output,
            ConflictPolicy = OfficeWorkflowConflictPolicy.Fail
        });

        Assert.True(result.Succeeded, result.Summary);
        Assert.Equal(3, result.SourceCount);
        Assert.Equal(3, result.PageCount);
        Assert.Equal(3, PdfDocument.Load(output).Inspect().PageCount);
        Assert.Equal(
            ["first.pdf", "image.png", "third.docx"],
            result.Diagnostics
                .Where(static item => item.Code == "AssemblySourceNormalized")
                .Select(static item => item.Details["name"]));
        Assert.Contains(result.Diagnostics, static item => item.Code == "RouteContract" && item.Details["route"] == "docx-pdf");
        Assert.Contains(result.Diagnostics, static item => item.Code == "AssemblyReopened");
    }

    [Fact]
    public async Task FolderAndZipExpansionAreDeterministicAndIgnoreUnrelatedFiles() {
        using var scope = new TestDirectory();
        string folder = Path.Combine(scope.Path, "folder");
        Directory.CreateDirectory(folder);
        string folderB = CreatePdf(folder, "02-b.pdf", "Folder B");
        string folderA = CreatePdf(folder, "01-a.pdf", "Folder A");
        Assert.False(PdfDocument.Load(folderA).Inspect().HasTaggedContent);
        Assert.False(PdfDocument.Load(folderB).Inspect().HasTaggedContent);
        await File.WriteAllTextAsync(Path.Combine(folder, "notes.txt"), "ignored");

        string zipOne = CreatePdf(scope.Path, "zip-one.pdf", "ZIP One");
        string zipTwo = CreatePdf(scope.Path, "zip-two.pdf", "ZIP Two");
        string archivePath = Path.Combine(scope.Path, "documents.zip");
        using (ZipArchive archive = ZipFile.Open(archivePath, ZipArchiveMode.Create)) {
            AddArchiveFile(archive, zipTwo, "02-two.pdf");
            AddArchiveFile(archive, zipOne, "01-one.pdf");
            ZipArchiveEntry ignored = archive.CreateEntry("readme.txt");
            await using StreamWriter writer = new(ignored.Open());
            await writer.WriteAsync("ignored");
        }

        PdfAssemblyResult folderOnly = await new OfficeWorkflowRunner().AssemblePdfAsync(new PdfAssemblyRequest {
            Sources = [folder],
            OutputPath = Path.Combine(scope.Path, "folder-only.pdf"),
            ConflictPolicy = OfficeWorkflowConflictPolicy.Fail
        });
        Assert.True(folderOnly.Succeeded, "Folder-only intake: " + folderOnly.Summary + " | " +
            string.Join("; ", folderOnly.Diagnostics.Select(static item => item.Code + ":" + item.Message)));
        PdfAssemblyResult archiveOnly = await new OfficeWorkflowRunner().AssemblePdfAsync(new PdfAssemblyRequest {
            Sources = [archivePath],
            OutputPath = Path.Combine(scope.Path, "archive-only.pdf"),
            ConflictPolicy = OfficeWorkflowConflictPolicy.Fail
        });
        Assert.True(archiveOnly.Succeeded, "Archive-only intake: " + archiveOnly.Summary);

        string output = Path.Combine(scope.Path, "ordered.pdf");
        PdfAssemblyResult result = await new OfficeWorkflowRunner().AssemblePdfAsync(new PdfAssemblyRequest {
            Sources = [folder, archivePath],
            OutputPath = output,
            ConflictPolicy = OfficeWorkflowConflictPolicy.Fail
        });

        Assert.True(result.Succeeded, result.Summary);
        Assert.Equal(4, result.SourceCount);
        PdfReadDocument readDocument = PdfReadDocument.Open(File.ReadAllBytes(output));
        Assert.Equal(["Folder A", "Folder B", "ZIP One", "ZIP Two"],
            readDocument.Pages.Select(static page => page.ExtractText().Trim()));
        Assert.Contains(result.Diagnostics, static item => item.Code == "ArchiveExpanded");
    }

    [Fact]
    public async Task FolderDiscoveryLimitCountsUnsupportedFilesBeforeFiltering() {
        using var scope = new TestDirectory();
        string folder = Path.Combine(scope.Path, "folder");
        Directory.CreateDirectory(folder);
        await File.WriteAllTextAsync(Path.Combine(folder, "one.bin"), "1");
        await File.WriteAllTextAsync(Path.Combine(folder, "two.bin"), "2");
        await File.WriteAllTextAsync(Path.Combine(folder, "three.bin"), "3");
        string output = Path.Combine(scope.Path, "must-not-exist.pdf");

        PdfAssemblyResult result = await new OfficeWorkflowRunner().AssemblePdfAsync(new PdfAssemblyRequest {
            Sources = [folder],
            OutputPath = output,
            Options = new PdfAssemblyOptions { MaximumDiscoveredEntries = 2 }
        });

        Assert.Equal(OfficeWorkflowStatus.Failed, result.Status);
        Assert.Contains("Discovered entry count", result.Summary, StringComparison.Ordinal);
        Assert.False(File.Exists(output));
    }

    [Fact]
    public async Task DiscoveryLimitIsAggregateAcrossArchivesAndCountsIgnoredEntries() {
        using var scope = new TestDirectory();
        string firstArchive = Path.Combine(scope.Path, "first.zip");
        string secondArchive = Path.Combine(scope.Path, "second.zip");
        CreateTextArchive(firstArchive, "one.txt", "two.txt");
        CreateTextArchive(secondArchive, "three.txt", "four.txt");
        string output = Path.Combine(scope.Path, "must-not-exist.pdf");

        PdfAssemblyResult result = await new OfficeWorkflowRunner().AssemblePdfAsync(new PdfAssemblyRequest {
            Sources = [firstArchive, secondArchive],
            OutputPath = output,
            Options = new PdfAssemblyOptions { MaximumDiscoveredEntries = 3 }
        });

        Assert.Equal(OfficeWorkflowStatus.Failed, result.Status);
        Assert.Contains("Discovered entry count", result.Summary, StringComparison.Ordinal);
        Assert.False(File.Exists(output));
    }

    [Fact]
    public async Task DeclaredZipEntryLimitIsRejectedBeforeArchiveMetadataMaterialization() {
        using var scope = new TestDirectory();
        string archive = Path.Combine(scope.Path, "declared-many.zip");
        CreateDeclaredEntryCountArchive(archive, 5_000);
        string output = Path.Combine(scope.Path, "must-not-exist.pdf");

        PdfAssemblyResult result = await new OfficeWorkflowRunner().AssemblePdfAsync(new PdfAssemblyRequest {
            Sources = [archive],
            OutputPath = output,
            Options = new PdfAssemblyOptions { MaximumArchiveEntries = 10 }
        });

        Assert.Equal(OfficeWorkflowStatus.Failed, result.Status);
        Assert.Equal(OfficeWorkflowFailureKind.UnsupportedInput, result.FailureKind);
        Assert.Contains("declared-many.zip", result.Summary, StringComparison.Ordinal);
        Assert.Contains("10-entry limit", result.Summary, StringComparison.Ordinal);
        Assert.False(File.Exists(output));
    }

    [Fact]
    public async Task CaseDistinctFolderInputsRemainDistinctOnCaseSensitivePlatforms() {
        if (OperatingSystem.IsWindows()) return;
        using var scope = new TestDirectory();
        string folder = Path.Combine(scope.Path, "folder");
        Directory.CreateDirectory(folder);
        CreatePdf(folder, "A.pdf", "Uppercase");
        CreatePdf(folder, "a.pdf", "Lowercase");
        string output = Path.Combine(scope.Path, "assembled.pdf");

        PdfAssemblyResult result = await new OfficeWorkflowRunner().AssemblePdfAsync(new PdfAssemblyRequest {
            Sources = [folder],
            OutputPath = output,
            ConflictPolicy = OfficeWorkflowConflictPolicy.Fail
        });

        Assert.True(result.Succeeded, result.Summary);
        Assert.Equal(2, result.SourceCount);
        Assert.Equal(["Uppercase", "Lowercase"],
            PdfReadDocument.Open(File.ReadAllBytes(output)).Pages.Select(static page => page.ExtractText().Trim()));
    }

    [Fact]
    public async Task PhysicalPathAliasCannotBeUsedAsBothAssemblyInputAndOutput() {
        if (OperatingSystem.IsWindows()) return;
        using var scope = new TestDirectory();
        string input = CreatePdf(scope.Path, "source.pdf", "Source");
        string alias = Path.Combine(scope.Path, "alias.pdf");
        File.CreateSymbolicLink(alias, input);

        PdfAssemblyResult result = await new OfficeWorkflowRunner().AssemblePdfAsync(new PdfAssemblyRequest {
            Sources = [input],
            OutputPath = alias,
            ConflictPolicy = OfficeWorkflowConflictPolicy.Replace
        });

        Assert.Equal(OfficeWorkflowStatus.Failed, result.Status);
        Assert.Equal(OfficeWorkflowFailureKind.ValidationFailed, result.FailureKind);
        Assert.Contains("explicit input", result.Summary, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task ZipTraversalFailsBeforeAnyOutputIsPublished() {
        using var scope = new TestDirectory();
        string source = CreatePdf(scope.Path, "source.pdf", "Traversal payload");
        string archivePath = Path.Combine(scope.Path, "traversal.zip");
        using (ZipArchive archive = ZipFile.Open(archivePath, ZipArchiveMode.Create)) {
            AddArchiveFile(archive, source, "../outside.pdf");
        }
        string output = Path.Combine(scope.Path, "must-not-exist.pdf");

        PdfAssemblyResult result = await new OfficeWorkflowRunner().AssemblePdfAsync(new PdfAssemblyRequest {
            Sources = [archivePath],
            OutputPath = output,
            ConflictPolicy = OfficeWorkflowConflictPolicy.Fail
        });

        Assert.Equal(OfficeWorkflowStatus.Failed, result.Status);
        Assert.Contains("outside", result.Summary, StringComparison.OrdinalIgnoreCase);
        Assert.False(File.Exists(output));
    }

    [Fact]
    public async Task PreCancelledAssemblyLeavesNoPublishedOrStagedArtifact() {
        using var scope = new TestDirectory();
        string input = CreatePdf(scope.Path, "source.pdf", "Cancelled");
        string output = Path.Combine(scope.Path, "cancelled.pdf");
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        PdfAssemblyResult result = await new OfficeWorkflowRunner().AssemblePdfAsync(new PdfAssemblyRequest {
            Sources = [input],
            OutputPath = output
        }, cancellationToken: cancellation.Token);

        Assert.Equal(OfficeWorkflowStatus.Cancelled, result.Status);
        Assert.False(File.Exists(output));
        Assert.Empty(Directory.GetFiles(scope.Path, ".*.tmp"));
    }

    [Fact]
    public void PrintPlannerCreatesOrderedTwoUpGeometryWithinResolvedPaper() {
        using var scope = new TestDirectory();
        string input = CreatePdf(scope.Path, "print.pdf", "First page", "Second page");

        PdfPrintPlan plan = PdfPrintPlanner.Create(new PdfPrintPlanRequest {
            InputPath = input,
            Pages = "2,1",
            PagesPerSheet = 2,
            PaperSize = PageSizes.A4,
            Orientation = PdfPrintOrientation.Landscape,
            ScaleMode = PdfPrintScaleMode.Fit,
            Margin = 18D
        });

        Assert.Equal([2, 1], plan.SelectedPages);
        PdfPrintSheet sheet = Assert.Single(plan.Sheets);
        Assert.True(sheet.PaperSize.Width > sheet.PaperSize.Height);
        Assert.Equal([2, 1], sheet.Placements.Select(static placement => placement.PageNumber));
        Assert.All(sheet.Placements, placement => {
            Assert.True(placement.X >= 0D);
            Assert.True(placement.Y >= 0D);
            Assert.True(placement.X + placement.Width <= sheet.PaperSize.Width + 0.01D);
            Assert.True(placement.Y + placement.Height <= sheet.PaperSize.Height + 0.01D);
            Assert.False(placement.IsClipped);
        });
    }

    [Fact]
    public void FillPrintPlanExposesSlotBoundsForClippingOversizedContent() {
        using var scope = new TestDirectory();
        string input = CreatePdf(scope.Path, "fill-print.pdf", "Wide", "Tall");

        PdfPrintPlan plan = PdfPrintPlanner.Create(new PdfPrintPlanRequest {
            InputPath = input,
            PagesPerSheet = 2,
            PaperSize = PageSizes.A4,
            Orientation = PdfPrintOrientation.Portrait,
            ScaleMode = PdfPrintScaleMode.Fill,
            Margin = 18D
        });

        PdfPrintSheet sheet = Assert.Single(plan.Sheets);
        Assert.All(sheet.Placements, placement => {
            Assert.True(placement.IsClipped);
            Assert.True(placement.SlotWidth > 0D);
            Assert.True(placement.SlotHeight > 0D);
            Assert.True(placement.X < placement.SlotX || placement.Y < placement.SlotY ||
                        placement.Width > placement.SlotWidth || placement.Height > placement.SlotHeight);
            Assert.True(placement.SlotX + placement.SlotWidth <= sheet.PaperSize.Width);
            Assert.True(placement.SlotY + placement.SlotHeight <= sheet.PaperSize.Height);
        });
    }

    private static string CreatePdf(string root, string fileName, params string[] pages) {
        string path = Path.Combine(root, fileName);
        PdfDocument.Create(compose => {
            foreach (string text in pages) {
                compose.Page(page => page.Content(content => content.Item(item => item.Paragraph(paragraph => paragraph.Text(text)))));
            }
        }).Save(path);
        return path;
    }

    private static byte[] CreatePng(string text) => PdfDocument.Create(compose =>
            compose.Page(page => page.Content(content => content.Item(item => item.Paragraph(paragraph => paragraph.Text(text))))))
        .ExportImages(OfficeImageExportFormat.Png)
        .Single()
        .Bytes;

    private static void AddArchiveFile(ZipArchive archive, string sourcePath, string entryName) {
        ZipArchiveEntry entry = archive.CreateEntry(entryName, CompressionLevel.Fastest);
        using Stream source = File.OpenRead(sourcePath);
        using Stream destination = entry.Open();
        source.CopyTo(destination);
    }

    private static void CreateTextArchive(string path, params string[] entryNames) {
        using ZipArchive archive = ZipFile.Open(path, ZipArchiveMode.Create);
        foreach (string entryName in entryNames) {
            ZipArchiveEntry entry = archive.CreateEntry(entryName);
            using var writer = new StreamWriter(entry.Open());
            writer.Write(entryName);
        }
    }

    private static void CreateDeclaredEntryCountArchive(string path, ushort entryCount) {
        using FileStream stream = File.Create(path);
        using var writer = new BinaryWriter(stream);
        writer.Write(0x06054B50U);
        writer.Write((ushort)0);
        writer.Write((ushort)0);
        writer.Write(entryCount);
        writer.Write(entryCount);
        writer.Write(0U);
        writer.Write(0U);
        writer.Write((ushort)0);
    }

    private sealed class TestDirectory : IDisposable {
        public TestDirectory() {
            Path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "officeimo-output-workflows-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(Path);
        }

        public string Path { get; }

        public void Dispose() {
            try {
                Directory.Delete(Path, recursive: true);
            } catch (IOException) {
                // Test cleanup is best effort on platforms that briefly retain package streams.
            } catch (UnauthorizedAccessException) {
                // Test cleanup is best effort on platforms that briefly retain package streams.
            }
        }
    }
}
