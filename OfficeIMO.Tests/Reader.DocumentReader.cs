using OfficeIMO.Excel;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.Reader;
using OfficeIMO.Word;
using System;
using System.IO;
using System.Linq;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>
/// Smoke tests for <see cref="DocumentReader"/> across supported formats.
/// </summary>
public sealed class ReaderDocumentReaderTests {
    [Fact]
    public void DocumentReader_CanReadWord() {
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");
        try {
            using (var doc = WordDocument.Create(path)) {
                doc.AddParagraph("Policy").Style = WordParagraphStyles.Heading1;
                doc.AddParagraph("This is the body.");
                doc.Save();
            }

            var bytes = File.ReadAllBytes(path);
            var chunks = DocumentReader.Read(bytes, "Policy.docx").ToList();
            Assert.NotEmpty(chunks);
            Assert.Contains(chunks, c => c.Kind == ReaderInputKind.Word && (c.Markdown ?? c.Text).Contains("# Policy", StringComparison.Ordinal));
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void DocumentReader_CanReadExcel() {
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
        try {
            using (var doc = ExcelDocument.Create(path)) {
                var sheet = doc.AddWorkSheet("Data");
                sheet.Cell(1, 1, "Name");
                sheet.Cell(1, 2, "Value");
                sheet.Cell(2, 1, "A");
                sheet.Cell(2, 2, 1);
                doc.Save();
            }

            var chunks = DocumentReader.Read(path, new ReaderOptions { ExcelSheetName = "Data", ExcelA1Range = "A1:B2" }).ToList();
            Assert.NotEmpty(chunks);
            Assert.Contains(chunks, c =>
                c.Kind == ReaderInputKind.Excel &&
                c.Tables != null &&
                c.Tables.Count > 0 &&
                c.Tables[0].Columns.Contains("Name", StringComparer.Ordinal));
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void DocumentReader_CanReadPowerPoint() {
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
        try {
            using (var presentation = PowerPointPresentation.Create(path)) {
                var slide = presentation.AddSlide();
                slide.AddTextBox("Hello Reader");
                slide.Notes.Text = "Notes";
                presentation.Save();
            }

            var chunks = DocumentReader.Read(path).ToList();
            Assert.NotEmpty(chunks);
            Assert.Contains(chunks, c => c.Kind == ReaderInputKind.PowerPoint && (c.Markdown ?? c.Text).Contains("Hello Reader", StringComparison.Ordinal));
            Assert.Contains(chunks, c => c.Kind == ReaderInputKind.PowerPoint && (c.Markdown ?? c.Text).Contains("Notes", StringComparison.Ordinal));
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void DocumentReader_CanReadMarkdown() {
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".md");
        try {
            File.WriteAllText(path, "# Top\n\nPara 1.\n\n## Child\n\nPara 2.\n");

            using var fs = File.OpenRead(path);
            var chunks = DocumentReader.Read(fs, "Notes.md").ToList();
            Assert.True(chunks.Count >= 2);
            Assert.Contains(chunks, c => c.Kind == ReaderInputKind.Markdown && (c.Location.HeadingPath?.Contains("Top", StringComparison.Ordinal) ?? false));
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void DocumentReader_CanReadPdf() {
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pdf");
        try {
            var pdf = PdfDoc.Create();
            pdf.H1("PDF Title");
            pdf.Paragraph(p => p.Text("This is a PDF body."));
            pdf.Save(path);

            var chunks = DocumentReader.Read(path).ToList();
            Assert.NotEmpty(chunks);
            Assert.Contains(chunks, c => c.Kind == ReaderInputKind.Pdf);
            Assert.Contains(chunks, c => c.Location.Page.HasValue && c.Location.Page.Value >= 1);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void DocumentReader_ReadFolder_IsDeterministicWithMaxFiles() {
        var folder = Path.Combine(Path.GetTempPath(), "officeimo-reader-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        var a = Path.Combine(folder, "a.md");
        var b = Path.Combine(folder, "b.md");

        try {
            File.WriteAllText(a, "alpha");
            File.WriteAllText(b, "beta");

            var chunks = DocumentReader.ReadFolder(
                folderPath: folder,
                folderOptions: new ReaderFolderOptions {
                    Recurse = false,
                    MaxFiles = 1,
                    DeterministicOrder = true
                }).ToList();

            Assert.NotEmpty(chunks);
            Assert.All(chunks, c => Assert.Contains("a.md", c.Location.Path ?? string.Empty, StringComparison.OrdinalIgnoreCase));
            Assert.DoesNotContain(chunks, c => (c.Text ?? string.Empty).Contains("beta", StringComparison.OrdinalIgnoreCase));
        } finally {
            if (Directory.Exists(folder)) Directory.Delete(folder, recursive: true);
        }
    }

    [Fact]
    public void DocumentReader_ReadFolder_SkipsBrokenFilesAndContinues() {
        var folder = Path.Combine(Path.GetTempPath(), "officeimo-reader-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        var badDocx = Path.Combine(folder, "broken.docx");
        var goodMarkdown = Path.Combine(folder, "good.md");

        try {
            // Not a real DOCX package; this should fail to parse and be skipped.
            File.WriteAllText(badDocx, "not-a-zip-package");
            File.WriteAllText(goodMarkdown, "# Ok\n\nBody");

            var chunks = DocumentReader.ReadFolder(
                folderPath: folder,
                folderOptions: new ReaderFolderOptions {
                    Recurse = false,
                    MaxFiles = 10,
                    DeterministicOrder = true
                }).ToList();

            Assert.NotEmpty(chunks);
            Assert.Contains(chunks, c => c.Kind == ReaderInputKind.Markdown && (c.Text ?? string.Empty).Contains("Body", StringComparison.Ordinal));
            Assert.Contains(chunks, c =>
                c.Kind == ReaderInputKind.Unknown &&
                string.Equals(c.Location.Path, badDocx, StringComparison.OrdinalIgnoreCase) &&
                (c.Warnings?.Any(w => w.Contains("read error", StringComparison.OrdinalIgnoreCase)) ?? false));
        } finally {
            if (Directory.Exists(folder)) Directory.Delete(folder, recursive: true);
        }
    }

    [Fact]
    public void DocumentReader_ReadFolder_RespectsRecursionAndExtensionFilter() {
        var folder = Path.Combine(Path.GetTempPath(), "officeimo-reader-" + Guid.NewGuid().ToString("N"));
        var nested = Path.Combine(folder, "nested");
        Directory.CreateDirectory(folder);
        Directory.CreateDirectory(nested);

        var rootMarkdown = Path.Combine(folder, "root.md");
        var nestedMarkdown = Path.Combine(nested, "nested.md");
        var nestedText = Path.Combine(nested, "ignored.txt");

        try {
            File.WriteAllText(rootMarkdown, "# Root");
            File.WriteAllText(nestedMarkdown, "# Nested");
            File.WriteAllText(nestedText, "Ignore me");

            var noRecurse = DocumentReader.ReadFolder(
                folderPath: folder,
                folderOptions: new ReaderFolderOptions {
                    Recurse = false,
                    Extensions = new[] { ".md" },
                    DeterministicOrder = true
                }).ToList();

            Assert.Contains(noRecurse, c => string.Equals(c.Location.Path, rootMarkdown, StringComparison.OrdinalIgnoreCase));
            Assert.DoesNotContain(noRecurse, c => string.Equals(c.Location.Path, nestedMarkdown, StringComparison.OrdinalIgnoreCase));
            Assert.DoesNotContain(noRecurse, c => string.Equals(c.Location.Path, nestedText, StringComparison.OrdinalIgnoreCase));

            var recurse = DocumentReader.ReadFolder(
                folderPath: folder,
                folderOptions: new ReaderFolderOptions {
                    Recurse = true,
                    Extensions = new[] { ".md" },
                    DeterministicOrder = true
                }).ToList();

            Assert.Contains(recurse, c => string.Equals(c.Location.Path, rootMarkdown, StringComparison.OrdinalIgnoreCase));
            Assert.Contains(recurse, c => string.Equals(c.Location.Path, nestedMarkdown, StringComparison.OrdinalIgnoreCase));
            Assert.DoesNotContain(recurse, c => string.Equals(c.Location.Path, nestedText, StringComparison.OrdinalIgnoreCase));
        } finally {
            if (Directory.Exists(folder)) Directory.Delete(folder, recursive: true);
        }
    }

    [Fact]
    public void DocumentReader_ReadFolder_EmitsWarningWhenFileExceedsMaxInputBytes() {
        var folder = Path.Combine(Path.GetTempPath(), "officeimo-reader-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);

        var smallMarkdown = Path.Combine(folder, "small.md");
        var largeMarkdown = Path.Combine(folder, "large.md");

        try {
            File.WriteAllText(smallMarkdown, "# Small\n\nok");
            File.WriteAllText(largeMarkdown, new string('x', 1024));

            var chunks = DocumentReader.ReadFolder(
                folderPath: folder,
                folderOptions: new ReaderFolderOptions {
                    Recurse = false,
                    DeterministicOrder = true
                },
                options: new ReaderOptions {
                    MaxInputBytes = 128
                }).ToList();

            Assert.Contains(chunks, c => string.Equals(c.Location.Path, smallMarkdown, StringComparison.OrdinalIgnoreCase));
            Assert.Contains(chunks, c =>
                c.Kind == ReaderInputKind.Unknown &&
                string.Equals(c.Location.Path, largeMarkdown, StringComparison.OrdinalIgnoreCase) &&
                (c.Warnings?.Any(w => w.Contains("MaxInputBytes", StringComparison.OrdinalIgnoreCase)) ?? false));
        } finally {
            if (Directory.Exists(folder)) Directory.Delete(folder, recursive: true);
        }
    }

    [Fact]
    public void DocumentReader_ReadFolderDetailed_ReturnsSummaryAndFileStatuses() {
        var folder = Path.Combine(Path.GetTempPath(), "officeimo-reader-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);

        var goodMarkdown = Path.Combine(folder, "good.md");
        var badDocx = Path.Combine(folder, "broken.docx");

        try {
            File.WriteAllText(goodMarkdown, "# Good\n\nBody");
            File.WriteAllText(badDocx, "not-a-zip-package");

            var result = DocumentReader.ReadFolderDetailed(
                folderPath: folder,
                folderOptions: new ReaderFolderOptions {
                    Recurse = false,
                    DeterministicOrder = true
                },
                options: new ReaderOptions {
                    ComputeHashes = true
                },
                includeChunks: true);

            Assert.NotNull(result);
            Assert.True(result.FilesScanned >= 2);
            Assert.True(result.FilesParsed >= 1);
            Assert.True(result.FilesSkipped >= 1);
            Assert.NotEmpty(result.Files);
            Assert.NotEmpty(result.Chunks);

            var good = result.Files.FirstOrDefault(f => string.Equals(f.Path, goodMarkdown, StringComparison.OrdinalIgnoreCase));
            Assert.NotNull(good);
            Assert.True(good!.Parsed);
            Assert.False(string.IsNullOrWhiteSpace(good.SourceId));
            Assert.False(string.IsNullOrWhiteSpace(good.SourceHash));

            var bad = result.Files.FirstOrDefault(f => string.Equals(f.Path, badDocx, StringComparison.OrdinalIgnoreCase));
            Assert.NotNull(bad);
            Assert.False(bad!.Parsed);
            Assert.True((bad.Warnings?.Count ?? 0) > 0);
            Assert.Contains(bad.Warnings!, w => w.Contains("read error", StringComparison.OrdinalIgnoreCase));
        } finally {
            if (Directory.Exists(folder)) Directory.Delete(folder, recursive: true);
        }
    }

    [Fact]
    public void DocumentReader_ReadFolder_ProgressCallback_EmitsLifecycleEvents() {
        var folder = Path.Combine(Path.GetTempPath(), "officeimo-reader-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        var goodMarkdown = Path.Combine(folder, "good.md");
        var badDocx = Path.Combine(folder, "broken.docx");
        var events = new System.Collections.Generic.List<ReaderProgress>();

        try {
            File.WriteAllText(goodMarkdown, "# Good\n\nBody");
            File.WriteAllText(badDocx, "not-a-zip-package");

            var chunks = DocumentReader.ReadFolder(
                folderPath: folder,
                folderOptions: new ReaderFolderOptions {
                    Recurse = false,
                    DeterministicOrder = true
                },
                options: new ReaderOptions(),
                onProgress: p => events.Add(p)).ToList();

            Assert.NotEmpty(chunks);
            Assert.NotEmpty(events);
            Assert.Contains(events, e => e.Kind == ReaderProgressEventKind.FileStarted);
            Assert.Contains(events, e => e.Kind == ReaderProgressEventKind.FileCompleted);
            Assert.Contains(events, e => e.Kind == ReaderProgressEventKind.FileSkipped);
            Assert.Contains(events, e => e.Kind == ReaderProgressEventKind.Completed);

            var final = events.Last();
            Assert.Equal(ReaderProgressEventKind.Completed, final.Kind);
            Assert.True(final.FilesScanned >= 2);
            Assert.True(final.FilesParsed >= 1);
            Assert.True(final.FilesSkipped >= 1);
        } finally {
            if (Directory.Exists(folder)) Directory.Delete(folder, recursive: true);
        }
    }

    [Fact]
    public void DocumentReader_Read_EmitsSourceAndChunkMetadata() {
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".md");
        try {
            File.WriteAllText(path, "# Title\n\nBody");

            var chunks = DocumentReader.Read(path, new ReaderOptions { ComputeHashes = true }).ToList();
            Assert.NotEmpty(chunks);
            Assert.All(chunks, c => {
                Assert.False(string.IsNullOrWhiteSpace(c.SourceId));
                Assert.False(string.IsNullOrWhiteSpace(c.SourceHash));
                Assert.False(string.IsNullOrWhiteSpace(c.ChunkHash));
                Assert.True(c.TokenEstimate.HasValue && c.TokenEstimate.Value >= 1);
                Assert.True(c.SourceLengthBytes.HasValue && c.SourceLengthBytes.Value > 0);
            });
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void DocumentReader_ReadFolderDocuments_ReturnsPerFilePayloadsForDatabaseIngestion() {
        var folder = Path.Combine(Path.GetTempPath(), "officeimo-reader-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);

        var goodMarkdown = Path.Combine(folder, "good.md");
        var badDocx = Path.Combine(folder, "broken.docx");

        try {
            File.WriteAllText(goodMarkdown, "# Good\n\nBody");
            File.WriteAllText(badDocx, "not-a-zip-package");

            var docs = DocumentReader.ReadFolderDocuments(
                folderPath: folder,
                folderOptions: new ReaderFolderOptions {
                    Recurse = false,
                    DeterministicOrder = true
                },
                options: new ReaderOptions {
                    ComputeHashes = true
                }).ToList();

            Assert.NotEmpty(docs);
            Assert.True(docs.Count >= 2);

            var good = docs.FirstOrDefault(d => string.Equals(d.Path, goodMarkdown, StringComparison.OrdinalIgnoreCase));
            Assert.NotNull(good);
            Assert.True(good!.Parsed);
            Assert.False(string.IsNullOrWhiteSpace(good.SourceId));
            Assert.False(string.IsNullOrWhiteSpace(good.SourceHash));
            Assert.True(good.ChunksProduced > 0);
            Assert.True(good.TokenEstimateTotal > 0);
            Assert.NotEmpty(good.Chunks);

            var bad = docs.FirstOrDefault(d => string.Equals(d.Path, badDocx, StringComparison.OrdinalIgnoreCase));
            Assert.NotNull(bad);
            Assert.False(bad!.Parsed);
            Assert.Equal(0, bad.ChunksProduced);
            Assert.Equal(0, bad.TokenEstimateTotal);
            Assert.Empty(bad.Chunks);
            Assert.True((bad.Warnings?.Count ?? 0) > 0);
            Assert.Contains(bad.Warnings!, w => w.Contains("read error", StringComparison.OrdinalIgnoreCase));
        } finally {
            if (Directory.Exists(folder)) Directory.Delete(folder, recursive: true);
        }
    }
}
