using OfficeIMO.Excel;
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
        } finally {
            if (Directory.Exists(folder)) Directory.Delete(folder, recursive: true);
        }
    }
}
