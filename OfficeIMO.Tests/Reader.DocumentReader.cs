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
}
