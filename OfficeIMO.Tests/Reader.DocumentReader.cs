using OfficeIMO.Excel;
using OfficeIMO.Markdown;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.Reader;
using OfficeIMO.Word;
using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
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
            var markdownChunks = chunks.Where(static c => c.Kind == ReaderInputKind.Markdown).ToList();
            Assert.Collection(
                markdownChunks.Select(c => c.Location.StartLine),
                line => Assert.Equal(1, line),
                line => Assert.Equal(5, line));
            Assert.Collection(
                markdownChunks.Select(c => c.Location.EndLine),
                line => Assert.Equal(3, line),
                line => Assert.Equal(7, line));
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void DocumentReader_MarkdownChunking_RecognizesSetextHeadings() {
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".md");
        try {
            File.WriteAllText(path, "Top\r\n===\r\n\r\nPara 1.\r\n\r\nChild\r\n---\r\n\r\nPara 2.\r\n");

            var chunks = DocumentReader.Read(path).Where(static c => c.Kind == ReaderInputKind.Markdown).ToList();

            Assert.True(chunks.Count >= 2);
            Assert.Contains(chunks, c => string.Equals(c.Location.HeadingPath, "Top", StringComparison.Ordinal));
            Assert.Contains(chunks, c => string.Equals(c.Location.HeadingPath, "Top > Child", StringComparison.Ordinal));
            Assert.Contains(chunks, c => (c.Markdown ?? string.Empty).Contains("# Top", StringComparison.Ordinal));
            Assert.Contains(chunks, c => (c.Markdown ?? string.Empty).Contains("## Child", StringComparison.Ordinal));
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void DocumentReader_MarkdownChunking_DoesNotTreatCodeFenceContentAsHeading() {
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".md");
        try {
            File.WriteAllText(path,
                "# Top\n\n```text\n# not a heading\nline 2\n```\n\nAfter fence.\n\n## Child\n\nDone.\n");

            var chunks = DocumentReader.Read(path).Where(static c => c.Kind == ReaderInputKind.Markdown).ToList();

            Assert.Equal(2, chunks.Count);
            Assert.Equal("Top", chunks[0].Location.HeadingPath);
            Assert.Contains("```text", chunks[0].Markdown ?? string.Empty, StringComparison.Ordinal);
            Assert.Contains("# not a heading", chunks[0].Markdown ?? string.Empty, StringComparison.Ordinal);
            Assert.Equal("Top > Child", chunks[1].Location.HeadingPath);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void DocumentReader_MarkdownChunking_PreservesWholeBlocksWhenTheyExceedMaxChars() {
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".md");
        try {
            var largePayload = new string('x', 320);
            File.WriteAllText(path,
                "# Top\n\nIntro paragraph.\n\n```text\n" + largePayload + "\n```\n\nTail paragraph.\n");

            var chunks = DocumentReader.Read(path, new ReaderOptions { MaxChars = 256 }).Where(static c => c.Kind == ReaderInputKind.Markdown).ToList();

            Assert.True(chunks.Count >= 3);
            Assert.Contains(chunks, c => (c.Warnings?.Any(w => w.Contains("single markdown block exceeded MaxChars", StringComparison.OrdinalIgnoreCase)) ?? false));

            var codeChunk = chunks.Single(c => (c.Markdown ?? string.Empty).Contains("```text", StringComparison.Ordinal));
            Assert.Contains(largePayload, codeChunk.Markdown ?? string.Empty, StringComparison.Ordinal);
            Assert.Contains("```", codeChunk.Markdown ?? string.Empty, StringComparison.Ordinal);
            Assert.Equal("Top", codeChunk.Location.HeadingPath);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void DocumentReader_MarkdownChunking_ExtractsMarkdownTables() {
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".md");
        try {
            File.WriteAllText(path,
                "# Data\n\n| Name | Value |\n| --- | ---: |\n| A | 1 |\n| B | 2 |\n");

            var chunk = DocumentReader.Read(path).Single(c => c.Kind == ReaderInputKind.Markdown && (c.Tables?.Count ?? 0) > 0);

            Assert.Equal("Data", chunk.Location.HeadingPath);
            Assert.NotNull(chunk.Tables);
            Assert.Single(chunk.Tables!);
            Assert.Equal(new[] { "Name", "Value" }, chunk.Tables![0].Columns);
            Assert.Equal(2, chunk.Tables[0].TotalRowCount);
            Assert.Equal("A", chunk.Tables[0].Rows[0][0]);
            Assert.Equal("2", chunk.Tables[0].Rows[1][1]);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void DocumentReader_MarkdownChunking_RespectsTableRowCapsAndFallbackHeaders() {
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".md");
        try {
            File.WriteAllText(path,
                "# Data\n\n| A | 1 |\n| B | 2 |\n| C | 3 |\n");

            var chunk = DocumentReader.Read(path, new ReaderOptions { MaxTableRows = 2 })
                .Single(c => c.Kind == ReaderInputKind.Markdown && (c.Tables?.Count ?? 0) > 0);

            Assert.NotNull(chunk.Tables);
            Assert.Single(chunk.Tables!);
            Assert.Equal(new[] { "Column1", "Column2" }, chunk.Tables![0].Columns);
            Assert.Equal(3, chunk.Tables[0].TotalRowCount);
            Assert.True(chunk.Tables[0].Truncated);
            Assert.Equal(2, chunk.Tables[0].Rows.Count);
            Assert.Equal("A", chunk.Tables[0].Rows[0][0]);
            Assert.Equal("2", chunk.Tables[0].Rows[1][1]);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void DocumentReader_MarkdownChunking_ExtractsIxDataViewTables() {
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".md");
        var raw = "{\"title\":\"Replication Summary\",\"summary\":\"Latest replication posture\",\"kind\":\"ix_tool_dataview_v1\",\"call_id\":\"call_123\",\"rows\":[[\"Server\",\"Fails\"],[\"AD0\",\"0\"],[\"AD1\",\"1\"]]}";
        try {
            File.WriteAllText(path,
                "# Visual\n\n```ix-dataview\n" + raw + "\n```\n");

            var chunk = DocumentReader.Read(path).Single(c => c.Kind == ReaderInputKind.Markdown && (c.Tables?.Count ?? 0) > 0);

            Assert.Equal("Visual", chunk.Location.HeadingPath);
            Assert.NotNull(chunk.Tables);
            Assert.Single(chunk.Tables!);
            Assert.Equal("Replication Summary", chunk.Tables![0].Title);
            Assert.Equal("ix_tool_dataview_v1", chunk.Tables[0].Kind);
            Assert.Equal("call_123", chunk.Tables[0].CallId);
            Assert.Equal("Latest replication posture", chunk.Tables[0].Summary);
            Assert.Equal(ComputeShortHash(raw), chunk.Tables[0].PayloadHash);
            Assert.Equal(new[] { "Server", "Fails" }, chunk.Tables[0].Columns);
            Assert.Equal(2, chunk.Tables[0].TotalRowCount);
            Assert.Equal("AD0", chunk.Tables[0].Rows[0][0]);
            Assert.Equal("1", chunk.Tables[0].Rows[1][1]);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    private static string ComputeShortHash(string input) {
        var data = Encoding.UTF8.GetBytes(input ?? string.Empty);
        byte[] hash;
        using (var sha = SHA256.Create()) {
            hash = sha.ComputeHash(data);
        }

        var sb = new StringBuilder(16);
        for (int i = 0; i < 8 && i < hash.Length; i++) {
            sb.Append(hash[i].ToString("x2", CultureInfo.InvariantCulture));
        }

        return sb.ToString();
    }

    [Fact]
    public void DocumentReader_MarkdownChunking_ExtractsIxDataViewColumnsAndObjectRecords() {
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".md");
        try {
            File.WriteAllText(path,
                "# Visual\n\n```ix-dataview\n{\"kind\":\"ix_tool_dataview_v1\",\"columns\":[\"Server\",\"Fails\"],\"records\":[{\"Server\":\"AD0\",\"Fails\":0},{\"Server\":\"AD1\",\"Fails\":1}]}\n```\n");

            var chunk = DocumentReader.Read(path).Single(c => c.Kind == ReaderInputKind.Markdown && (c.Tables?.Count ?? 0) > 0);

            Assert.NotNull(chunk.Tables);
            Assert.Single(chunk.Tables!);
            Assert.Equal("ix_tool_dataview_v1", chunk.Tables![0].Title);
            Assert.Equal(new[] { "Server", "Fails" }, chunk.Tables[0].Columns);
            Assert.Equal(2, chunk.Tables[0].TotalRowCount);
            Assert.Equal("AD0", chunk.Tables[0].Rows[0][0]);
            Assert.Equal("1", chunk.Tables[0].Rows[1][1]);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void DocumentReader_MarkdownChunking_RespectsIxDataViewRowCaps() {
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".md");
        try {
            File.WriteAllText(path,
                "# Visual\n\n```ix-dataview\n{\"rows\":[[\"Server\",\"Fails\"],[\"AD0\",\"0\"],[\"AD1\",\"1\"],[\"AD2\",\"2\"]]}\n```\n");

            var chunk = DocumentReader.Read(path, new ReaderOptions { MaxTableRows = 2 })
                .Single(c => c.Kind == ReaderInputKind.Markdown && (c.Tables?.Count ?? 0) > 0);

            Assert.NotNull(chunk.Tables);
            Assert.Single(chunk.Tables!);
            Assert.Equal(new[] { "Server", "Fails" }, chunk.Tables![0].Columns);
            Assert.Equal(3, chunk.Tables[0].TotalRowCount);
            Assert.True(chunk.Tables[0].Truncated);
            Assert.Equal(2, chunk.Tables[0].Rows.Count);
            Assert.Equal("AD0", chunk.Tables[0].Rows[0][0]);
            Assert.Equal("1", chunk.Tables[0].Rows[1][1]);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void DocumentReader_MarkdownChunking_CanNormalize_CompactIxDataViewFenceBodies() {
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".md");
        try {
            File.WriteAllText(path,
                "# Visual\n\n```ix-dataview{\"kind\":\"ix_tool_dataview_v1\",\"rows\":[[\"Server\",\"Fails\"],[\"AD0\",\"0\"]]}\n```\n");

            var chunk = DocumentReader.Read(path, new ReaderOptions {
                MarkdownInputNormalization = new MarkdownInputNormalizationOptions {
                    NormalizeCompactFenceBodyBoundaries = true
                }
            }).Single(c => c.Kind == ReaderInputKind.Markdown && (c.Tables?.Count ?? 0) > 0);

            Assert.NotNull(chunk.Tables);
            Assert.Single(chunk.Tables!);
            Assert.Equal("ix_tool_dataview_v1", chunk.Tables![0].Title);
            Assert.Equal(new[] { "Server", "Fails" }, chunk.Tables[0].Columns);
            Assert.Equal("AD0", chunk.Tables[0].Rows[0][0]);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void DocumentReader_MarkdownChunking_EmitsLineRangesAndBlockKinds() {
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".md");
        try {
            File.WriteAllText(path,
                "# Top\n\nPara 1.\n\n## Child\n\nPara 2.\n");

            var chunks = DocumentReader.Read(path).Where(static c => c.Kind == ReaderInputKind.Markdown).ToList();

            var topChunk = chunks.Single(c => string.Equals(c.Location.HeadingPath, "Top", StringComparison.Ordinal));
            Assert.Equal(1, topChunk.Location.StartLine);
            Assert.Equal(3, topChunk.Location.EndLine);
            Assert.Equal(1, topChunk.Location.NormalizedStartLine);
            Assert.Equal(3, topChunk.Location.NormalizedEndLine);
            Assert.Equal("top", topChunk.Location.HeadingSlug);
            Assert.Equal("heading", topChunk.Location.SourceBlockKind);
            Assert.Equal("top", topChunk.Location.BlockAnchor);
            Assert.Equal(0, topChunk.Location.SourceBlockIndex);

            var childChunk = chunks.Single(c => string.Equals(c.Location.HeadingPath, "Top > Child", StringComparison.Ordinal));
            Assert.Equal(5, childChunk.Location.StartLine);
            Assert.Equal(7, childChunk.Location.EndLine);
            Assert.Equal(5, childChunk.Location.NormalizedStartLine);
            Assert.Equal(7, childChunk.Location.NormalizedEndLine);
            Assert.Equal("child", childChunk.Location.HeadingSlug);
            Assert.Equal("heading", childChunk.Location.SourceBlockKind);
            Assert.Equal("child", childChunk.Location.BlockAnchor);
            Assert.Equal(2, childChunk.Location.SourceBlockIndex);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void DocumentReader_MarkdownChunking_CanApply_InputNormalization() {
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".md");
        try {
            File.WriteAllText(path, "## Wynik ogólny- **Replication:** wcześniej zdrowa ✅- **FSMO:** technicznie OK");

            var chunks = DocumentReader.Read(path, new ReaderOptions {
                MarkdownInputNormalization = new MarkdownInputNormalizationOptions {
                    NormalizeHeadingListBoundaries = true,
                    NormalizeCompactStrongLabelListBoundaries = true
                }
            }).Where(static c => c.Kind == ReaderInputKind.Markdown).ToList();

            var chunk = Assert.Single(chunks);
            Assert.Equal("Wynik ogólny", chunk.Location.HeadingPath);
            Assert.Contains("## Wynik ogólny", chunk.Markdown ?? string.Empty, StringComparison.Ordinal);
            Assert.Contains("- **Replication:**", chunk.Markdown ?? string.Empty, StringComparison.Ordinal);
            Assert.Contains("- **FSMO:**", chunk.Markdown ?? string.Empty, StringComparison.Ordinal);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void DocumentReader_MarkdownChunking_CanApply_BlockBoundaryNormalization() {
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".md");
        try {
            File.WriteAllText(path, "previous shutdown was unexpected### Reason- **Unplanned / unexpected reboot**");

            var chunks = DocumentReader.Read(path, new ReaderOptions {
                MarkdownInputNormalization = new MarkdownInputNormalizationOptions {
                    NormalizeCompactHeadingBoundaries = true,
                    NormalizeHeadingListBoundaries = true,
                    NormalizeCompactStrongLabelListBoundaries = true
                }
            }).Where(static c => c.Kind == ReaderInputKind.Markdown).ToList();

            Assert.Equal(2, chunks.Count);
            Assert.Equal("Reason", chunks[1].Location.HeadingPath);
            Assert.Contains("### Reason", chunks[1].Markdown ?? string.Empty, StringComparison.Ordinal);
            Assert.Contains("- **Unplanned / unexpected reboot**", chunks[1].Markdown ?? string.Empty, StringComparison.Ordinal);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void DocumentReader_MarkdownChunking_AssignsSubBlockAnchorsWhenChunksSplitWithinAHeading() {
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".md");
        try {
            var largePayload = new string('x', 320);
            File.WriteAllText(path,
                "# Top\n\nIntro paragraph.\n\n```text\n" + largePayload + "\n```\n\nTail paragraph.\n");

            var chunks = DocumentReader.Read(path, new ReaderOptions { MaxChars = 256 })
                .Where(static c => c.Kind == ReaderInputKind.Markdown)
                .ToList();

            Assert.True(chunks.Count >= 3);

            var headingChunk = chunks.Single(c => string.Equals(c.Location.BlockAnchor, "top", StringComparison.Ordinal));
            Assert.Contains("# Top", headingChunk.Markdown ?? string.Empty, StringComparison.Ordinal);

            var codeChunk = chunks.Single(c => string.Equals(c.Location.BlockAnchor, "top--code-2", StringComparison.Ordinal));
            Assert.Equal("code", codeChunk.Location.SourceBlockKind);
            Assert.Equal("top", codeChunk.Location.HeadingSlug);
            Assert.Contains("```text", codeChunk.Markdown ?? string.Empty, StringComparison.Ordinal);

            var trailingParagraphChunk = chunks.Single(c => string.Equals(c.Location.BlockAnchor, "top--paragraph-3", StringComparison.Ordinal));
            Assert.Equal("paragraph", trailingParagraphChunk.Location.SourceBlockKind);
            Assert.Equal("top", trailingParagraphChunk.Location.HeadingSlug);
            Assert.Contains("Tail paragraph.", trailingParagraphChunk.Markdown ?? string.Empty, StringComparison.Ordinal);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void DocumentReader_MarkdownChunking_AssignsUniqueHeadingSlugsForDuplicateHeadings() {
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".md");
        try {
            File.WriteAllText(path,
                "# Repeat\n\nOne.\n\n## Child\n\nA.\n\n# Repeat\n\nTwo.\n");

            var chunks = DocumentReader.Read(path).Where(static c => c.Kind == ReaderInputKind.Markdown).ToList();

            Assert.Contains(chunks, c => string.Equals(c.Location.HeadingPath, "Repeat", StringComparison.Ordinal) && string.Equals(c.Location.HeadingSlug, "repeat", StringComparison.Ordinal));
            Assert.Contains(chunks, c => string.Equals(c.Location.HeadingPath, "Repeat > Child", StringComparison.Ordinal) && string.Equals(c.Location.HeadingSlug, "child", StringComparison.Ordinal));
            Assert.Contains(chunks, c => string.Equals(c.Location.HeadingPath, "Repeat", StringComparison.Ordinal) && string.Equals(c.Location.HeadingSlug, "repeat-1", StringComparison.Ordinal));
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void DocumentReader_MarkdownChunking_AssignsDeterministicSlugsForNonAsciiHeadings() {
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".md");
        try {
            File.WriteAllText(path,
                "# !!!\n\nOne.\n\n# !!!\n\nTwo.\n\n# ążźć\n\nThree.\n");

            var chunks = DocumentReader.Read(path).Where(static c => c.Kind == ReaderInputKind.Markdown).ToList();

            var punctuationChunks = chunks.Where(c => string.Equals(c.Location.HeadingPath, "!!!", StringComparison.Ordinal)).ToList();
            Assert.Equal(2, punctuationChunks.Count);
            Assert.All(punctuationChunks, c => Assert.StartsWith("heading-", c.Location.HeadingSlug ?? string.Empty, StringComparison.Ordinal));
            Assert.NotEqual(punctuationChunks[0].Location.HeadingSlug, punctuationChunks[1].Location.HeadingSlug);

            var nonAsciiChunk = chunks.Single(c => string.Equals(c.Location.HeadingPath, "ążźć", StringComparison.Ordinal));
            Assert.StartsWith("heading-", nonAsciiChunk.Location.HeadingSlug ?? string.Empty, StringComparison.Ordinal);
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
