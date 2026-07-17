using OfficeIMO.PowerPoint;
using OfficeIMO.Reader;
using System;
using System.IO;
using System.Linq;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Tests;

[Collection("ReaderRegistryNonParallel")]
public sealed class ReaderPowerPointExtractionTests {
    [Theory]
    [InlineData("pptx", false)]
    [InlineData("ppt", true)]
    [InlineData("pot", true)]
    [InlineData("pps", true)]
    public void RichExtractionHasPptxAndPptPathStreamParity(
        string extension, bool binary) {
        string path = Path.Combine(Path.GetTempPath(),
            Guid.NewGuid().ToString("N") + "." + extension);
        try {
            byte[] bytes = CreateRichPresentation(binary);
            File.WriteAllBytes(path, bytes);

            OfficeDocumentReadResult pathResult =
                OfficeDocumentReader.Default.ReadDocument(path);
            using var stream = new MemoryStream(bytes, writable: false);
            OfficeDocumentReadResult streamResult =
                OfficeDocumentReader.Default.ReadDocument(stream,
                    "reader-contract." + extension);
            ReaderDetectionResult pathDetection =
                OfficeDocumentReader.Default.Detect(path);
            ReaderDetectionResult streamDetection =
                OfficeDocumentReader.Default.Detect(bytes,
                    "reader-contract." + extension);

            AssertPowerPointRichExtraction(pathResult);
            AssertPowerPointRichExtraction(streamResult);
            Assert.Equal(pathResult.Markdown, streamResult.Markdown);
            Assert.Equal(ReaderInputKind.PowerPoint, pathDetection.Kind);
            Assert.Equal(ReaderInputKind.PowerPoint, streamDetection.Kind);
            Assert.Contains("officeimo.powerpoint.shape-model",
                pathResult.CapabilitiesUsed);
            Assert.Contains("officeimo.powerpoint.shape-model",
                streamResult.CapabilitiesUsed);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void ExtensionlessBinaryPowerPointUsesCompoundContentDetection() {
        byte[] bytes = CreateRichPresentation(binary: true);
        string path = Path.Combine(Path.GetTempPath(),
            Guid.NewGuid().ToString("N"));
        try {
            File.WriteAllBytes(path, bytes);
            ReaderDetectionResult byteDetection =
                OfficeDocumentReader.Default.Detect(bytes);
            ReaderDetectionResult pathDetection =
                OfficeDocumentReader.Default.Detect(path);
            using var stream = new MemoryStream(bytes, writable: false);
            OfficeDocumentReadResult streamResult =
                OfficeDocumentReader.Default.ReadDocument(stream);
            OfficeDocumentReadResult pathResult =
                OfficeDocumentReader.Default.ReadDocument(path);

            foreach (ReaderDetectionResult detection in new[] {
                         byteDetection, pathDetection
                     }) {
                Assert.Equal(ReaderInputKind.PowerPoint, detection.Kind);
                Assert.Equal(ReaderInputKind.PowerPoint,
                    detection.ContentKind);
                Assert.Equal("application/vnd.ms-powerpoint",
                    detection.MediaType);
                Assert.Contains("container:ole-powerpoint-presentation",
                    detection.Evidence);
            }
            AssertPowerPointRichExtraction(streamResult);
            AssertPowerPointRichExtraction(pathResult);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void RichExtractionRecursesThroughGroupsInDrawingOrder() {
        byte[] bytes;
        using (PowerPointPresentation presentation =
               PowerPointPresentation.Create()) {
            PowerPointSlide slide = presentation.AddSlide();
            slide.AddTextBox("Before group");
            PowerPointTextBox groupedText =
                slide.AddTextBox("Grouped paragraph");
            PowerPointTable groupedTable = slide.AddTable(2, 1);
            groupedTable.GetCell(0, 0).Text = "Header";
            groupedTable.GetCell(1, 0).Text = "Grouped cell";
            slide.GroupShapes(new PowerPointShape[] {
                groupedText,
                groupedTable
            }, "Content group");
            slide.AddTextBox("After group");
            bytes = presentation.ToBytes();
        }

        OfficeDocumentReadResult result =
            OfficeDocumentReader.Default.ReadDocument(
                new MemoryStream(bytes, writable: false),
                "grouped-content.pptx");

        int before = result.Blocks.ToList().FindIndex(block =>
            block.Text == "Before group");
        int groupedTextIndex = result.Blocks.ToList().FindIndex(block =>
            block.Text == "Grouped paragraph");
        int groupedTableIndex = result.Blocks.ToList().FindIndex(block =>
            block.Kind == "table"
            && block.Text.Contains("Grouped cell", StringComparison.Ordinal));
        int after = result.Blocks.ToList().FindIndex(block =>
            block.Text == "After group");
        Assert.True(before < groupedTextIndex
            && groupedTextIndex < groupedTableIndex
            && groupedTableIndex < after);
        Assert.Contains("Grouped paragraph", result.Markdown ?? string.Empty,
            StringComparison.Ordinal);
        Assert.Contains("Grouped cell", result.Markdown ?? string.Empty,
            StringComparison.Ordinal);
        Assert.Contains(result.Tables, table => table.Rows
            .SelectMany(static row => row)
            .Contains("Grouped cell", StringComparer.Ordinal));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void InternalRunHyperlinksBecomeMarkdownSlideLinks(bool binary) {
        byte[] bytes;
        using (PowerPointPresentation presentation =
               PowerPointPresentation.Create()) {
            PowerPointSlide source = presentation.AddSlide();
            PowerPointSlide target = presentation.AddSlide();
            PowerPointTextBox link = source.AddTextBox("Open ");
            link.Paragraphs[0].AddRun("destination")
                .SetHyperlink(target, "Jump to destination");
            target.AddTitle("Destination");
            bytes = binary
                ? presentation.ToBytes(PowerPointFileFormat.Ppt)
                : presentation.ToBytes();
        }

        OfficeDocumentReadResult result =
            OfficeDocumentReader.Default.ReadDocument(
                new MemoryStream(bytes, writable: false),
                binary ? "internal-link.ppt" : "internal-link.pptx");

        Assert.Contains("[destination](<#slide-2>)",
            result.Markdown ?? string.Empty, StringComparison.Ordinal);
        Assert.Contains(result.Links, link => link.Uri == "#slide-2"
            && link.Text == "destination");
    }

    [Fact]
    public void TableContractKeepsPptxAndPptSemanticParity() {
        byte[] pptxBytes;
        byte[] pptBytes;
        using (PowerPointPresentation presentation =
               PowerPointPresentation.Create()) {
            PowerPointSlide slide = presentation.AddSlide();
            slide.AddTextBox("Before table");
            PowerPointTable table = slide.AddTable(2, 2);
            table.GetCell(0, 0).Text = "Region";
            table.GetCell(0, 1).Text = "Revenue";
            table.GetCell(1, 0).Text = "North";
            table.GetCell(1, 1).Text = "120";
            slide.AddTextBox("After table");

            Assert.True(presentation.AnalyzeLegacyPptWrite().CanWrite);
            pptxBytes = presentation.ToBytes();
            pptBytes = presentation.ToBytes(PowerPointFileFormat.Ppt);
        }

        OfficeDocumentReadResult pptxResult =
            OfficeDocumentReader.Default.ReadDocument(
                new MemoryStream(pptxBytes, writable: false),
                "reader-table.pptx");
        OfficeDocumentReadResult pptResult =
            OfficeDocumentReader.Default.ReadDocument(
                new MemoryStream(pptBytes, writable: false),
                "reader-table.ppt");

        foreach (OfficeDocumentReadResult result in new[] {
                     pptxResult, pptResult
                 }) {
            ReaderTable semanticTable = Assert.Single(result.Tables);
            Assert.Contains("Region", semanticTable.Columns.Concat(
                semanticTable.Rows.SelectMany(static row => row)),
                StringComparer.Ordinal);
            Assert.Contains("120", semanticTable.Columns.Concat(
                semanticTable.Rows.SelectMany(static row => row)),
                StringComparer.Ordinal);
        }

        int pptxBefore = pptxResult.Blocks.ToList().FindIndex(block =>
            block.Text == "Before table");
        int pptxTable = pptxResult.Blocks.ToList().FindIndex(block =>
            block.Kind == "table");
        int pptxAfter = pptxResult.Blocks.ToList().FindIndex(block =>
            block.Text == "After table");
        Assert.True(pptxBefore < pptxTable && pptxTable < pptxAfter);

        int pptBefore = pptResult.Blocks.ToList().FindIndex(block =>
            block.Text == "Before table");
        int pptTable = pptResult.Blocks.ToList().FindIndex(block =>
            block.Kind == "table");
        int pptAfter = pptResult.Blocks.ToList().FindIndex(block =>
            block.Text == "After table");
        Assert.True(pptBefore < pptTable && pptTable < pptAfter);
    }

    private static byte[] CreateRichPresentation(bool binary) {
        using PowerPointPresentation presentation =
            PowerPointPresentation.Create();
        PowerPointSlide slide = presentation.AddSlide();
        slide.AddTitle("Reader extraction contract");
        PowerPointTextBox list = slide.AddTextBox(string.Empty);
        list.SetBullets(new[] { "Parent bullet" });
        list.AddBullets(new[] { "Nested bullet" }, level: 1);
        PowerPointTextBox numbered = slide.AddTextBox(string.Empty);
        numbered.SetNumberedList(new[] { "Third item", "Fourth item" },
            A.TextAutoNumberSchemeValues.AlphaLowerCharacterParenR,
            startAt: 3);
        PowerPointTextBox link = slide.AddTextBox("Read ");
        link.Paragraphs[0].AddRun("the guide")
            .SetHyperlink("https://example.test/guide");
        slide.Notes.Text = "Speaker reminder";
        var preflight = presentation.AnalyzeLegacyPptWrite();
        Assert.True(preflight.CanWrite,
            string.Join(Environment.NewLine, preflight.Findings));
        return binary
            ? presentation.ToBytes(PowerPointFileFormat.Ppt)
            : presentation.ToBytes();
    }

    private static void AssertPowerPointRichExtraction(
        OfficeDocumentReadResult result) {
        Assert.Equal(ReaderInputKind.PowerPoint, result.Kind);
        Assert.Contains("### Reader extraction contract",
            result.Markdown ?? string.Empty, StringComparison.Ordinal);
        Assert.Contains("- • Parent bullet", result.Markdown ?? string.Empty,
            StringComparison.Ordinal);
        Assert.Contains("    - • Nested bullet",
            result.Markdown ?? string.Empty, StringComparison.Ordinal);
        Assert.Contains("3. c) Third item", result.Markdown ?? string.Empty,
            StringComparison.Ordinal);
        Assert.Contains("4. d) Fourth item", result.Markdown ?? string.Empty,
            StringComparison.Ordinal);
        Assert.Contains("Read [the guide](<https://example.test/guide>)",
            result.Markdown ?? string.Empty, StringComparison.Ordinal);
        Assert.Contains("Speaker reminder", result.Markdown ?? string.Empty,
            StringComparison.Ordinal);

        OfficeDocumentBlock heading = Assert.Single(result.Blocks,
            block => block.Kind == "heading"
                && block.Text == "Reader extraction contract");
        Assert.Equal(1, heading.Level);
        OfficeDocumentBlock nested = Assert.Single(result.Blocks,
            block => block.Kind == "list-item"
                && block.Text == "Nested bullet");
        Assert.Equal(1, nested.Level);
        Assert.NotNull(nested.Marker);
        OfficeDocumentBlock numbered = Assert.Single(result.Blocks,
            block => block.Kind == "list-item"
                && block.Text == "Third item");
        Assert.Equal("c)", numbered.Marker);
        Assert.Contains(result.Links, link =>
            link.Uri == "https://example.test/guide"
            && link.Text == "the guide");
        Assert.Contains(result.Blocks, block =>
            block.Kind == "speaker-notes"
            && block.Text.Contains("Speaker reminder",
                StringComparison.Ordinal));
    }
}
