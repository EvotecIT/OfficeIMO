using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;
using OfficeIMO.Provenance;

namespace OfficeIMO.Workflows.Tests;

public sealed partial class OfficeProvenanceWorkflowTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void BufferWorkflowBudgetsFinalReinspectionWithoutMutatingCallerLimits(bool embeddedCarrier) {
        using var scope = new TempScope();
        string path = Path.Combine(scope.Path, "budget.docx");
        using (var document = WordDocument.Create(path)) {
            document.AddParagraph("Budgeted document");
            if (embeddedCarrier) document.AddParagraph().AddImage(Path.Combine(AppContext.BaseDirectory, "samples", "provenance-demo.png"));
            document.Save();
        }
        byte[] bytes = File.ReadAllBytes(path);
        var ownerResult = WordDocument.RemoveProvenance(bytes, "budget.docx");
        // These small packages contain XML and a PNG, so inspection expands every entry once.
        static long ExpandedBytes(byte[] package) {
            using var stream = new MemoryStream(package);
            using var archive = new System.IO.Compression.ZipArchive(stream);
            return archive.Entries.Sum(entry => entry.Length);
        }
        long reopenBytes = ExpandedBytes(ownerResult.ToArray());
        long ownerBytes = ExpandedBytes(bytes) + reopenBytes;
        Assert.True(ownerBytes > 0);
        Assert.True(reopenBytes > 0);
        var limits = new OfficeProvenanceRemovalOptions { Limits = { MaxExpandedContainerBytes = ownerBytes + reopenBytes - 1 } };
        Assert.Throws<InvalidDataException>(() => OfficeProvenanceBufferWorkflow.Remove(bytes, "budget.docx", limits));
        Assert.Equal(ownerBytes + reopenBytes - 1, limits.Limits.MaxExpandedContainerBytes);
        limits.Limits.MaxExpandedContainerBytes++;
        var result = OfficeProvenanceBufferWorkflow.Remove(bytes, "budget.docx", limits);
        Assert.Equal(embeddedCarrier, result.WasChanged);
        Assert.Equal(ownerResult.ToArray(), result.ToArray());
        Assert.Equal(ownerBytes + reopenBytes, limits.Limits.MaxExpandedContainerBytes);
        Assert.Equal(File.ReadAllBytes(path), bytes);
    }

    [Fact]
    public void BufferWorkflowUsesPackageOwnersAndPreservesUnchangedDocuments() {
        using var scope = new TempScope();
        string word = Path.Combine(scope.Path, "report.docx");
        using (var document = WordDocument.Create(word)) {
            document.AddParagraph("Preserved content"); document.Save();
        }
        string excel = Path.Combine(scope.Path, "report.xlsx");
        using (var document = ExcelDocument.Create(excel)) {
            document.AddWorksheet("Data").Cell(1, 1, "Preserved content"); document.Save();
        }
        string deck = Path.Combine(scope.Path, "report.pptx");
        using (var presentation = PowerPointPresentation.Create(deck)) {
            presentation.AddSlide().AddTextBoxPoints("Preserved content", 20, 20, 300, 60); presentation.Save();
        }
        foreach (string path in new[] { word, excel, deck }) {
            byte[] bytes = File.ReadAllBytes(path);
            var report = OfficeProvenanceBufferWorkflow.Inspect(bytes, Path.GetFileName(path));
            Assert.Empty(report.Evidence);
            var result = OfficeProvenanceBufferWorkflow.Remove(bytes, Path.GetFileName(path));
            Assert.False(result.WasChanged);
            Assert.Equal(bytes, result.ToArray());
            Assert.Equal(report.Format, result.After.Format);
        }
        Assert.Throws<InvalidDataException>(() => OfficeProvenanceBufferWorkflow.Inspect(File.ReadAllBytes(word), "renamed.xlsx"));
    }
    [Fact]
    public void BufferWorkflowRetainsTheOwnersEffectiveOutputAndEmbeddedLimits() {
        string sample = Path.Combine(AppContext.BaseDirectory, "samples", "provenance-demo.png");
        var small = OfficeProvenanceBufferWorkflow.Remove(File.ReadAllBytes(sample), "sample.png", new OfficeProvenanceRemovalOptions { MaxOutputBytes = 1_000_000 });
        Assert.True(small.WasChanged); Assert.Empty(small.After.Evidence);
        using var scope = new TempScope();
        string path = Path.Combine(scope.Path, "embedded.docx");
        using (var document = WordDocument.Create(path)) {
            document.AddParagraph().AddImage(sample); document.Save();
        }
        byte[] original = File.ReadAllBytes(path);
        Assert.NotEmpty(OfficeProvenanceBufferWorkflow.Inspect(original, "embedded.docx").Evidence);
        var disabled = new OfficeProvenanceRemovalOptions { Limits = { ProcessEmbeddedAssets = false } };
        var unchanged = OfficeProvenanceBufferWorkflow.Remove(original, "embedded.docx", disabled);
        Assert.False(unchanged.WasChanged); Assert.Equal(original, unchanged.ToArray());
        var bounded = new OfficeProvenanceRemovalOptions { Limits = { MaxEmbeddedAssets = 1 }, MaxEmbeddedAssets = 20 };
        Assert.True(OfficeProvenanceBufferWorkflow.Remove(original, "embedded.docx", bounded).WasChanged);
    }

}
