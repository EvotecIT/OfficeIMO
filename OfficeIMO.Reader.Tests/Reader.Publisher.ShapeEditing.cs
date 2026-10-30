using OfficeIMO.Drawing;
using OfficeIMO.Publisher;
using OfficeIMO.Reader.Publisher;
using OfficeIMO.Tests;
using Xunit;

namespace OfficeIMO.Reader.Tests;

public sealed class ReaderPublisherShapeEditingTests {
    [Fact]
    public void InMemoryPublisherProjectionRetainsSemanticContentWhileVectorArtworkChanges() {
        string path = Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.Publisher.Tests", "Fixtures",
            "NativeComparison", "Rectangle", "native-rectangle.pub");
        byte[] input = File.ReadAllBytes(path);
        PublisherDocument document = PublisherDocument.Load(input);
        PublisherShape original = Assert.Single(document.Pages[0].Shapes, shape => shape.CanEdit);
        var sourceReport = document.SourceReadReport;
        var options = new ReaderOptions { ComputeHashes = true };
        OfficeDocumentReadResult before = document.ToOfficeDocumentReadResult("shape-edited.pub", options);
        document.SetShapeBounds(original.Id, 90, 144, 180, 90);
        document.SetShapePaint(original.Id, OfficeColor.FromRgb(0, 90, 170), OfficeColor.Black, 2);
        var inputShape = OfficeShape.Rectangle(60, 40);
        inputShape.FillColor = OfficeColor.FromRgb(0, 128, 0);
        PublisherShape created = document.AddShape(document.Pages[0].Id, inputShape, 306, 144);
        inputShape.FillColor = OfficeColor.Black;
        Assert.NotEqual(original.Id, created.Id);
        Assert.Equal(2, document.Pages[0].Shapes.Count);
        Assert.Equal(document.Pages[0].Id, created.PageId);
        Assert.Equal(72, original.X, 5);
        Assert.Equal(90, Assert.Single(document.Pages[0].Shapes, shape => shape.Id == original.Id).X);
        Assert.Equal(OfficeColor.FromRgb(0, 128, 0), created.CreateShape().FillColor);

        OfficeDocumentReadResult after = document.ToOfficeDocumentReadResult("shape-edited.pub", options);
        Assert.Equal(before.Blocks.Select(block => (block.Id, block.Text)), after.Blocks.Select(block => (block.Id, block.Text)));
        Assert.Equal(before.Chunks.Select(chunk => (chunk.Id, chunk.Text, chunk.ChunkHash)),
            after.Chunks.Select(chunk => (chunk.Id, chunk.Text, chunk.ChunkHash)));
        Assert.Equal(before.Markdown, after.Markdown);
        Assert.Equal(before.Pages.Count, after.Pages.Count);
        Assert.Empty(after.Assets);
        Assert.Empty(after.Tables);
        Assert.Contains(after.Diagnostics, diagnostic => diagnostic.Code == "PUB_READER_LAYOUT_OMITTED");
        foreach (var finding in document.ReadReport.FidelityDiagnostics)
            Assert.Contains(after.Diagnostics, diagnostic => diagnostic.Code == finding.Code && diagnostic.Message == finding.Message);
        Assert.Null(after.Source.SourceHash);
        Assert.Null(after.Source.LengthBytes);
        Assert.Equal(OfficeDocumentPageProvenance.Native, after.GetPageProvenance());
        Assert.Same(sourceReport, document.SourceReadReport);
        Assert.Equal(input, File.ReadAllBytes(path));
    }
}
