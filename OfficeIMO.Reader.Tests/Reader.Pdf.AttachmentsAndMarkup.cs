using OfficeIMO.Pdf;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

[Collection("ReaderRegistryNonParallel")]
public sealed class ReaderPdfAttachmentsAndMarkupTests {
    [Fact]
    public void ReadDocument_ReadsEmbeddedPdfAttachmentAsNestedDocumentWithItsContainerFacts() {
        byte[] inner = CreatePdf("Embedded quarterly notes");
        PdfDocument outer = PdfDocument.Create(new PdfCreateOptions()
            .AddEmbeddedFile("notes.pdf", inner, "application/pdf", description: "Quarter notes")
            .AddEmbeddedFile("payload.bin", new byte[] { 1, 2, 3 }, "application/octet-stream"));
        outer.Paragraph(paragraph => paragraph.AddText("Cover page"));
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddPdfHandler().Build();

        OfficeDocumentReadResult result = reader.ReadDocument(outer.ToBytes(), "outer.pdf");

        // Only an attachment with a registered handler becomes a nested document; payload.bin stays metadata.
        OfficeDocumentNestedResult nested = Assert.Single(result.NestedDocuments);
        Assert.Equal("outer.pdf!/attachments/0000/notes.pdf", nested.Path);
        Assert.Equal(ReaderInputKind.Pdf, nested.Document.Kind);
        Assert.Contains("Embedded quarterly notes", nested.Document.Markdown, StringComparison.Ordinal);
        OfficeDocumentMetadataEntry container = Assert.Single(nested.Document.Metadata, entry => entry.Category == "pdf.container.attachment");
        Assert.Equal("notes.pdf", container.Attributes["fileName"]);
        Assert.Equal("application/pdf", container.Attributes["mimeType"]);
        Assert.Equal("Quarter notes", container.Attributes["description"]);
        Assert.Equal(inner.Length.ToString(System.Globalization.CultureInfo.InvariantCulture), container.Attributes["decodedSizeBytes"]);
        Assert.Equal("attachment", container.Attributes["containerRelationship"]);
        Assert.Equal(2, result.Metadata.Count(entry => entry.Category == "pdf.attachment" && entry.ValueType == "object"));

        // Chunk-only consumers see the attachment text after the page chunks, located at the attachment path.
        ReaderChunk attachmentChunk = Assert.Single(result.Chunks, chunk => chunk.Location.Path == nested.Path);
        Assert.Contains("Embedded quarterly notes", attachmentChunk.Markdown ?? attachmentChunk.Text, StringComparison.Ordinal);
        Assert.Equal(1, attachmentChunk.Location.Page);
        Assert.Same(result.Chunks[result.Chunks.Count - 1], attachmentChunk);
    }

    [Fact]
    public void ReadDocument_MarksEmbeddedFilesOfAPortfolioAsPortfolioItems() {
        PdfDocument portfolio = PdfDocument.Create(new PdfCreateOptions()
            .AddEmbeddedFile("first.pdf", CreatePdf("First portfolio item"), "application/pdf")
            .SetPortfolio(new PdfPortfolioOptions { InitialDocumentFileName = "first.pdf" }));
        portfolio.Paragraph(paragraph => paragraph.AddText("Portfolio cover"));
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddPdfHandler().Build();

        OfficeDocumentReadResult result = reader.ReadDocument(portfolio.ToBytes(), "portfolio.pdf");

        OfficeDocumentNestedResult item = Assert.Single(result.NestedDocuments);
        Assert.Contains("First portfolio item", item.Document.Markdown, StringComparison.Ordinal);
        OfficeDocumentMetadataEntry container = Assert.Single(item.Document.Metadata, entry => entry.Category == "pdf.container.attachment");
        Assert.Equal("portfolio-item", container.Attributes["containerRelationship"]);
    }

    [Fact]
    public void ReadDocument_ChargesAttachmentPayloadToTheSharedNestedInputBudget() {
        byte[] inner = CreatePdf("Budgeted attachment");
        PdfDocument outer = PdfDocument.Create(new PdfCreateOptions().AddEmbeddedFile("inner.pdf", inner, "application/pdf"));
        outer.Paragraph(paragraph => paragraph.AddText("Cover page"));
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddPdfHandler().Build();
        var options = new ReaderOptions {
            ResourceLimits = new ReaderResourceLimits { MaxNestedInputBytes = inner.Length - 1 }
        };

        ReaderResourceLimitException error = Assert.Throws<ReaderResourceLimitException>(() =>
            reader.ReadDocument(outer.ToBytes(), "outer.pdf", options));

        Assert.Equal(nameof(ReaderResourceLimits.MaxNestedInputBytes), error.LimitName);
    }

    [Fact]
    public void ReadDocument_ReportsTextCoveredByMarkupAnnotationsInMetadataAndPageMarkdown() {
        PdfDocument source = PdfDocument.Create();
        source.Paragraph(paragraph => paragraph.AddText("Quarterly revenue grew strongly"));
        source.Paragraph(paragraph => paragraph.AddText("Operating costs stayed flat"));
        byte[] plain = source.ToBytes();
        PdfLogicalPage layout = Assert.Single(PdfDocument.Load(new MemoryStream(plain, writable: false)).Read().Pages);
        PdfLogicalTextBlock revenue = Assert.Single(layout.TextBlocks, block => block.Text.StartsWith("Quarterly", StringComparison.Ordinal));
        PdfLogicalTextBlock costs = Assert.Single(layout.TextBlocks, block => block.Text.StartsWith("Operating", StringComparison.Ordinal));

        PdfDocument annotated = PdfDocument.Load(new MemoryStream(plain, writable: false)).Annotations.Add(new PdfAnnotationCreateOptions {
            Subtype = "Highlight", Title = "Alice", Contents = "Confirm the figure",
            Rectangle = LineRectangle(revenue), QuadPoints = LineQuad(revenue)
        }).ToDocument();
        annotated = annotated.Annotations.Add(new PdfAnnotationCreateOptions {
            Subtype = "StrikeOut", Rectangle = LineRectangle(costs), QuadPoints = LineQuad(costs)
        }).ToDocument();
        annotated = annotated.Annotations.Add(new PdfAnnotationCreateOptions {
            Subtype = "Text", Contents = "General remark", Rectangle = new[] { 20D, 20D, 40D, 40D }
        }).ToDocument();

        OfficeDocumentReadResult result = PdfReaderAdapter.ReadDocument(annotated.ToBytes(), "review.pdf");

        OfficeDocumentMetadataEntry highlight = Assert.Single(result.Metadata, entry => entry.Category == "pdf.annotation" && entry.Name == "Highlight");
        Assert.Equal("Quarterly revenue grew strongly", highlight.Attributes["coveredText"]);
        Assert.Equal("Alice", highlight.Attributes["author"]);
        Assert.Equal("Confirm the figure", highlight.Value);
        Assert.Equal(1, highlight.Location!.Page);
        OfficeDocumentMetadataEntry strikeOut = Assert.Single(result.Metadata, entry => entry.Category == "pdf.annotation" && entry.Name == "StrikeOut");
        Assert.Equal("Operating costs stayed flat", strikeOut.Attributes["coveredText"]);
        OfficeDocumentMetadataEntry note = Assert.Single(result.Metadata, entry => entry.Category == "pdf.annotation" && entry.Name == "Text");
        Assert.False(note.Attributes.ContainsKey("coveredText"));
        Assert.Equal("2", Assert.Single(result.Metadata, entry => entry.Id == "pdf-annotation-covered-text-count").Value);

        string markdown = Assert.Single(result.Chunks).Markdown!;
        Assert.EndsWith(
            "[Highlight by Alice: \"Quarterly revenue grew strongly\"; note: Confirm the figure]" + Environment.NewLine + Environment.NewLine +
            "[StrikeOut: \"Operating costs stayed flat\"]",
            markdown.Replace("\r\n", "\n").Replace("\n", Environment.NewLine), StringComparison.Ordinal);
        Assert.DoesNotContain("General remark", markdown, StringComparison.Ordinal);
    }

    private static byte[] CreatePdf(string text) {
        PdfDocument document = PdfDocument.Create();
        document.Paragraph(paragraph => paragraph.AddText(text));
        return document.ToBytes();
    }

    // A markup quad spanning one rendered line: from slightly below the baseline to the top of the glyphs.
    private static double[] LineQuad(PdfLogicalTextBlock block) {
        double bottom = block.BaselineY - block.FontSize * 0.3D;
        double top = block.BaselineY + block.FontSize;
        return new[] { block.XStart - 1D, top, block.XEnd + 1D, top, block.XStart - 1D, bottom, block.XEnd + 1D, bottom };
    }

    private static double[] LineRectangle(PdfLogicalTextBlock block) =>
        new[] { block.XStart - 1D, block.BaselineY - block.FontSize * 0.3D, block.XEnd + 1D, block.BaselineY + block.FontSize };
}
