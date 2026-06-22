using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfRedactionApplierTests {
    [Fact]
    public void Apply_RemovesMatchedTextAndKeepsUnmatchedTextExtractable() {
        byte[] source = BuildRedactionSource();
        PdfRedactionArea area = FindAreaForText(source, "Secret account 123-45");

        PdfRedactionPlan plan = PdfRedactionPlanner.Plan(source, new[] { area });
        Assert.True(plan.HasMatches);
        Assert.Contains(plan.Matches, match => match.Text != null && match.Text.Contains("Secret account", StringComparison.Ordinal));

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });
        string text = PdfTextExtractor.ExtractAllText(redacted);

        Assert.Contains("Visible before", text, StringComparison.Ordinal);
        Assert.Contains("Visible after", text, StringComparison.Ordinal);
        Assert.DoesNotContain("Secret account", text, StringComparison.Ordinal);
        Assert.DoesNotContain("123-45", text, StringComparison.Ordinal);

        PdfRedactionPlan redactedPlan = PdfRedactionPlanner.Plan(redacted, new[] { area });
        Assert.DoesNotContain(redactedPlan.Matches, match => match.Text != null && match.Text.Contains("Secret account", StringComparison.Ordinal));
    }

    [Fact]
    public void ApplyRedactions_FacadeReturnsRedactedDocumentAndTryResult() {
        byte[] source = BuildRedactionSource();
        PdfRedactionArea area = FindAreaForText(source, "Secret account 123-45");
        using PdfDocument document = PdfDocument.Open(source);

        PdfDocument redacted = document.ApplyRedactions(new[] { area });
        PdfOperationResult<PdfDocument> result = document.TryApplyRedactions(new[] { area });

        Assert.DoesNotContain("Secret account", redacted.Read.Text(), StringComparison.Ordinal);
        Assert.True(result.Succeeded);
        Assert.DoesNotContain("Secret account", result.RequireValue().Read.Text(), StringComparison.Ordinal);
    }

    [Fact]
    public void Apply_DecodesOctalEscapesBeforeMatchingTextObjects() {
        byte[] source = BuildOctalRedactionSource();
        PdfRedactionArea area = FindAreaForText(source, "Secret account 123-45");
        Assert.Contains("Secret account 123-45", PdfTextExtractor.ExtractAllText(source), StringComparison.Ordinal);

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });
        string text = PdfTextExtractor.ExtractAllText(redacted);

        Assert.DoesNotContain("Secret account", text, StringComparison.Ordinal);
        Assert.DoesNotContain("123-45", text, StringComparison.Ordinal);
    }

    private static byte[] BuildRedactionSource() {
        return PdfDocument.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .Paragraph(paragraph => paragraph.Text("Visible before"))
            .Paragraph(paragraph => paragraph.Text("Secret account 123-45"))
            .Paragraph(paragraph => paragraph.Text("Visible after"))
            .ToBytes();
    }

    private static byte[] BuildOctalRedactionSource() {
        string streamContent = string.Join("\n", new[] {
            "BT",
            "/F1 12 Tf",
            "72 720 Td",
            "(Visible before) Tj",
            "0 -18 Td",
            "(Secret\\040account\\040123-45) Tj",
            "0 -18 Td",
            "(Visible after) Tj",
            "ET"
        });
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            "endobj",
            "5 0 obj",
            "<< /Length " + Encoding.ASCII.GetByteCount(streamContent).ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>",
            "stream",
            streamContent,
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "startxref",
            "123",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static PdfRedactionArea FindAreaForText(byte[] pdf, string text) {
        PdfLogicalTextBlock block = PdfLogicalDocument.Load(pdf)
            .TextBlocks
            .Single(item => item.Text.Contains(text, StringComparison.Ordinal));

        double x = Math.Min(block.XStart, block.XEnd) - 2D;
        double width = Math.Abs(block.XEnd - block.XStart) + 4D;
        return new PdfRedactionArea(block.PageNumber, x, block.BaselineY - 14D, width, 20D, "secret");
    }
}
