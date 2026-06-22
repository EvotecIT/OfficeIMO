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
    public void Apply_RemovesOnlyRepeatedTextInsideSelectedArea() {
        byte[] source = PdfDocument.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .Paragraph(paragraph => paragraph.Text("Repeated secret"))
            .Paragraph(paragraph => paragraph.Text("Repeated secret"))
            .ToBytes();
        PdfRedactionArea[] areas = FindAreasForText(source, "Repeated secret");

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { areas[0] });
        string text = PdfTextExtractor.ExtractAllText(redacted);

        Assert.Equal(1, CountOccurrences(text, "Repeated secret"));
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

    private static PdfRedactionArea FindAreaForText(byte[] pdf, string text) {
        return FindAreasForText(pdf, text).Single();
    }

    private static PdfRedactionArea[] FindAreasForText(byte[] pdf, string text) {
        return PdfLogicalDocument.Load(pdf)
            .TextBlocks
            .Where(item => item.Text.Contains(text, StringComparison.Ordinal))
            .Select(static block => {
                double x = Math.Min(block.XStart, block.XEnd) - 2D;
                double width = Math.Abs(block.XEnd - block.XStart) + 4D;
                return new PdfRedactionArea(block.PageNumber, x, block.BaselineY - 14D, width, 20D, "secret");
            })
            .ToArray();
    }

    private static int CountOccurrences(string value, string search) {
        int count = 0;
        int index = 0;
        while ((index = value.IndexOf(search, index, StringComparison.Ordinal)) >= 0) {
            count++;
            index += search.Length;
        }

        return count;
    }
}
