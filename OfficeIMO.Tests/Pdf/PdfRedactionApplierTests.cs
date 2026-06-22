using System.Globalization;
using System.Text;
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

    [Fact]
    public void Apply_RejectsRedactionAreasOutsideDocument() {
        byte[] source = BuildRedactionSource();

        var exception = Assert.Throws<ArgumentOutOfRangeException>(() => PdfRedactionApplier.Apply(source, new[] {
            new PdfRedactionArea(2, 0, 0, 100, 100)
        }));

        Assert.Contains("outside the document page count", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RedactionArea_RejectsNonFiniteCoordinates() {
        Assert.Equal("x", Assert.Throws<ArgumentOutOfRangeException>(() => new PdfRedactionArea(1, double.NaN, 0, 100, 100)).ParamName);
        Assert.Equal("y", Assert.Throws<ArgumentOutOfRangeException>(() => new PdfRedactionArea(1, 0, double.PositiveInfinity, 100, 100)).ParamName);
        Assert.Equal("width", Assert.Throws<ArgumentOutOfRangeException>(() => new PdfRedactionArea(1, 0, 0, double.NegativeInfinity, 100)).ParamName);
        Assert.Equal("height", Assert.Throws<ArgumentOutOfRangeException>(() => new PdfRedactionArea(1, 0, 0, 100, double.NaN)).ParamName);
    }

    [Fact]
    public void Apply_PreservesGraphicsTransformWhenRemovingMatchedText() {
        byte[] source = BuildTransformedTextPdf();

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] {
            new PdfRedactionArea(1, 95, 82, 180, 30)
        });

        string text = PdfTextExtractor.ExtractAllText(redacted);
        Assert.DoesNotContain("Transformed secret", text, StringComparison.Ordinal);
    }

    [Fact]
    public void Apply_ScrubsTextInsideFormXObjects() {
        byte[] source = BuildFormXObjectTextPdf();

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] {
            new PdfRedactionArea(1, 95, 82, 180, 30)
        });

        string text = PdfTextExtractor.ExtractAllText(redacted);
        Assert.DoesNotContain("Form secret", text, StringComparison.Ordinal);
        Assert.DoesNotContain("Form secret", PdfEncoding.Latin1GetString(redacted), StringComparison.Ordinal);
    }

    [Fact]
    public void Apply_RemovesDirectAnnotationDictionariesAndLinkedPopups() {
        byte[] source = BuildDirectAnnotationWithPopupPdf();

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] {
            new PdfRedactionArea(1, 15, 15, 140, 80)
        });

        PdfDocumentInfo info = PdfInspector.Inspect(redacted);
        string raw = PdfEncoding.Latin1GetString(redacted);
        Assert.Equal(0, info.AnnotationCount);
        Assert.DoesNotContain("Direct redaction note", raw, StringComparison.Ordinal);
        Assert.DoesNotContain("/Subtype /Popup", raw, StringComparison.Ordinal);
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

    private static byte[] BuildTransformedTextPdf() {
        var objects = new List<string> {
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            BuildStream("q\n1 0 0 1 100 100 cm\nBT\n/F1 12 Tf\n0 0 Td\n(Transformed secret) Tj\nET\nQ")
        };

        return Encoding.ASCII.GetBytes(BuildPdf(objects));
    }

    private static byte[] BuildFormXObjectTextPdf() {
        var objects = new List<string> {
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Resources << /XObject << /Fm1 5 0 R >> >> /Contents 6 0 R >>",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            BuildStream("BT\n/F1 12 Tf\n0 0 Td\n(Form secret) Tj\nET", "/Type /XObject /Subtype /Form /BBox [0 0 200 50] /Resources << /Font << /F1 4 0 R >> >>"),
            BuildStream("q\n1 0 0 1 100 100 cm\n/Fm1 Do\nQ")
        };

        return Encoding.ASCII.GetBytes(BuildPdf(objects));
    }

    private static byte[] BuildDirectAnnotationWithPopupPdf() {
        var objects = new List<string> {
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Annots [<< /Type /Annot /Subtype /Text /Rect [20 20 40 40] /Contents (Direct redaction note) /Popup 5 0 R >> 5 0 R] /Contents 4 0 R >>",
            BuildStream("BT\n/F1 12 Tf\n72 720 Td\n(Annotation carrier) Tj\nET"),
            "<< /Type /Annot /Subtype /Popup /Rect [45 20 120 80] >>"
        };

        return Encoding.ASCII.GetBytes(BuildPdf(objects));
    }

    private static string BuildStream(string content, string dictionaryEntries = "") {
        byte[] bytes = Encoding.ASCII.GetBytes(content);
        return "<< " + dictionaryEntries + (dictionaryEntries.Length == 0 ? string.Empty : " ") + "/Length " + bytes.Length.ToString(CultureInfo.InvariantCulture) + " >>\nstream\n" + content + "\nendstream";
    }

    private static string BuildPdf(IReadOnlyList<string> objects) {
        var builder = new StringBuilder();
        builder.AppendLine("%PDF-1.7");
        for (int i = 0; i < objects.Count; i++) {
            builder.Append((i + 1).ToString(CultureInfo.InvariantCulture)).AppendLine(" 0 obj");
            builder.AppendLine(objects[i]);
            builder.AppendLine("endobj");
        }

        builder.AppendLine("trailer");
        builder.Append("<< /Root 1 0 R /Size ").Append(objects.Count + 1).AppendLine(" >>");
        builder.AppendLine("startxref");
        builder.AppendLine("123");
        builder.AppendLine("%%EOF");
        return builder.ToString();
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
