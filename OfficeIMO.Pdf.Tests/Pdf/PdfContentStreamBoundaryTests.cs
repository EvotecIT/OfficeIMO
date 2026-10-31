using System.Globalization;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfContentStreamBoundaryTests {
    [Theory]
    [InlineData("1.7")]
    [InlineData("2.0")]
    public void ReadingSeparatesAdjacentOperatorsAndKeepsStreamIdentity(string version) {
        byte[] source = BuildPdf(version,
            "q 2 0 0 2 0 0 cm BT /F1 12 Tf 1 0 0 1 10 50 Tm (FIRST) Tj ET Q",
            "q BT /F1 12 Tf 1 0 0 1 10 70 Tm (SECOND) Tj ET Q");

        var spans = PdfReadDocument.Open(source).Pages[0].GetTextSpans();

        Assert.Equal(new[] { "FIRST", "SECOND" }, spans.Select(span => span.Text));
        Assert.Equal(20D, spans[0].X, 5);
        Assert.Equal(10D, spans[1].X, 5);
        Assert.Equal(70D, spans[1].Y, 5);
        Assert.NotEqual(spans[0].ContentStreamObjectNumber, spans[1].ContentStreamObjectNumber);
    }

    [Theory]
    [InlineData("1.7")]
    [InlineData("2.0")]
    public void ImportingPagePreservesOperatorsAndOperandsAcrossStreamBoundaries(string version) {
        string[] contents = {
            "q 2 0 0 2 0 0 cm Q",
            "q BT /F1 12 Tf 1 0 0 1 10",
            "70 Tm (SECOND) Tj ET Q"
        };
        byte[] source = BuildPdf(version, contents);
        byte[] expected = BuildPdf(version, string.Join("\n", contents));
        PdfDocument target = PdfDocument.Load(BuildPdf(version, ""));

        byte[] imported = target.Stamp.OverlayPage(source, new PdfPageOverlayOptions {
            X = 0D, Y = 0D, Fit = PdfPageOverlayFit.None
        }).ToBytes();

        Assert.Equal(Render(expected), Render(source));
        Assert.Equal(Render(expected), Render(imported));
        using var independent = UglyToad.PdfPig.PdfDocument.Open(imported);
        Assert.Equal("SECOND", independent.GetPage(1).Text);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RedactionUsesTheSameBoundariesAndOffsetsAsReading(bool splitTextObject) {
        var contents = new List<string> { "BT /F1 12 Tf 1 0 0 1 10 110 Tm (BEFORE) Tj ET" };
        if (splitTextObject) contents.AddRange(new[] { "BT /F1 12 Tf 1 0 0 1 10", "75 Tm (SECRET) Tj ET" });
        else contents.Add("BT /F1 12 Tf 1 0 0 1 10 75 Tm (SECRET) Tj ET");
        contents.Add("BT /F1 12 Tf 1 0 0 1 10 40 Tm (AFTER) Tj ET");
        byte[] source = BuildPdf("2.0", contents.ToArray());
        PdfDocument document = PdfDocument.Load(source);
        PdfRedactionPlan plan = document.Redactions.Search(new PdfRedactionSearchOptions().AddLiteral("SECRET"));

        PdfRedactionApplyResult result = document.Redactions.ApplyWithEvidence(plan,
            verificationOptions: new PdfRedactionVerificationOptions { RequireCompleteStreamInspection = true }
                .RequireRemovedText("SECRET").RequireRetainedText("BEFORE", "AFTER"));

        Assert.True(result.IsVerified, result.Evidence.Summary);
        var spans = PdfReadDocument.Open(result.ToDocument().ToBytes()).Pages[0].GetTextSpans();
        Assert.Equal(new[] { "BEFORE", "AFTER" }, spans.Select(span => span.Text));
        Assert.Equal(new[] { 110D, 40D }, spans.Select(span => span.Y));
        Assert.Equal(source, document.ToBytes());
    }

    [Fact]
    public void RedactionFindsAFormWhoseResourceNameAndInvocationAreInDifferentStreams() {
        byte[] source = BuildPdf("2.0", "/Sensitive", "Do");
        source = PdfDocumentObjectGraphRewriter.Rewrite(source, null, null, (objects, security) => {
            const int formNumber = 7;
            var dictionary = new PdfDictionary();
            dictionary.Items["Type"] = new PdfName("XObject");
            dictionary.Items["Subtype"] = new PdfName("Form");
            var box = new PdfArray();
            foreach (int coordinate in new[] { 0, 0, 200, 150 }) box.Items.Add(new PdfNumber(coordinate));
            dictionary.Items["BBox"] = box;
            var fonts = new PdfDictionary();
            fonts.Items["F1"] = new PdfReference(4, 0);
            var resources = new PdfDictionary();
            resources.Items["Font"] = fonts;
            dictionary.Items["Resources"] = resources;
            objects[formNumber] = new PdfIndirectObject(formNumber, 0,
                new PdfStream(dictionary, Encoding.ASCII.GetBytes("BT /F1 12 Tf 1 0 0 1 10 75 Tm (SECRET) Tj ET")));
            var xObjects = new PdfDictionary();
            xObjects.Items["Sensitive"] = new PdfReference(formNumber, 0);
            var page = (PdfDictionary)objects[3].Value;
            ((PdfDictionary)page.Items["Resources"]).Items["XObject"] = xObjects;
            return security.InfoObjectNumber;
        });
        PdfDocument document = PdfDocument.Load(source);
        var plan = document.Redactions.Search(new PdfRedactionSearchOptions().AddLiteral("SECRET"));
        Assert.Single(plan.Areas);

        var result = document.Redactions.ApplyWithEvidence(plan,
            verificationOptions: new PdfRedactionVerificationOptions { RequireCompleteStreamInspection = true }
                .RequireRemovedText("SECRET"));

        Assert.True(result.IsVerified, result.Evidence.Summary);
        Assert.Empty(result.ToDocument().Read().TextBlocks);
        Assert.Equal(source, document.ToBytes());
    }

    [Fact]
    public void RedactionRetainsParagraphsAdjacentToTheReviewedBlockCoordinates() {
        byte[] source = PdfDocument.Create(new PdfCreateOptions { CompressContentStreams = false })
            .Paragraph(p => p.Text("Visible before"))
            .Paragraph(p => p.Text("Secret account 123-45"))
            .Paragraph(p => p.Text("Visible after"))
            .ToBytes();
        var page = PdfReadDocument.Open(source).Pages[0];
        var block = Assert.Single(PdfDocument.Load(source).Read().TextBlocks,
            block => block.Text.Contains("Secret account", StringComparison.Ordinal));
        var area = new PdfRedactionArea(1, Math.Min(block.XStart, block.XEnd) - 2D,
            block.BaselineY - block.FontSize * 0.25D,
            Math.Abs(block.XEnd - block.XStart) + 4D, block.FontSize * 1.1D);
        var before = page.GetTextSpans().Where(span => !span.Text.Contains("Secret account", StringComparison.Ordinal)).ToArray();

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });

        var after = PdfReadDocument.Open(redacted).Pages[0].GetTextSpans();
        Assert.Equal(before.Select(span => span.Text), after.Select(span => span.Text));
        Assert.Equal(before.Select(span => (span.X, span.Y)), after.Select(span => (span.X, span.Y)));
        using var independent = UglyToad.PdfPig.PdfDocument.Open(redacted);
        Assert.Contains("Visible before", independent.GetPage(1).Text, StringComparison.Ordinal);
        Assert.Contains("Visible after", independent.GetPage(1).Text, StringComparison.Ordinal);
        Assert.DoesNotContain("123-45", independent.GetPage(1).Text, StringComparison.Ordinal);
    }

    private static byte[] Render(byte[] pdf) {
        var result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf,
            options: new PdfPageRenderOptions { ContinueOnError = false }));
        Assert.True(result.Succeeded);
        return Assert.IsType<byte[]>(result.Bytes);
    }

    private static byte[] BuildPdf(string version, params string[] contents) {
        var objects = new List<string> {
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 150] /Resources << /Font << /F1 4 0 R >> >> /Contents [" +
                string.Join(" ", Enumerable.Range(5, contents.Length).Select(number => number + " 0 R")) + "] >>",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>"
        };
        objects.AddRange(contents.Select(content => "<< /Length " + content.Length.ToString(CultureInfo.InvariantCulture) +
            " >>\nstream\n" + content + "\nendstream"));
        var pdf = new StringBuilder("%PDF-" + version + "\n");
        var offsets = new List<int>();
        for (int index = 0; index < objects.Count; index++) {
            offsets.Add(pdf.Length);
            pdf.Append(index + 1).Append(" 0 obj\n").Append(objects[index]).Append("\nendobj\n");
        }
        int xref = pdf.Length;
        pdf.Append("xref\n0 ").Append(objects.Count + 1).Append("\n0000000000 65535 f \n");
        foreach (int offset in offsets) pdf.Append(offset.ToString("D10", CultureInfo.InvariantCulture)).Append(" 00000 n \n");
        pdf.Append("trailer << /Root 1 0 R /Size ").Append(objects.Count + 1).Append(" >>\nstartxref\n").Append(xref).Append("\n%%EOF\n");
        return Encoding.ASCII.GetBytes(pdf.ToString());
    }
}
