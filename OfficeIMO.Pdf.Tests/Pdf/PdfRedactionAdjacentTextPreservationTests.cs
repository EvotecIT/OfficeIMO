using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfRedactionAdjacentTextPreservationTests {
    [Theory]
    [InlineData(PdfRedactionTextSelection.LogicalBlocks)]
    [InlineData(PdfRedactionTextSelection.MatchedGlyphs)]
    public void LiteralRedactionPreservesTheLargerHeadingAboveTheSelectedAccount(PdfRedactionTextSelection selection) {
        byte[] bytes = PdfDocument.Create(pdf => pdf.Content(content => content
            .H1("Supplier onboarding")
            .Paragraph(paragraph => paragraph.Text("Account: 123-45-6789"))
            .Paragraph(paragraph => paragraph.Text("Delivery contact: service@example.test")))).ToBytes();
        PdfDocument source = PdfDocument.Load(bytes);
        PdfRedactionPlan plan = source.Redactions.Search(new PdfRedactionSearchOptions {
            TextSelection = selection
        }.AddLiteral("123-45-6789"));
        Assert.Single(plan.Areas);
        Assert.DoesNotContain(plan.Matches, match => match.Text?.Contains("Supplier", StringComparison.Ordinal) == true);
        var verification = new PdfRedactionVerificationOptions {
            RequireCompleteStreamInspection = true,
            CheckManagedRendering = true
        }.RequireRemovedText("123-45-6789");

        PdfRedactionApplyResult result = source.Redactions.ApplyWithEvidence(plan, verificationOptions: verification);

        Assert.True(result.Evidence.IsVerified, result.Evidence.Summary);
        AssertSavedText(result.ToDocument().ToBytes());
        Assert.Equal(bytes, source.ToBytes());
    }

    [Theory]
    [InlineData(false, "1 0 0 1 0 0")]
    [InlineData(true, "1 0 0 1 0 0")]
    [InlineData(false, "1 0 0.2 0.8 20 30")]
    [InlineData(true, "0 1 -1 0 800 0")]
    public void RectangleRedactionPreservesAdjacentTextAcrossTextObjectsAndTransforms(bool sameTextObject, string matrix) {
        byte[] bytes = BuildSource(sameTextObject, matrix);
        PdfReadPage page = PdfReadDocument.Open(bytes).Pages[0];
        PdfTextSpan account = Assert.Single(page.GetTextSpans(), span => span.Text == "Account: 123-45-6789");
        PdfTextSpanBounds bounds = PdfTextSpanGeometry.GetRedactionGlyphBounds(account, 0D, account.Advance);
        var area = new PdfRedactionArea(1, bounds.Left, bounds.Bottom, bounds.Width, bounds.Height, "account");
        PdfDocument source = PdfDocument.Load(bytes);
        PdfRedactionPlan plan = source.Redactions.Plan(new[] { area });
        Assert.DoesNotContain(plan.Matches, match => match.Text?.Contains("Supplier", StringComparison.Ordinal) == true);

        PdfRedactionApplyResult result = source.Redactions.ApplyWithEvidence(plan,
            verificationOptions: new PdfRedactionVerificationOptions {
                RequireCompleteStreamInspection = true
            }.RequireRemovedText("123-45-6789"));

        Assert.True(result.Evidence.IsVerified, result.Evidence.Summary);
        AssertSavedText(result.ToDocument().ToBytes());
        Assert.Equal(bytes, source.ToBytes());
    }

    private static void AssertSavedText(byte[] output) {
        using var independent = UglyToad.PdfPig.PdfDocument.Open(output);
        string text = string.Concat(independent.GetPage(1).Text.Where(character => !char.IsWhiteSpace(character)));
        Assert.Contains("Supplieronboarding", text, StringComparison.Ordinal);
        Assert.Contains("Deliverycontact:service@example.test", text, StringComparison.Ordinal);
        Assert.DoesNotContain("123-45-6789", text, StringComparison.Ordinal);
    }

    private static byte[] BuildSource(bool sameTextObject, string matrix) {
        byte[] seed = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Fixture seed")).ToBytes();
        int pageObject = PdfReadDocument.Open(seed).Pages[0].ObjectNumber;
        string split = sameTextObject ? string.Empty : "ET BT ";
        string text = "q " + matrix + " cm BT /F1 24 Tf 1 0 0 1 72 702.24 Tm "
            + "[(Supplier onboar) 20 (ding)] TJ " + split
            + "/F1 11 Tf 1 0 0 1 72 674.36 Tm (Account: 123-45-6789) Tj " + split
            + "/F1 11 Tf 1 0 0 1 72 654.34 Tm (Delivery contact: service@example.test) Tj ET Q";
        return PdfDocumentObjectGraphRewriter.Rewrite(seed, null, null, (objects, security) => {
            int number = objects.Keys.Max() + 1;
            objects[number] = new PdfIndirectObject(number, 0, new PdfStream(new PdfDictionary(), Encoding.ASCII.GetBytes(text)));
            ((PdfDictionary)objects[pageObject].Value).Items["Contents"] = new PdfReference(number, 0);
            return security.InfoObjectNumber;
        });
    }
}
