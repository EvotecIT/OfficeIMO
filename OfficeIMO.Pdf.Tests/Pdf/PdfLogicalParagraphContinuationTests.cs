using System.Globalization;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfLogicalParagraphContinuationTests {
    [Fact]
    public void ParagraphContinuations_MergesPageEdgeSoftHyphenWithStructuralEvidence() {
        PdfDocumentReadResult document = LoadTwoPageDocument(
            "The paragraph ends with a discr\u00AD",
            "etionary break and continues here");

        PdfLogicalParagraphContinuationGroup group = Assert.Single(document.GetParagraphContinuationGroups(
            new PdfLogicalParagraphContinuationOptions { RejoinSoftHyphens = true }));

        Assert.True(group.SpansPages);
        Assert.Equal(2, group.Segments.Count);
        Assert.Equal("The paragraph ends with a discretionary break and continues here", group.Text);
        Assert.Equal(1, group.FirstPageNumber);
        Assert.Equal(2, group.LastPageNumber);
        Assert.Equal(1, group.RejoinedSoftHyphenCount);
        Assert.True(group.Confidence >= 0.75D);
        Assert.True(group.Evidence.HasFlag(PdfLogicalParagraphContinuationEvidence.AdjacentPages));
        Assert.True(group.Evidence.HasFlag(PdfLogicalParagraphContinuationEvidence.PageEdges));
        Assert.True(group.Evidence.HasFlag(PdfLogicalParagraphContinuationEvidence.CompatibleGeometry));
        Assert.True(group.Evidence.HasFlag(PdfLogicalParagraphContinuationEvidence.CompatibleTypography));
        Assert.True(group.Evidence.HasFlag(PdfLogicalParagraphContinuationEvidence.SoftHyphenBreak));
    }

    [Fact]
    public void ParagraphContinuations_PreservesAuthoredHyphenWhenSoftHyphenJoiningIsEnabled() {
        PdfDocumentReadResult document = LoadTwoPageDocument(
            "A state-of-the-",
            "art system continues here");

        PdfLogicalParagraphContinuationGroup group = Assert.Single(document.GetParagraphContinuationGroups(
            new PdfLogicalParagraphContinuationOptions { RejoinSoftHyphens = true }));

        Assert.Equal("A state-of-the- art system continues here", group.Text);
        Assert.Equal(0, group.RejoinedSoftHyphenCount);
        Assert.False(group.Evidence.HasFlag(PdfLogicalParagraphContinuationEvidence.SoftHyphenBreak));
    }

    [Fact]
    public void ParagraphContinuations_DoesNotMergeCompletedSentence() {
        PdfDocumentReadResult document = LoadTwoPageDocument(
            "This sentence is complete.",
            "another paragraph starts here");

        IReadOnlyList<PdfLogicalParagraphContinuationGroup> groups = document.GetParagraphContinuationGroups();

        Assert.Equal(2, groups.Count);
        Assert.All(groups, group => Assert.False(group.SpansPages));
        Assert.All(groups, group => Assert.Equal(1D, group.Confidence));
        Assert.All(groups, group => Assert.Equal(PdfLogicalParagraphContinuationEvidence.None, group.Evidence));
    }

    [Fact]
    public void ParagraphContinuations_DoesNotMergeDifferentColumns() {
        PdfDocumentReadResult document = LoadTwoPageDocument(
            "The paragraph continues without punctuation",
            "another column begins here",
            secondPageX: 170);

        IReadOnlyList<PdfLogicalParagraphContinuationGroup> groups = document.GetParagraphContinuationGroups();

        Assert.Equal(2, groups.Count);
        Assert.All(groups, group => Assert.False(group.SpansPages));
    }

    [Fact]
    public void ParagraphContinuations_DoesNotRequireLowercaseContinuation() {
        PdfDocumentReadResult document = LoadTwoPageDocument(
            "The paragraph continues without punctuation",
            "Stockholm remains in the same paragraph");

        PdfLogicalParagraphContinuationGroup group = Assert.Single(document.GetParagraphContinuationGroups());

        Assert.True(group.SpansPages);
        Assert.DoesNotContain("Lowercase", group.Evidence.ToString(), StringComparison.Ordinal);
    }

    [Fact]
    public void ParagraphContinuations_KeepCompactPageEdgeFragmentsAsBodyWithoutRepeatedOrTaggedEvidence() {
        string firstContent =
            "BT /F1 12 Tf 40 150 Td (Earlier complete sentence.) Tj ET\n" +
            "BT /F1 12 Tf 40 20 Td (continues without terminal) Tj ET";
        string secondContent =
            "BT /F1 12 Tf 40 280 Td (onto the next page) Tj ET\n" +
            "BT /F1 12 Tf 40 150 Td (Later complete sentence.) Tj ET";
        byte[] pdf = BuildPdf(
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Kids [3 0 R 4 0 R] /Count 2 /Resources << /Font << /F1 7 0 R >> >> >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Contents 5 0 R >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Contents 6 0 R >>",
            StreamObject(firstContent),
            StreamObject(secondContent),
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>");

        PdfDocumentReadResult document = PdfDocumentReadResult.Load(
            pdf,
            new PdfTextLayoutOptions { ForceSingleColumn = true });

        Assert.Empty(document.Pages.SelectMany(static page => page.Headers));
        Assert.Empty(document.Pages.SelectMany(static page => page.Footers));
        PdfLogicalParagraphContinuationGroup continuation = Assert.Single(
            document.GetParagraphContinuationGroups(), static group => group.SpansPages);
        Assert.Equal("continues without terminal onto the next page", continuation.Text);
    }

    [Fact]
    public void ParagraphContinuations_CanDisableCrossPageInference() {
        PdfDocumentReadResult document = LoadTwoPageDocument(
            "The paragraph continues without punctuation",
            "another segment begins here");

        IReadOnlyList<PdfLogicalParagraphContinuationGroup> groups = document.GetParagraphContinuationGroups(
            new PdfLogicalParagraphContinuationOptions { MergePageContinuations = false });

        Assert.Equal(2, groups.Count);
        Assert.All(groups, group => Assert.False(group.SpansPages));
    }

    [Fact]
    public void ParagraphContinuations_RejectsInvalidConfidence() {
        PdfDocumentReadResult document = LoadTwoPageDocument("first segment", "second segment");

        Assert.Throws<ArgumentOutOfRangeException>(() => document.GetParagraphContinuationGroups(
            new PdfLogicalParagraphContinuationOptions { MinimumConfidence = 1.1D }));
    }

    [Fact]
    public void ReaderParagraphContinuations_UsesPublicFacadeAndPreflight() {
        byte[] pdf = BuildTwoPagePdf(
            "The paragraph continues without punctuation",
            "another segment begins here",
            secondPageX: 40);
        PdfDocument source = PdfDocument.Load(pdf);

        PdfLogicalParagraphContinuationGroup group = Assert.Single(source.Reader.ParagraphContinuations());
        PdfOperationResult<IReadOnlyList<PdfLogicalParagraphContinuationGroup>> attempt = source.Reader.TryParagraphContinuations();

        Assert.True(group.SpansPages);
        Assert.True(attempt.Succeeded);
        Assert.Equal(PdfPreflightCapability.ReadLogicalObjects, attempt.Capability);
        Assert.Single(attempt.RequireValue());
    }

    [Fact]
    public void ReaderParagraphContinuations_PageSelector_UsesDocumentRelativeSelectionContract() {
        PdfDocument source = PdfDocument.Load(BuildTwoPagePdf(
            "The paragraph continues without punctuation",
            "another segment begins here",
            secondPageX: 40));

        IReadOnlyList<PdfLogicalParagraphContinuationGroup> groups = source.Reader.ParagraphContinuations(PdfPageSelector.Parse("1..last"));
        PdfOperationResult<IReadOnlyList<PdfLogicalParagraphContinuationGroup>> attempt = source.Reader.TryParagraphContinuations(PdfPageSelector.Parse("all"));

        Assert.True(Assert.Single(groups).SpansPages);
        Assert.True(attempt.Succeeded);
        Assert.True(Assert.Single(attempt.RequireValue()).SpansPages);
    }

    private static PdfDocumentReadResult LoadTwoPageDocument(
        string firstPageText,
        string secondPageText,
        double secondPageX = 40) {
        return PdfDocumentReadResult.Load(
            BuildTwoPagePdf(firstPageText, secondPageText, secondPageX),
            new PdfTextLayoutOptions { ForceSingleColumn = true });
    }

    private static byte[] BuildTwoPagePdf(string firstPageText, string secondPageText, double secondPageX) {
        string firstContent = "BT /F1 12 Tf 40 20 Td (" + Escape(firstPageText) + ") Tj ET";
        string secondContent = "BT /F1 12 Tf " + secondPageX.ToString("0.###", CultureInfo.InvariantCulture) + " 280 Td (" + Escape(secondPageText) + ") Tj ET";
        return BuildPdf(
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Kids [3 0 R 4 0 R] /Count 2 /Resources << /Font << /F1 7 0 R >> >> >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Contents 5 0 R >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Contents 6 0 R >>",
            StreamObject(firstContent),
            StreamObject(secondContent),
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>");
    }

    private static string Escape(string value) =>
        value.Replace("\\", "\\\\").Replace("(", "\\(").Replace(")", "\\)");

    private static string StreamObject(string content) {
        int length = Encoding.Latin1.GetByteCount(content);
        return "<< /Length " + length.ToString(CultureInfo.InvariantCulture) + " >>\nstream\n" + content + "\nendstream";
    }

    private static byte[] BuildPdf(params string[] objects) {
        var builder = new StringBuilder("%PDF-1.7\n");
        var offsets = new List<int>(objects.Length);
        for (int i = 0; i < objects.Length; i++) {
            offsets.Add(Encoding.Latin1.GetByteCount(builder.ToString()));
            builder.Append(i + 1).Append(" 0 obj\n").Append(objects[i]).Append("\nendobj\n");
        }

        int xrefOffset = Encoding.Latin1.GetByteCount(builder.ToString());
        builder.Append("xref\n0 ").Append(objects.Length + 1).Append("\n0000000000 65535 f \n");
        for (int i = 0; i < offsets.Count; i++) {
            builder.Append(offsets[i].ToString("D10", CultureInfo.InvariantCulture)).Append(" 00000 n \n");
        }
        builder.Append("trailer\n<< /Root 1 0 R /Size ").Append(objects.Length + 1).Append(" >>\nstartxref\n")
            .Append(xrefOffset.ToString(CultureInfo.InvariantCulture)).Append("\n%%EOF\n");
        return Encoding.Latin1.GetBytes(builder.ToString());
    }
}
