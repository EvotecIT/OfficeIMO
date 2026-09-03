using System;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Filters;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfRedactionVerificationTests {
    [Fact]
    public void AppliedPlanVerificationAcceptsWholeTextObjectRemovalWhenOnlyOneSpanWasReviewed() {
        byte[] source = BuildTextObjectWithTwoSpansPdf();
        PdfRedactionPlan plan = PdfRedactionPlanner.Plan(source, [
            new PdfRedactionArea(1, 8D, 88D, 60D, 18D, "first span")
        ]);

        byte[] redacted = PdfRedactionApplier.Apply(source, plan);
        PdfRedactionVerificationReport report = PdfRedactionVerification.VerifyAppliedPlan(
            redacted,
            plan,
            new PdfRedactionVerificationOptions { RequireCompleteStreamInspection = true });

        Assert.True(report.IsVerified, string.Join("; ", report.Issues.Select(static issue => issue.Message)));
        Assert.DoesNotContain("FIRST", PdfReadDocument.Open(redacted).Pages[0].ExtractText(), StringComparison.Ordinal);
        Assert.DoesNotContain("SECOND", PdfReadDocument.Open(redacted).Pages[0].ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void AppliedPlanVerificationUsesOrientedBoundsForRotatedText() {
        byte[] source = BuildRotatedTextIdentityPdf();
        PdfRedactionPlan plan = PdfRedactionPlanner.Plan(source, [
            new PdfRedactionArea(1, 88D, 55D, 20D, 28D, "vertical text")
        ]);

        byte[] redacted = PdfRedactionApplier.Apply(source, plan);
        PdfRedactionVerificationReport report = PdfRedactionVerification.VerifyAppliedPlan(
            redacted,
            plan,
            new PdfRedactionVerificationOptions { RequireCompleteStreamInspection = true });

        Assert.True(report.IsVerified, string.Join("; ", report.Issues.Select(static issue => issue.Message)));
        Assert.DoesNotContain("VERTICAL", PdfReadDocument.Open(redacted).Pages[0].ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void AppliedPlanVerificationRejectsUnchangedRotatedTextInsideReviewedArea() {
        byte[] source = BuildRotatedTextIdentityPdf();
        PdfRedactionPlan plan = PdfRedactionPlanner.Plan(source, [
            new PdfRedactionArea(1, 88D, 55D, 20D, 28D, "vertical text")
        ]);

        PdfRedactionVerificationReport report = PdfRedactionVerification.VerifyAppliedPlan(
            source,
            plan,
            new PdfRedactionVerificationOptions { RequireCompleteStreamInspection = true });

        Assert.False(report.IsVerified);
        Assert.Contains(report.Issues, static issue => issue.Feature == "RedactionPlanResidual");
    }

    [Fact]
    public void SearchRedactionsUsesOrientedBoundsForRotatedText() {
        byte[] source = BuildRotatedTextIdentityPdf();

        PdfRedactionPlan plan = PdfRedactionPlanner.Search(
            source,
            new PdfRedactionSearchOptions().AddLiteral("VERTICAL"));

        PdfRedactionArea area = Assert.Single(plan.Areas);
        Assert.True(area.Height > area.Width);
        Assert.Contains(plan.Matches, static match => match.Kind == PdfRedactionMatchKind.TextBlock);
    }

    [Fact]
    public void AppliedPlanUsesThePlannerAscentAndDescentBoundsForTextScrubbing() {
        byte[] source = BuildTextPaintIdentityPdf("0 g");
        PdfRedactionPlan plan = PdfRedactionPlanner.Plan(source, [
            new PdfRedactionArea(1, 20D, 102D, 70D, 4D, "glyph ascent")
        ]);

        Assert.True(plan.HasMatches);
        byte[] redacted = PdfRedactionApplier.Apply(source, plan);
        PdfRedactionVerificationReport report = PdfRedactionVerification.VerifyAppliedPlan(
            redacted,
            plan,
            new PdfRedactionVerificationOptions { RequireCompleteStreamInspection = true });

        Assert.True(report.IsVerified, string.Join("; ", report.Issues.Select(static issue => issue.Message)));
        Assert.DoesNotContain("Visible text", PdfReadDocument.Open(redacted).Pages[0].ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void AppliedPlanVerificationRejectsReorderedPagesThatDifferOnlyByVectorPaths() {
        byte[] source = PdfDocument.Create(compose => {
            compose.Page(page => page.Content(content => content.Item(item => item
                .Rectangle(40D, 30D, strokeColor: PdfColor.FromRgb(180, 20, 20), fillColor: PdfColor.FromRgb(255, 220, 220)))));
            compose.Page(page => page.Content(content => content.Item(item => item
                .Rectangle(80D, 30D, strokeColor: PdfColor.FromRgb(20, 20, 180), fillColor: PdfColor.FromRgb(220, 220, 255)))));
        }).ToBytes();
        PdfRedactionPlan plan = PdfRedactionPlanner.Plan(source, [
            new PdfRedactionArea(1, 10D, 10D, 10D, 10D, "reviewed area")
        ]);
        byte[] reordered = PdfPageExtractor.ExtractPages(source, 2, 1);

        PdfRedactionVerificationReport report = PdfRedactionVerification.VerifyAppliedPlan(
            reordered,
            plan,
            new PdfRedactionVerificationOptions { RequireCompleteStreamInspection = true });

        Assert.False(report.IsVerified);
        Assert.Contains(report.Issues, static issue => issue.Feature == "RedactionPlanPageIdentityChanged");
    }

    [Fact]
    public void AppliedPlanVerificationRejectsChangedUnredactedImageDecodeSemantics() {
        byte[] source = BuildImageIdentitySource(includeInvertedDecode: false);
        PdfRedactionPlan plan = PdfRedactionPlanner.Plan(source, [
            new PdfRedactionArea(1, 150D, 80D, 20D, 20D, "reviewed blank area")
        ]);
        byte[] rewritten = BuildImageIdentitySource(includeInvertedDecode: true);

        PdfRedactionVerificationReport report = PdfRedactionVerification.VerifyAppliedPlan(
            rewritten,
            plan,
            new PdfRedactionVerificationOptions { RequireCompleteStreamInspection = true });

        Assert.False(report.IsVerified);
        Assert.Contains(report.Issues, static issue => issue.Feature == "RedactionPlanPageIdentityChanged");
    }

    [Fact]
    public void AppliedPlanVerificationRejectsChangedUnredactedTextPaintState() {
        byte[] source = BuildTextPaintIdentityPdf("1 0 0 rg");
        PdfRedactionPlan plan = PdfRedactionPlanner.Plan(source, [
            new PdfRedactionArea(1, 150D, 80D, 20D, 20D, "reviewed blank area")
        ]);
        byte[] rewritten = BuildTextPaintIdentityPdf("0 0 1 rg");

        PdfRedactionVerificationReport report = PdfRedactionVerification.VerifyAppliedPlan(
            rewritten,
            plan,
            new PdfRedactionVerificationOptions { RequireCompleteStreamInspection = true });

        Assert.False(report.IsVerified);
        Assert.Contains(report.Issues, static issue => issue.Feature == "RedactionPlanPageIdentityChanged");
    }

    [Fact]
    public void AppliedPlanVerificationRejectsChangedUnredactedTextStrokePaint() {
        AssertPlanIdentityChanged(
            BuildTextVisualIdentityPdf("1 0 0 RG 2 Tr", "1 0 0 1 20 100"),
            BuildTextVisualIdentityPdf("0 0 1 RG 2 Tr", "1 0 0 1 20 100"));
    }

    [Fact]
    public void AppliedPlanVerificationRejectsChangedUnredactedTextStrokeOpacity() {
        AssertPlanIdentityChanged(
            BuildTextOpacityIdentityPdf("GS1"),
            BuildTextOpacityIdentityPdf("GS2"));
    }

    [Fact]
    public void AppliedPlanVerificationRejectsChangedUnredactedTextTransform() {
        AssertPlanIdentityChanged(
            BuildTextVisualIdentityPdf("0 g", "1 0 0 1 20 100"),
            BuildTextVisualIdentityPdf("0 g", "1 0.15 0.2 1 20 100"));
    }

    [Fact]
    public void AppliedPlanVerificationRejectsChangedFontProgramGraph() {
        AssertPlanIdentityChanged(
            BuildEmbeddedFontIdentityPdf("source-font-program"),
            BuildEmbeddedFontIdentityPdf("changed-font-program"));
    }

    [Fact]
    public void AppliedPlanVerificationRejectsChangedUnredactedVectorClip() {
        byte[] source = BuildVectorClipIdentityPdf("10 10 100 100 re W n");
        PdfRedactionPlan plan = PdfRedactionPlanner.Plan(source, [
            new PdfRedactionArea(1, 150D, 20D, 10D, 10D, "reviewed blank area")
        ]);
        byte[] rewritten = BuildVectorClipIdentityPdf("30 30 30 30 re W n");

        PdfRedactionVerificationReport report = PdfRedactionVerification.VerifyAppliedPlan(
            rewritten,
            plan,
            new PdfRedactionVerificationOptions { RequireCompleteStreamInspection = true });

        Assert.False(report.IsVerified);
        Assert.Contains(report.Issues, static issue => issue.Feature == "RedactionPlanPageIdentityChanged");
    }

    [Fact]
    public void ApplyWithEvidenceRejectsResidualVectorPathInsideFormXObject() {
        byte[] source = BuildNestedFormVectorPathPdf();
        PdfDocument document = PdfDocument.Load(source);
        PdfRedactionPlan plan = document.Redactions.Plan([
            new PdfRedactionArea(1, 55D, 55D, 10D, 10D, "nested vector")
        ]);

        PdfRedactionMatch plannedPath = Assert.Single(
            plan.Matches,
            static match => match.Kind == PdfRedactionMatchKind.VectorPath);
        PdfRedactionApplyResult result = document.Redactions.ApplyWithEvidence(
            plan,
            verificationOptions: new PdfRedactionVerificationOptions {
                RequireCompleteStreamInspection = true
            });

        Assert.False(result.IsVerified);
        Assert.Contains(result.Evidence.ResidualMatches, match =>
            match.Kind == PdfRedactionMatchKind.VectorPath &&
            match.Area == plannedPath.Area);
        Assert.Contains(result.Evidence.Verification.Issues, static issue =>
            issue.Feature == "RedactionPlanResidual" &&
            issue.Marker.StartsWith("VectorPath@page:1", StringComparison.Ordinal));
    }

    [Theory]
    [InlineData("0 J 0 j", "2 J 0 j")]
    [InlineData("0 J 0 j", "0 J 2 j")]
    public void AppliedPlanVerificationRejectsChangedUnredactedVectorLineStyle(
        string sourceStyle,
        string rewrittenStyle) {
        AssertPlanIdentityChanged(
            BuildVectorStyleIdentityPdf(sourceStyle),
            BuildVectorStyleIdentityPdf(rewrittenStyle));
    }

    [Theory]
    [InlineData("GS1", "GS2")]
    [InlineData("GS3", "GS4")]
    public void AppliedPlanVerificationRejectsChangedUnredactedGraphicsEffectSelection(
        string sourceState,
        string rewrittenState) {
        AssertPlanIdentityChanged(
            BuildVectorEffectIdentityPdf(sourceState),
            BuildVectorEffectIdentityPdf(rewrittenState));
    }

    [Theory]
    [InlineData("0 0 1 rg 0 0 5 5 re f", 10D)]
    [InlineData("1 0 0 rg 0 0 5 5 re f", 12D)]
    public void AppliedPlanVerificationRejectsChangedUnredactedTilingPattern(
        string rewrittenTileContent,
        double rewrittenHorizontalStep) {
        byte[] source = BuildTilingPatternIdentityPdf("1 0 0 rg 0 0 5 5 re f", 10D);
        PdfRedactionPlan plan = PdfRedactionPlanner.Plan(source, [
            new PdfRedactionArea(1, 150D, 20D, 10D, 10D, "reviewed blank area")
        ]);
        byte[] rewritten = BuildTilingPatternIdentityPdf(rewrittenTileContent, rewrittenHorizontalStep);

        PdfRedactionVerificationReport report = PdfRedactionVerification.VerifyAppliedPlan(
            rewritten,
            plan,
            new PdfRedactionVerificationOptions { RequireCompleteStreamInspection = true });

        Assert.False(report.IsVerified);
        Assert.Contains(report.Issues, static issue => issue.Feature == "RedactionPlanPageIdentityChanged");
    }

    [Fact]
    public void AppliedPlanVerificationRejectsChangedUnredactedAnnotationAppearanceGraph() {
        byte[] source = BuildAnnotationAppearanceIdentityPdf("1 0 0 rg");
        PdfRedactionPlan plan = PdfRedactionPlanner.Plan(source, [
            new PdfRedactionArea(1, 150D, 80D, 20D, 20D, "reviewed blank area")
        ]);
        byte[] rewritten = BuildAnnotationAppearanceIdentityPdf("0 0 1 rg");

        PdfRedactionVerificationReport report = PdfRedactionVerification.VerifyAppliedPlan(
            rewritten,
            plan,
            new PdfRedactionVerificationOptions { RequireCompleteStreamInspection = true });

        Assert.False(report.IsVerified);
        Assert.Contains(report.Issues, static issue => issue.Feature == "RedactionPlanPageIdentityChanged");
    }

    [Theory]
    [InlineData("/CA 0.4", "/CA 0.8")]
    [InlineData("/BS << /W 1 /S /S >>", "/BS << /W 2 /S /D /D [3 2] >>")]
    public void AppliedPlanVerificationRejectsChangedUnredactedAnnotationStyle(
        string sourceStyle,
        string rewrittenStyle) {
        byte[] source = BuildAnnotationAppearanceIdentityPdf("1 0 0 rg", sourceStyle);
        PdfRedactionPlan plan = PdfRedactionPlanner.Plan(source, [
            new PdfRedactionArea(1, 150D, 80D, 20D, 20D, "reviewed blank area")
        ]);
        byte[] rewritten = BuildAnnotationAppearanceIdentityPdf("1 0 0 rg", rewrittenStyle);

        PdfRedactionVerificationReport report = PdfRedactionVerification.VerifyAppliedPlan(
            rewritten,
            plan,
            new PdfRedactionVerificationOptions { RequireCompleteStreamInspection = true });

        Assert.False(report.IsVerified);
        Assert.Contains(report.Issues, static issue => issue.Feature == "RedactionPlanPageIdentityChanged");
    }

    [Fact]
    public void AppliedPlanVerificationRejectsChangedUnredactedAnnotationAppearanceState() {
        AssertPlanIdentityChanged(
            BuildAnnotationAppearanceIdentityPdf("1 0 0 rg", "/AS /On"),
            BuildAnnotationAppearanceIdentityPdf("1 0 0 rg", "/AS /Off"));
    }

    [Fact]
    public void AppliedPlanVerificationAcceptsPreservedUnredactedImageRenderingSemantics() {
        byte[] source = BuildImageIdentitySource(includeInvertedDecode: false);
        PdfRedactionPlan plan = PdfRedactionPlanner.Plan(source, [
            new PdfRedactionArea(1, 150D, 80D, 20D, 20D, "reviewed blank area")
        ]);

        byte[] redacted = PdfRedactionApplier.Apply(source, plan);
        PdfRedactionVerificationReport report = PdfRedactionVerification.VerifyAppliedPlan(
            redacted,
            plan,
            new PdfRedactionVerificationOptions { RequireCompleteStreamInspection = true });

        Assert.True(report.IsVerified, string.Join("; ", report.Issues.Select(static issue => issue.Message)));
    }

    [Fact]
    public void AppliedPlanVerificationReportsResidualContentInsideReviewedArea() {
        byte[] source = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Still present")).ToBytes();
        var area = new PdfRedactionArea(1, 0D, 0D, 600D, 800D, "whole page");
        PdfRedactionPlan plan = PdfDocument.Load(source).Redactions.Plan([area]);

        PdfRedactionVerificationReport report = PdfDocument.Load(source).Redactions.VerifyAppliedPlan(
            plan,
            new PdfRedactionVerificationOptions { RequireCompleteStreamInspection = true });

        Assert.False(report.IsVerified);
        Assert.Contains(report.Issues, issue => issue.Feature == "RedactionPlanResidual");
    }

    [Fact]
    public void AppliedPlanVerificationRejectsARewriteThatDropsAReviewedPage() {
        byte[] source = PdfDocument.Create(compose => {
            compose.Page(page => page.Content(content => content.Item(item => item.Paragraph(paragraph => paragraph.Text("Retained first page")))));
            compose.Page(page => page.Content(content => content.Item(item => item.Paragraph(paragraph => paragraph.Text("Reviewed second page")))));
        }).ToBytes();
        var area = new PdfRedactionArea(2, 0D, 0D, 600D, 800D, "reviewed page");
        PdfRedactionPlan plan = PdfDocument.Load(source).Redactions.Plan([area]);
        byte[] rewritten = PdfPageExtractor.ExtractPages(source, 1);

        PdfRedactionVerificationReport report = PdfDocument.Load(rewritten).Redactions.VerifyAppliedPlan(
            plan,
            new PdfRedactionVerificationOptions { RequireCompleteStreamInspection = true });

        Assert.False(report.IsVerified);
        Assert.Contains(report.Issues, issue => issue.Feature == "RedactionPlanPageCountChanged");
        Assert.Contains(report.Issues, issue => issue.Feature == "RedactionPlanPageMissing");
    }

    [Fact]
    public void AppliedPlanVerificationRejectsSameCountRewriteWithChangedPageGeometry() {
        byte[] source = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Reviewed page")).ToBytes();
        PdfRedactionPlan plan = PdfDocument.Load(source).Redactions.Plan([
            new PdfRedactionArea(1, 0D, 0D, 600D, 800D, "reviewed page")
        ]);
        byte[] rewritten = PdfPageEditor.RotatePages(source, 90, 1);

        PdfRedactionVerificationReport report = PdfDocument.Load(rewritten).Redactions.VerifyAppliedPlan(
            plan,
            new PdfRedactionVerificationOptions { RequireCompleteStreamInspection = true });

        Assert.False(report.IsVerified);
        Assert.Contains(report.Issues, issue => issue.Feature == "RedactionPlanPageIdentityChanged");
    }

    [Fact]
    public void AppliedPlanVerificationRejectsSameCountRewriteWithReorderedPages() {
        byte[] source = PdfDocument.Create(compose => {
            compose.Page(page => page.Content(content => content.Item(item => item.Paragraph(paragraph => paragraph.Text("Reviewed first page")))));
            compose.Page(page => page.Content(content => content.Item(item => item.Paragraph(paragraph => paragraph.Text("Reviewed second page")))));
        }).ToBytes();
        PdfRedactionPlan plan = PdfDocument.Load(source).Redactions.Plan([
            new PdfRedactionArea(1, 0D, 0D, 600D, 800D, "reviewed first page")
        ]);
        byte[] rewritten = PdfPageExtractor.ExtractPages(source, 2, 1);

        PdfRedactionVerificationReport report = PdfDocument.Load(rewritten).Redactions.VerifyAppliedPlan(
            plan,
            new PdfRedactionVerificationOptions { RequireCompleteStreamInspection = true });

        Assert.False(report.IsVerified);
        Assert.Contains(report.Issues, issue => issue.Feature == "RedactionPlanPageIdentityChanged");
    }

    [Fact]
    public void AppliedPlanVerificationRejectsReorderedPagesThatDifferOnlyByUnredactedAnnotations() {
        byte[] source = PdfDocument.Create(compose => {
            compose.Page(page => page.Content(content => content.Item(item => item.Paragraph(paragraph => paragraph.Text("Same page")))));
            compose.Page(page => page.Content(content => content.Item(item => item.Paragraph(paragraph => paragraph.Text("Same page")))));
        }).ToBytes();
        byte[] firstAnnotated = PdfDocument.Load(source).Annotations.Add(new PdfAnnotationCreateOptions {
            PageNumber = 1,
            Subtype = "Link",
            Rectangle = new[] { 100D, 100D, 120D, 120D },
            Contents = "Same annotation",
            LinkUri = "https://example.test/first"
        }).Bytes;
        byte[] annotated = PdfDocument.Load(firstAnnotated).Annotations.Add(new PdfAnnotationCreateOptions {
            PageNumber = 2,
            Subtype = "Link",
            Rectangle = new[] { 100D, 100D, 120D, 120D },
            Contents = "Same annotation",
            LinkUri = "https://example.test/second"
        }).Bytes;
        PdfRedactionPlan plan = PdfDocument.Load(annotated).Redactions.Plan([
            new PdfRedactionArea(1, 10D, 10D, 20D, 20D, "reviewed area")
        ]);
        byte[] rewritten = PdfPageExtractor.ExtractPages(annotated, 2, 1);

        PdfRedactionVerificationReport report = PdfDocument.Load(rewritten).Redactions.VerifyAppliedPlan(
            plan,
            new PdfRedactionVerificationOptions { RequireCompleteStreamInspection = true });

        Assert.False(report.IsVerified);
        Assert.Contains(report.Issues, issue => issue.Feature == "RedactionPlanPageIdentityChanged");
    }

    [Fact]
    public void AppliedPlanVerificationAcceptsPageObjectRenumbering() {
        byte[] source = BuildSparsePageObjectPdf();
        PdfRedactionPlan plan = PdfDocument.Load(source).Redactions.Plan([
            new PdfRedactionArea(1, 10D, 10D, 20D, 20D, "reviewed area")
        ]);

        byte[] rewritten = PdfRedactionApplier.Apply(source, plan);
        PdfRedactionVerificationReport report = PdfRedactionVerification.VerifyAppliedPlan(
            rewritten,
            plan,
            new PdfRedactionVerificationOptions { RequireCompleteStreamInspection = true });

        Assert.NotEqual(
            PdfReadDocument.Open(source).Pages[0].ObjectNumber,
            PdfReadDocument.Open(rewritten).Pages[0].ObjectNumber);
        Assert.True(report.IsVerified);
        Assert.DoesNotContain(report.Issues, issue => issue.Feature == "RedactionPlanPageIdentityChanged");
    }

    [Fact]
    public void AppliedPlanVerificationAcceptsAnEmptySearchDrivenPlan() {
        byte[] source = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Nothing confidential")).ToBytes();
        PdfRedactionPlan plan = PdfRedactionPlanner.Search(
            source,
            new PdfRedactionSearchOptions().AddLiteral("absent secret"));
        byte[] rewritten = PdfRedactionApplier.Apply(source, plan);

        PdfRedactionVerificationReport report = PdfRedactionVerification.VerifyAppliedPlan(
            rewritten,
            plan,
            new PdfRedactionVerificationOptions { RequireCompleteStreamInspection = true });

        Assert.True(plan.IsSearchDriven);
        Assert.Empty(plan.Areas);
        Assert.True(report.IsVerified);
    }

    [Fact]
    public void ReviewedPlanCannotBeAppliedToDifferentSourceBytes() {
        byte[] reviewed = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Reviewed source")).ToBytes();
        byte[] different = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Different source")).ToBytes();
        PdfRedactionPlan plan = PdfDocument.Load(reviewed).Redactions.Plan([
            new PdfRedactionArea(1, 0D, 0D, 600D, 800D, "whole page")
        ]);

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
            PdfRedactionApplier.Apply(different, plan));

        Assert.Contains("different source PDF bytes", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void AssertVerified_ConfirmsRemovedAndRetainedTextMarkersAfterApply() {
        PdfRedactionProofResult proof = PdfRedactionProofTestSupport.BuildAndVerifyRedactionRemovalProof();

        Assert.True(proof.Plan.HasMatches);
        Assert.Contains(proof.Plan.Matches, match => match.Text != null && match.Text.Contains("PAY-SECRET-2026", StringComparison.Ordinal));
        Assert.True(proof.Verification.IsVerified);
        Assert.True(proof.Verification.RawPdfBytesChecked);
        Assert.True(proof.Verification.EncodedPdfStringsChecked);
        Assert.True(proof.Verification.DecodedPdfStreamsChecked);
        Assert.Empty(proof.Verification.Issues);
        Assert.DoesNotContain("PAY-SECRET-2026", proof.Verification.ExtractedText, StringComparison.Ordinal);
        Assert.Contains("Visible compliance marker", proof.Verification.ExtractedText, StringComparison.Ordinal);
        Assert.Contains("Public summary marker", proof.Verification.ExtractedText, StringComparison.Ordinal);
    }

    [Fact]
    public void Verify_ReportsRemovedMarkersThatRemainInUnredactedPdf() {
        PdfRedactionProofResult proof = PdfRedactionProofTestSupport.BuildAndVerifyRedactionRemovalProof();

        PdfRedactionVerificationReport report = PdfRedactionVerification.Verify(
            proof.Source,
            PdfRedactionProofTestSupport.CreateVerificationOptions());

        Assert.False(report.IsVerified);
        Assert.Contains(report.Issues, issue => issue.Feature == "RemovedTextMarker" && issue.Marker == "PAY-SECRET-2026");
        Assert.DoesNotContain(report.Issues, issue => issue.Feature == "RetainedTextMarker");
        Assert.Contains("PDF redaction verification failed", report.Summary, StringComparison.Ordinal);

        var exception = Assert.Throws<InvalidOperationException>(() => report.ThrowIfFailed());
        Assert.Contains("PAY-SECRET-2026", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Verify_ReportsRemovedMarkersThatRemainInPdfHexStringBytes() {
        PdfRedactionProofResult proof = PdfRedactionProofTestSupport.BuildAndVerifyRedactionRemovalProof();
        byte[] rewrittenWithEncodedResidue = AppendPdfHexStringResidue(proof.Redacted, "PAY-SECRET-2026");

        PdfRedactionVerificationReport report = PdfRedactionVerification.Verify(
            rewrittenWithEncodedResidue,
            PdfRedactionProofTestSupport.CreateVerificationOptions());

        Assert.False(report.IsVerified);
        Assert.True(report.RawPdfBytesChecked);
        Assert.True(report.EncodedPdfStringsChecked);
        Assert.DoesNotContain("PAY-SECRET-2026", report.ExtractedText, StringComparison.Ordinal);
        Assert.DoesNotContain(report.Issues, issue => issue.Feature == "RemovedRawMarker" && issue.Marker == "PAY-SECRET-2026");
        Assert.Contains(report.Issues, issue => issue.Feature == "RemovedEncodedMarker" && issue.Marker == "PAY-SECRET-2026");
    }

    [Fact]
    public void Verify_CaseInsensitiveProfileFindsDifferentlyCasedEncodedResidue() {
        PdfRedactionProofResult proof = PdfRedactionProofTestSupport.BuildAndVerifyRedactionRemovalProof();
        byte[] rewrittenWithEncodedResidue = AppendPdfHexStringResidue(proof.Redacted, "PAY-SECRET-2026");
        var options = new PdfRedactionVerificationOptions { MatchCase = false };
        options.RequireRemovedText("pay-secret-2026");

        PdfRedactionVerificationReport report = PdfRedactionVerification.Verify(rewrittenWithEncodedResidue, options);

        Assert.False(report.IsVerified);
        Assert.Contains(report.Issues, issue => issue.Feature == "RemovedEncodedMarker" && issue.Marker == "pay-secret-2026");
    }

    [Fact]
    public void Verify_ReportsRemovedMarkersThatRemainInEscapedPdfLiteralBytes() {
        PdfRedactionProofResult proof = PdfRedactionProofTestSupport.BuildAndVerifyRedactionRemovalProof();
        byte[] rewrittenWithEncodedResidue = AppendPdfLiteralStringResidue(proof.Redacted, "PAY\\055SECRET\\0552026");

        PdfRedactionVerificationReport report = PdfRedactionVerification.Verify(
            rewrittenWithEncodedResidue,
            PdfRedactionProofTestSupport.CreateVerificationOptions());

        Assert.False(report.IsVerified);
        Assert.DoesNotContain("PAY-SECRET-2026", report.ExtractedText, StringComparison.Ordinal);
        Assert.DoesNotContain(report.Issues, issue => issue.Feature == "RemovedRawMarker" && issue.Marker == "PAY-SECRET-2026");
        Assert.Contains(report.Issues, issue => issue.Feature == "RemovedEncodedMarker" && issue.Marker == "PAY-SECRET-2026");
    }

    [Fact]
    public void Verify_CanSkipEncodedPdfStringResidueChecksWhenRequested() {
        PdfRedactionProofResult proof = PdfRedactionProofTestSupport.BuildAndVerifyRedactionRemovalProof();
        byte[] rewrittenWithEncodedResidue = AppendPdfHexStringResidue(proof.Redacted, "PAY-SECRET-2026");
        PdfRedactionVerificationOptions options = PdfRedactionProofTestSupport.CreateVerificationOptions();
        options.CheckEncodedPdfStrings = false;

        PdfRedactionVerificationReport report = PdfRedactionVerification.Verify(rewrittenWithEncodedResidue, options);

        Assert.True(report.IsVerified);
        Assert.False(report.EncodedPdfStringsChecked);
        Assert.DoesNotContain(report.Issues, issue => issue.Feature == "RemovedEncodedMarker");
    }

    [Fact]
    public void Verify_ReportsRemovedMarkersThatRemainInDecodedCompressedStreams() {
        PdfRedactionProofResult proof = PdfRedactionProofTestSupport.BuildAndVerifyRedactionRemovalProof();
        byte[] rewrittenWithCompressedResidue = AppendFlateStreamResidue(proof.Redacted, "PAY-SECRET-2026");

        PdfRedactionVerificationReport report = PdfRedactionVerification.Verify(
            rewrittenWithCompressedResidue,
            PdfRedactionProofTestSupport.CreateVerificationOptions());

        Assert.False(report.IsVerified);
        Assert.True(report.DecodedPdfStreamsChecked);
        Assert.DoesNotContain("PAY-SECRET-2026", report.ExtractedText, StringComparison.Ordinal);
        Assert.DoesNotContain(report.Issues, issue => issue.Feature == "RemovedRawMarker" && issue.Marker == "PAY-SECRET-2026");
        Assert.DoesNotContain(report.Issues, issue => issue.Feature == "RemovedEncodedMarker" && issue.Marker == "PAY-SECRET-2026");
        Assert.Contains(report.Issues, issue => issue.Feature == "RemovedDecodedStreamMarker" && issue.Marker == "PAY-SECRET-2026");
    }

    [Fact]
    public void Verify_CanSkipDecodedCompressedStreamResidueChecksWhenRequested() {
        PdfRedactionProofResult proof = PdfRedactionProofTestSupport.BuildAndVerifyRedactionRemovalProof();
        byte[] rewrittenWithCompressedResidue = AppendFlateStreamResidue(proof.Redacted, "PAY-SECRET-2026");
        PdfRedactionVerificationOptions options = PdfRedactionProofTestSupport.CreateVerificationOptions();
        options.CheckDecodedPdfStreams = false;

        PdfRedactionVerificationReport report = PdfRedactionVerification.Verify(rewrittenWithCompressedResidue, options);

        Assert.True(report.IsVerified);
        Assert.False(report.DecodedPdfStreamsChecked);
        Assert.DoesNotContain(report.Issues, issue => issue.Feature == "RemovedDecodedStreamMarker");
    }

    [Fact]
    public void Verify_CompleteStreamInspectionOverridesDisabledDecodedMarkerChecks() {
        PdfRedactionProofResult proof = PdfRedactionProofTestSupport.BuildAndVerifyRedactionRemovalProof();
        byte[] rewrittenWithCompressedResidue = AppendFlateStreamResidue(proof.Redacted, "PAY-SECRET-2026");
        PdfRedactionVerificationOptions options = PdfRedactionProofTestSupport.CreateVerificationOptions();
        options.CheckDecodedPdfStreams = false;
        options.RequireCompleteStreamInspection = true;

        PdfRedactionVerificationReport report = PdfRedactionVerification.Verify(rewrittenWithCompressedResidue, options);

        Assert.False(report.IsVerified);
        Assert.True(report.DecodedPdfStreamsChecked);
        Assert.Contains(report.Issues, issue => issue.Feature == "RemovedDecodedStreamMarker" && issue.Marker == "PAY-SECRET-2026");
    }

    [Fact]
    public void Verify_FailsClosedWhenPdfStreamCannotBeDecoded() {
        PdfRedactionProofResult proof = PdfRedactionProofTestSupport.BuildAndVerifyRedactionRemovalProof();
        byte[] rewrittenWithUndecodableStream = AppendUnsupportedFilteredStream(proof.Redacted);

        PdfRedactionVerificationReport report = PdfRedactionVerification.Verify(
            rewrittenWithUndecodableStream,
            PdfRedactionProofTestSupport.CreateVerificationOptions());

        Assert.False(report.IsVerified);
        Assert.True(report.DecodedPdfStreamsChecked);
        Assert.DoesNotContain("PAY-SECRET-2026", report.ExtractedText, StringComparison.Ordinal);
        Assert.Contains(report.Issues, issue => issue.Feature == "UndecodablePdfStream" && issue.Marker == "996");
    }

    [Fact]
    public void Verify_CanOptOutOfUndecodableStreamProofFailure() {
        PdfRedactionProofResult proof = PdfRedactionProofTestSupport.BuildAndVerifyRedactionRemovalProof();
        byte[] rewrittenWithUndecodableStream = AppendUnsupportedFilteredStream(proof.Redacted);
        PdfRedactionVerificationOptions options = PdfRedactionProofTestSupport.CreateVerificationOptions();
        options.FailOnUndecodablePdfStreams = false;

        PdfRedactionVerificationReport report = PdfRedactionVerification.Verify(rewrittenWithUndecodableStream, options);

        Assert.True(report.IsVerified);
        Assert.True(report.DecodedPdfStreamsChecked);
        Assert.DoesNotContain(report.Issues, issue => issue.Feature == "UndecodablePdfStream");
    }

    [Theory]
    [InlineData(false, true)]
    [InlineData(true, false)]
    public void Verify_CompleteStreamInspectionOverridesDisabledMarkerOrFailureSwitches(
        bool checkDecodedPdfStreams,
        bool failOnUndecodablePdfStreams) {
        PdfRedactionProofResult proof = PdfRedactionProofTestSupport.BuildAndVerifyRedactionRemovalProof();
        byte[] rewrittenWithUndecodableStream = AppendUnsupportedFilteredStream(proof.Redacted);
        PdfRedactionVerificationOptions options = PdfRedactionProofTestSupport.CreateVerificationOptions();
        options.CheckDecodedPdfStreams = checkDecodedPdfStreams;
        options.FailOnUndecodablePdfStreams = failOnUndecodablePdfStreams;
        options.RequireCompleteStreamInspection = true;

        PdfRedactionVerificationReport report = PdfRedactionVerification.Verify(rewrittenWithUndecodableStream, options);

        Assert.False(report.IsVerified);
        Assert.True(report.DecodedPdfStreamsChecked);
        Assert.Contains(report.Issues, issue => issue.Feature == "UndecodablePdfStream");
    }

    [Fact]
    public void Verify_CompleteStreamInspectionAcceptsOpaqueImageCodecStreams() {
        byte[] source = BuildJpegImageRedactionSource();

        PdfRedactionVerificationReport report = PdfRedactionVerification.Verify(
            source,
            new PdfRedactionVerificationOptions { RequireCompleteStreamInspection = true });

        Assert.True(report.IsVerified, report.Summary);
        Assert.DoesNotContain(report.Issues, issue => issue.Feature == "UndecodablePdfStream");
    }

    [Fact]
    public void Verify_CompleteStreamInspectionRejectsOpaqueImageCodecCombinedWithUnknownFilter() {
        byte[] source = BuildJpegImageRedactionSource("[/DCTDecode /UnknownDecode]");

        PdfRedactionVerificationReport report = PdfRedactionVerification.Verify(
            source,
            new PdfRedactionVerificationOptions { RequireCompleteStreamInspection = true });

        Assert.False(report.IsVerified);
        Assert.Contains(report.Issues, issue => issue.Feature == "UndecodablePdfStream");
    }

    [Theory]
    [InlineData("[/DCTDecode /JPXDecode]")]
    [InlineData("[/DCTDecode /DCTDecode]")]
    [InlineData("[/FlateDecode /DCTDecode]")]
    public void Verify_CompleteStreamInspectionRejectsMultipleOpaqueImageFilters(string filters) {
        byte[] source = BuildJpegImageRedactionSource(filters);

        PdfRedactionVerificationReport report = PdfRedactionVerification.Verify(
            source,
            new PdfRedactionVerificationOptions { RequireCompleteStreamInspection = true });

        Assert.False(report.IsVerified);
        Assert.Contains(report.Issues, issue => issue.Feature == "UndecodablePdfStream");
    }

    [Fact]
    public void ApplyWithEvidenceVerifiesTextRedactionAlongsideUntouchedJpeg() {
        byte[] source = BuildJpegImageRedactionSource();
        PdfDocument document = PdfDocument.Load(source);
        PdfRedactionPlan plan = document.Redactions.Search(
            new PdfRedactionSearchOptions().AddLiteral("REMOVE-JPEG-DOC-TEXT"));

        PdfRedactionApplyResult result = document.Redactions.ApplyWithEvidence(
            plan,
            verificationOptions: new PdfRedactionVerificationOptions {
                RequireCompleteStreamInspection = true
            }.RequireRemovedText("REMOVE-JPEG-DOC-TEXT"));

        Assert.True(result.IsVerified, result.Evidence.Summary);
        Assert.Single(PdfImageExtractor.ExtractImages(result.Pdf));
        Assert.Contains("/DCTDecode", PdfEncoding.Latin1GetString(result.Pdf), StringComparison.Ordinal);
    }

    [Fact]
    public void Plan_ReportsIntersectingImagePlacementsAsRedactionRisk() {
        byte[] source = BuildImageRedactionPlanningSource();
        PdfLogicalImage image = GetSingleImage(source);
        PdfImagePlacement placement = image.PrimaryPlacement!;
        PdfRedactionArea area = CreateImageIntersectionArea(image);

        PdfRedactionPlan plan = PdfRedactionPlanner.Plan(source, new[] { area }, new PdfTextLayoutOptions {
            ForceSingleColumn = true
        });

        PdfRedactionMatch match = Assert.Single(plan.Matches, item => item.Kind == PdfRedactionMatchKind.ImagePlacement);
        Assert.Equal(area, match.Area);
        Assert.Equal(placement.PageNumber, match.PageNumber);
        Assert.Equal(placement.X, match.X, 3);
        Assert.Equal(placement.Y, match.Y, 3);
        Assert.Equal(placement.Width, match.Width, 3);
        Assert.Equal(placement.Height, match.Height, 3);
        Assert.Equal(placement.ResourceName, match.ResourceName);
        Assert.Equal(placement.ObjectNumber, match.ObjectNumber);
        Assert.Null(match.Text);
        PdfDiagnosticFinding finding = Assert.Single(plan.Findings, finding =>
            finding.Code == "RedactionPlanImageIntersection" &&
            finding.Severity == PdfDiagnosticSeverity.Warning &&
            finding.PageNumber == image.PageNumber);
        Assert.Contains("rewrites supported image pixels", finding.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Apply_FailsClosedWhenRedactionAreaIntersectsImagePlacement() {
        byte[] source = BuildJpegImageRedactionSource();
        PdfLogicalImage image = GetSingleImage(source);
        PdfRedactionArea area = CreateImageIntersectionArea(image);

        var exception = Assert.Throws<InvalidOperationException>(() => PdfRedactionApplier.Apply(source, new[] { area }));

        Assert.Contains("intersects image placement", exception.Message, StringComparison.Ordinal);
        Assert.Contains("AllowImagePlacementOverlays", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Apply_CanSecurelyRemoveWholeJpegPlacementForPartialIntersection() {
        byte[] source = BuildJpegImageRedactionSource();
        PdfLogicalImage image = GetSingleImage(source);
        PdfRedactionArea area = CreateImageIntersectionArea(image);

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area }, new PdfRedactionApplyOptions { UnsupportedImagePolicy = PdfRedactionUnsupportedImagePolicy.RemoveWholePlacement });

        Assert.Empty(PdfImageExtractor.ExtractImages(redacted));
        Assert.DoesNotContain("/DCTDecode", PdfEncoding.Latin1GetString(redacted), StringComparison.Ordinal);
    }

    [Fact]
    public void Apply_RemovesFullyCoveredImagePlacementByDefault() {
        byte[] source = BuildImageRedactionPlanningSource();
        PdfLogicalImage image = GetSingleImage(source);
        PdfRedactionArea area = CreateImageCoveringArea(image);

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });

        string text = PdfTextExtractor.ExtractAllText(redacted);
        string raw = PdfEncoding.Latin1GetString(redacted);
        Assert.Contains("Visible image redaction planning marker", text, StringComparison.Ordinal);
        Assert.Contains("Retained text after image", text, StringComparison.Ordinal);
        Assert.Empty(PdfImageExtractor.ExtractImages(redacted));
        Assert.DoesNotContain("/Subtype /Image", raw, StringComparison.Ordinal);
        Assert.DoesNotContain("/Im", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void Apply_RemovesFullyCoveredNestedFormImagePlacement() {
        byte[] source = BuildNestedFormImageRedactionSource();
        PdfLogicalImage image = GetSingleImage(source);
        PdfRedactionArea area = CreateImageCoveringArea(image);

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });

        string raw = PdfEncoding.Latin1GetString(redacted);
        Assert.Empty(PdfImageExtractor.ExtractImages(redacted));
        Assert.DoesNotContain("/Subtype /Image", raw, StringComparison.Ordinal);
        Assert.DoesNotContain("/ImNested", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void Apply_HonorsConfiguredContentNestingLimitDuringImageCleanup() {
        string nestedOperand = new string('[', 129) + "0" + new string(']', 129);
        string pageContent = nestedOperand + " n\nq\n1 0 0 1 100 200 cm\n/Fx Do\nQ\n";
        const string formContent = "q\n10 0 0 10 0 0 cm\n/ImNested Do\nQ\n";
        byte[] source = BuildNestedImagePdf(pageContent, "<< /Fx 6 0 R >>", formContent, "ImNested");
        var readOptions = new PdfLoadOptions {
            Limits = new PdfReadLimits { MaxContentNestingDepth = 256 }
        };
        PdfImagePlacement placement = Assert.Single(
            PdfImageExtractor.ExtractImagePlacements(PdfReadDocument.Open(source, readOptions)));
        byte[] redacted = PdfRedactionApplier.RemoveImagePlacements(source, new[] { placement }, readOptions);

        Assert.Empty(PdfImageExtractor.ExtractImages(PdfReadDocument.Open(redacted, readOptions)));
    }

    [Fact]
    public void Apply_ClonesRepeatedFormInvocationBeforeRemovingNestedImagePlacement() {
        byte[] source = BuildRepeatedFormImageRedactionSource();
        PdfLogicalImage image = GetSingleImage(source);
        PdfImagePlacement firstPlacement = image.Placements.OrderBy(placement => placement.X).First();
        PdfRedactionArea area = CreateImageCoveringArea(image, firstPlacement);

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });

        IReadOnlyList<PdfImagePlacement> placements = PdfImageExtractor.ExtractImagePlacements(redacted);
        PdfImagePlacement remaining = Assert.Single(placements);
        Assert.Equal(120D, remaining.X, 3);
        Assert.Single(PdfImageExtractor.ExtractImages(redacted));
    }

    [Fact]
    public void Apply_RewritesPartiallyCoveredSimpleImagePixelsByDefault() {
        byte[] source = BuildSimpleFlateImageRedactionSource();
        PdfLogicalImage image = GetSingleImage(source);
        PdfImagePlacement placement = image.PrimaryPlacement!;
        PdfRedactionArea area = CreateImageLeftHalfArea(image, placement);

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });

        IReadOnlyList<PdfImagePlacement> placements = PdfImageExtractor.ExtractImagePlacements(redacted);
        PdfImagePlacement remainingPlacement = Assert.Single(placements);
        Assert.Equal(placement.X, remainingPlacement.X, 3);
        Assert.Equal(placement.Y, remainingPlacement.Y, 3);
        Assert.Single(PdfImageExtractor.ExtractImages(redacted));

        byte[] pixels = DecodeSingleImagePixels(redacted);
        Assert.Equal(24, pixels.Length);
        AssertRedactedLeftHalf(pixels, width: 4, height: 2, components: 3);
    }

    [Fact]
    public void ApplyWithEvidenceVerifiesAppliedPartialImagePixelRewrite() {
        byte[] source = BuildSimpleFlateImageRedactionSource();
        PdfLogicalImage image = GetSingleImage(source);
        PdfImagePlacement placement = image.PrimaryPlacement!;
        PdfRedactionArea area = CreateImageLeftHalfArea(image, placement);
        PdfDocument document = PdfDocument.Load(source);
        PdfRedactionPlan plan = document.Redactions.Plan([area]);

        PdfRedactionApplyResult result = document.Redactions.ApplyWithEvidence(
            plan,
            verificationOptions: new PdfRedactionVerificationOptions {
                RequireCompleteStreamInspection = true
            });

        Assert.True(result.IsVerified, result.Evidence.Summary);
        Assert.Empty(result.Evidence.ResidualMatches);
        PdfRedactionEvidenceItem item = Assert.Single(
            result.Evidence.Items,
            evidence => evidence.ReviewedMatch.Kind == PdfRedactionMatchKind.ImagePlacement);
        Assert.Equal(PdfRedactionEvidenceStatus.VerifiedAbsent, item.Status);
        AssertRedactedLeftHalf(DecodeSingleImagePixels(result.Pdf), width: 4, height: 2, components: 3);
    }

    [Fact]
    public void VerifyAppliedPlanWithoutMutationProofKeepsPartialImageResidualFailClosed() {
        byte[] source = BuildSimpleFlateImageRedactionSource();
        PdfLogicalImage image = GetSingleImage(source);
        PdfImagePlacement placement = image.PrimaryPlacement!;
        PdfRedactionArea area = CreateImageLeftHalfArea(image, placement);
        PdfRedactionPlan plan = PdfRedactionPlanner.Plan(source, [area]);
        byte[] redacted = PdfRedactionApplier.Apply(source, plan);

        PdfRedactionVerificationReport report = PdfRedactionVerification.VerifyAppliedPlan(
            redacted,
            plan,
            new PdfRedactionVerificationOptions { RequireCompleteStreamInspection = true });

        Assert.False(report.IsVerified);
        Assert.Contains(report.Issues, issue =>
            issue.Feature == "RedactionPlanResidual" &&
            issue.Marker == "ImagePlacement@page:1");
    }

    [Fact]
    public void ApplyWithEvidencePairsOneAppliedRewriteWithOneRepeatedImagePlacement() {
        byte[] source = BuildRepeatedSimpleFlateImageRedactionSource();
        PdfLogicalImage image = GetSingleImage(source);
        PdfImagePlacement firstPlacement = image.Placements.OrderBy(placement => placement.X).First();
        PdfRedactionArea area = CreateImageLeftHalfArea(image, firstPlacement);
        PdfDocument document = PdfDocument.Load(source);
        PdfRedactionPlan plan = document.Redactions.Plan([area]);

        PdfRedactionApplyResult result = document.Redactions.ApplyWithEvidence(
            plan,
            verificationOptions: new PdfRedactionVerificationOptions {
                RequireCompleteStreamInspection = true
            });

        Assert.True(result.IsVerified, result.Evidence.Summary);
        Assert.Empty(result.Evidence.ResidualMatches);
        PdfImagePlacement[] placements = PdfImageExtractor.ExtractImagePlacements(result.Pdf).ToArray();
        Assert.Equal(2, placements.Length);
        byte[][] images = DecodeImagePixelStreams(result.Pdf);
        Assert.Contains(images, pixels => PixelRowsMatch(pixels, CreateSimpleFlateImagePixels()));
        Assert.Contains(images, pixels => LeftHalfIsRedacted(pixels, width: 4, height: 2, components: 3));
    }

    [Fact]
    public void ApplyWithEvidenceRecordsEachIdenticalImageInvocationMutation() {
        byte[] source = BuildRepeatedIdenticalSimpleFlateImageRedactionSource();
        PdfLogicalImage image = GetSingleImage(source);
        PdfImagePlacement placement = image.PrimaryPlacement!;
        PdfRedactionArea area = CreateImageLeftHalfArea(image, placement);
        PdfDocument document = PdfDocument.Load(source);
        PdfRedactionPlan plan = document.Redactions.Plan([area]);
        Assert.Equal(2, plan.Matches.Count(match => match.Kind == PdfRedactionMatchKind.ImagePlacement));

        PdfRedactionApplyResult result = document.Redactions.ApplyWithEvidence(
            plan,
            verificationOptions: new PdfRedactionVerificationOptions {
                RequireCompleteStreamInspection = true
            });

        Assert.True(result.IsVerified, result.Evidence.Summary);
        Assert.Empty(result.Evidence.ResidualMatches);
        Assert.Equal(2, PdfImageExtractor.ExtractImagePlacements(result.Pdf).Count);
        byte[][] images = DecodeImagePixelStreams(result.Pdf);
        Assert.Equal(2, images.Length);
        Assert.All(images, pixels => Assert.True(LeftHalfIsRedacted(pixels, width: 4, height: 2, components: 3)));
    }

    [Fact]
    public void ApplyWithEvidenceAccountsForInPlacePageStreamGrowthFromImageAliases() {
        byte[] source = BuildManyIdenticalSimpleFlateImageRedactionSource();
        Dictionary<int, PdfIndirectObject> objects = PdfSyntax.ParseObjects(source).Map;
        PdfStream[] sourceStreams = objects.Values
            .Select(static item => item.Value)
            .OfType<PdfStream>()
            .ToArray();
        byte[][] decodedStreams = sourceStreams
            .Select(stream => StreamDecoder.Decode(stream.Dictionary, stream.Data, objects, int.MaxValue))
            .ToArray();
        PdfStream pageContentStream = Assert.IsType<PdfStream>(objects[4].Value);
        int sourcePageContentBytes = StreamDecoder.Decode(pageContentStream.Dictionary, pageContentStream.Data, objects, int.MaxValue).Length;
        var readOptions = new PdfLoadOptions {
            Limits = new PdfReadLimits {
                MaxInputBytes = source.LongLength,
                MaxRawStreamBytes = sourceStreams.Max(static stream => stream.Data.Length),
                MaxDecodedStreamBytes = decodedStreams.Max(static stream => stream.Length),
                MaxTotalDecodedStreamBytes = decodedStreams.Sum(static stream => (long)stream.LongLength),
                MaxPageContentBytes = sourcePageContentBytes,
                MaxRetainedContentBytes = sourcePageContentBytes
            }
        };
        PdfDocument document = PdfDocument.Load(source, readOptions);
        PdfLogicalImage image = GetSingleImage(source);
        PdfRedactionArea area = CreateImageLeftHalfArea(image, image.PrimaryPlacement!);
        PdfRedactionPlan plan = document.Redactions.Plan([area]);
        Assert.Equal(64, plan.Matches.Count(match => match.Kind == PdfRedactionMatchKind.ImagePlacement));

        PdfRedactionApplyResult result = document.Redactions.ApplyWithEvidence(
            plan,
            verificationOptions: new PdfRedactionVerificationOptions {
                RequireCompleteStreamInspection = true,
                CheckManagedRendering = false
            });

        Assert.True(result.IsVerified, result.Evidence.Summary);
        Assert.Equal(64, PdfImageExtractor.ExtractImagePlacements(result.Pdf).Count);
        Assert.DoesNotContain("/ASCIIHexDecode", PdfEncoding.Latin1GetString(result.Pdf), StringComparison.Ordinal);
    }

    [Fact]
    public void Apply_RewritesPartiallyCoveredSoftMaskedSimpleImagePixelsAndMask() {
        byte[] source = BuildSoftMaskedSimpleFlateImageRedactionSource();
        PdfLogicalImage image = GetSingleImage(source);
        PdfImagePlacement placement = image.PrimaryPlacement!;
        PdfRedactionArea area = CreateImageLeftHalfArea(image, placement);

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });

        PdfExtractedImage extracted = Assert.Single(PdfImageExtractor.ExtractImages(redacted));
        Assert.True(extracted.HasTransparencyMask);
        Assert.True(extracted.TransparencyMaskResolved);
        Assert.Equal("soft-mask", extracted.TransparencyMaskKind);
        AssertRedactedLeftHalf(DecodeSingleImagePixels(redacted), width: 4, height: 2, components: 3);
        AssertSoftMaskLeftHalfOpaque(DecodeSoftMaskPixels(redacted, extracted.ObjectNumber), width: 4, height: 2);
    }

    [Fact]
    public void Apply_RewritesPartiallyCoveredDecodeAwareImagePixelsAndMask() {
        byte[] source = BuildDecodeAwareSoftMaskedSimpleFlateImageRedactionSource();
        PdfLogicalImage image = GetSingleImage(source);
        PdfImagePlacement placement = image.PrimaryPlacement!;
        PdfRedactionArea area = CreateImageLeftHalfArea(image, placement);

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });

        PdfExtractedImage extracted = Assert.Single(PdfImageExtractor.ExtractImages(redacted));
        Assert.True(extracted.HasTransparencyMask);
        Assert.True(extracted.TransparencyMaskResolved);
        AssertInvertedDecodeRedactedLeftHalf(DecodeSingleImagePixels(redacted), width: 4, height: 2, components: 3);
        AssertInvertedDecodeSoftMaskLeftHalfOpaque(DecodeSoftMaskPixels(redacted, extracted.ObjectNumber), width: 4, height: 2);
    }

    [Fact]
    public void Apply_ClonesRepeatedImageInvocationBeforeRewritingPixels() {
        byte[] source = BuildRepeatedSimpleFlateImageRedactionSource();
        PdfLogicalImage image = GetSingleImage(source);
        PdfImagePlacement firstPlacement = image.Placements.OrderBy(placement => placement.X).First();
        PdfRedactionArea area = CreateImageLeftHalfArea(image, firstPlacement);

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });

        string raw = PdfEncoding.Latin1GetString(redacted);
        Assert.Contains("/ImSimpleRedacted1 Do", raw, StringComparison.Ordinal);
        Assert.Contains("/ImSimple Do", raw, StringComparison.Ordinal);
        Assert.Equal(2, PdfImageExtractor.ExtractImagePlacements(redacted).Count);
        byte[][] images = DecodeImagePixelStreams(redacted);
        Assert.Equal(2, images.Length);
        Assert.Contains(images, pixels => PixelRowsMatch(pixels, CreateSimpleFlateImagePixels()));
        Assert.Contains(images, pixels => LeftHalfIsRedacted(pixels, width: 4, height: 2, components: 3));
    }

    [Fact]
    public void Apply_ClonesSharedFormResourceBeforeRemovingNestedImagePlacement() {
        byte[] source = BuildSharedAliasFormImageRedactionSource();
        PdfLogicalImage image = GetSingleImage(source);
        PdfImagePlacement firstPlacement = image.Placements.OrderBy(placement => placement.X).First();
        PdfRedactionArea area = CreateImageCoveringArea(image, firstPlacement);

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });
        IReadOnlyList<PdfImagePlacement> placements = PdfImageExtractor.ExtractImagePlacements(redacted);

        PdfImagePlacement remaining = Assert.Single(placements);
        Assert.Equal(120D, remaining.X, 3);
        Assert.Single(PdfImageExtractor.ExtractImages(redacted));
    }

    [Fact]
    public void Apply_PreservesPriorSharedContentArrayImageReplacements() {
        byte[] source = BuildSharedIndirectContentArrayImageRedactionSource();
        IReadOnlyList<PdfImagePlacement> sourcePlacements = PdfImageExtractor.ExtractImagePlacements(source);
        Assert.Equal(4, sourcePlacements.Count);
        var area = new PdfRedactionArea(1, 0, 0, 200, 120, "shared-content-images");

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });

        IReadOnlyList<PdfImagePlacement> placements = PdfImageExtractor.ExtractImagePlacements(redacted);
        Assert.DoesNotContain(placements, placement => placement.PageNumber == 1);
        Assert.Contains(placements, placement => placement.PageNumber == 2);
    }

    [Fact]
    public void Apply_PreservesSharedImageAliasesWhenAContentOwnerCannotBeDecoded() {
        const string pageContent = "q 20 0 0 20 20 30 cm /ImShared Do Q";
        const string undecodableFormContent = "/ImShared Do";
        const string imageBytes = "abc";
        byte[] source = Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 120] /Resources << /XObject << /ImShared 6 0 R /Fx 7 0 R >> >> /Contents 4 0 R >>", "endobj",
            "4 0 obj", "<< /Length " + pageContent.Length.ToString(CultureInfo.InvariantCulture) + " >>", "stream", pageContent, "endstream", "endobj",
            "6 0 obj", "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Length " + imageBytes.Length.ToString(CultureInfo.InvariantCulture) + " >>", "stream", imageBytes, "endstream", "endobj",
            "7 0 obj", "<< /Type /XObject /Subtype /Form /BBox [0 0 1 1] /Resources << /XObject << /ImShared 6 0 R >> >> /Filter /RunLengthDecode /Length " + undecodableFormContent.Length.ToString(CultureInfo.InvariantCulture) + " >>", "stream", undecodableFormContent, "endstream", "endobj",
            "trailer", "<< /Root 1 0 R /Size 8 >>", "%%EOF"
        }));
        PdfLogicalImage image = GetSingleImage(source);
        PdfRedactionArea area = CreateImageCoveringArea(image, Assert.Single(image.Placements));

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });
        Dictionary<int, PdfIndirectObject> objects = PdfSyntax.ParseObjects(redacted).Map;
        PdfStream form = Assert.IsType<PdfStream>(Assert.Single(objects.Values, static item =>
            item.Value is PdfStream stream && stream.Dictionary.Get<PdfName>("Subtype")?.Name == "Form").Value);
        PdfDictionary resources = Assert.IsType<PdfDictionary>(form.Dictionary.Items["Resources"]);
        PdfDictionary xObjects = Assert.IsType<PdfDictionary>(resources.Items["XObject"]);

        Assert.True(xObjects.Items.ContainsKey("ImShared"));
    }

    [Fact]
    public void Apply_AllowsExplicitImageOverlayWhenWeakerOutcomeIsAccepted() {
        byte[] source = BuildImageRedactionPlanningSource();
        PdfLogicalImage image = GetSingleImage(source);
        PdfRedactionArea area = CreateImageIntersectionArea(image);

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area }, new PdfRedactionApplyOptions {
            AllowImagePlacementOverlays = true
        });

        Assert.Contains("Retained text after image", PdfTextExtractor.ExtractAllText(redacted), StringComparison.Ordinal);
        Assert.Single(PdfImageExtractor.ExtractImages(redacted));
    }

    private static byte[] BuildImageRedactionPlanningSource() {
        return PdfDocument.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .Paragraph(paragraph => paragraph.Text("Visible image redaction planning marker"))
            .Image(PdfPngTestImages.CreateRgbPng(3, 2), 48, 32, alternativeText: "Sensitive chart pixels")
            .Paragraph(paragraph => paragraph.Text("Retained text after image"))
            .ToBytes();
    }

    private static byte[] BuildImageIdentitySource(bool includeInvertedDecode) {
        const string pageContent = "q\n20 0 0 20 20 30 cm\n/Im1 Do\nQ\n";
        const string imageBytes = "abc";
        string decode = includeInvertedDecode ? " /Decode [1 0 1 0 1 0]" : string.Empty;
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 120] /Resources << /XObject << /Im1 5 0 R >> >> /Contents 4 0 R >>", "endobj",
            "4 0 obj", "<< /Length " + Encoding.ASCII.GetByteCount(pageContent).ToString(CultureInfo.InvariantCulture) + " >>", "stream", pageContent, "endstream", "endobj",
            "5 0 obj", "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8" + decode + " /Length 3 >>", "stream", imageBytes, "endstream", "endobj",
            "trailer", "<< /Root 1 0 R /Size 6 >>", "%%EOF"
        }) + "\n";
        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildNestedFormImageRedactionSource() {
        const string pageContent = "q\n1 0 0 1 100 200 cm\n/Fx Do\nQ\n";
        const string formContent = "q\n10 0 0 10 0 0 cm\n/ImNested Do\nQ\n";
        return BuildNestedImagePdf(pageContent, "<< /Fx 6 0 R >>", formContent, "ImNested");
    }

    private static byte[] BuildRepeatedFormImageRedactionSource() {
        const string pageContent = "q\n1 0 0 1 20 30 cm\n/Fx Do\nQ\nq\n1 0 0 1 120 30 cm\n/Fx Do\nQ\n";
        const string formContent = "q\n10 0 0 10 0 0 cm\n/ImShared Do\nQ\n";
        return BuildNestedImagePdf(pageContent, "<< /Fx 6 0 R >>", formContent, "ImShared");
    }

    private static byte[] BuildSharedAliasFormImageRedactionSource() {
        const string pageContent = "q\n1 0 0 1 20 30 cm\n/FxA Do\nQ\nq\n1 0 0 1 120 30 cm\n/FxB Do\nQ\n";
        const string formContent = "q\n10 0 0 10 0 0 cm\n/ImShared Do\nQ\n";
        return BuildNestedImagePdf(pageContent, "<< /FxA 6 0 R /FxB 6 0 R >>", formContent, "ImShared");
    }

    private static byte[] BuildSharedIndirectContentArrayImageRedactionSource() {
        const string firstContent = "q\n20 0 0 20 20 30 cm\n/ImSharedA Do\nQ\n";
        const string secondContent = "q\n20 0 0 20 80 30 cm\n/ImSharedB Do\nQ\n";
        const string imageBytes = "abc";

        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 2 /Kids [3 0 R 4 0 R] /MediaBox [0 0 200 120] /Resources << /XObject << /ImSharedA 9 0 R /ImSharedB 11 0 R >> >> >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /Contents 8 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Type /Page /Parent 2 0 R /Contents 10 0 R >>",
            "endobj",
            "5 0 obj",
            "<< /Length " + Encoding.ASCII.GetByteCount(firstContent).ToString(CultureInfo.InvariantCulture) + " >>",
            "stream",
            firstContent,
            "endstream",
            "endobj",
            "6 0 obj",
            "<< /Length " + Encoding.ASCII.GetByteCount(secondContent).ToString(CultureInfo.InvariantCulture) + " >>",
            "stream",
            secondContent,
            "endstream",
            "endobj",
            "8 0 obj",
            "[5 0 R 6 0 R]",
            "endobj",
            "9 0 obj",
            "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Length " + Encoding.ASCII.GetByteCount(imageBytes).ToString(CultureInfo.InvariantCulture) + " >>",
            "stream",
            imageBytes,
            "endstream",
            "endobj",
            "10 0 obj",
            "[5 0 R 6 0 R]",
            "endobj",
            "11 0 obj",
            "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Length " + Encoding.ASCII.GetByteCount(imageBytes).ToString(CultureInfo.InvariantCulture) + " >>",
            "stream",
            imageBytes,
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildNestedImagePdf(string pageContent, string pageXObjects, string formContent, string imageName) {
        const string imageBytes = "abc";
        int pageStreamLength = Encoding.ASCII.GetByteCount(pageContent);
        int formStreamLength = Encoding.ASCII.GetByteCount(formContent);
        int imageLength = Encoding.ASCII.GetByteCount(imageBytes);

        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] /Resources 4 0 R >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /Contents 5 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /XObject " + pageXObjects + " >>",
            "endobj",
            "5 0 obj",
            "<< /Length " + pageStreamLength.ToString(CultureInfo.InvariantCulture) + " >>",
            "stream",
            pageContent.TrimEnd('\n'),
            "endstream",
            "endobj",
            "6 0 obj",
            "<< /Type /XObject /Subtype /Form /BBox [0 0 20 20] /Resources << /XObject << /" + imageName + " 7 0 R >> >> /Length " + formStreamLength.ToString(CultureInfo.InvariantCulture) + " >>",
            "stream",
            formContent.TrimEnd('\n'),
            "endstream",
            "endobj",
            "7 0 obj",
            "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Length " + imageLength.ToString(CultureInfo.InvariantCulture) + " >>",
            "stream",
            imageBytes,
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildSimpleFlateImageRedactionSource() {
        const string pageContent = "q\n40 0 0 20 20 30 cm\n/ImSimple Do\nQ\n";
        return BuildSimpleFlateImagePdf(pageContent);
    }

    private static byte[] BuildTextPaintIdentityPdf(string colorOperator) {
        string content = $"BT /F1 12 Tf {colorOperator} 20 100 Td (Visible text) Tj ET";
        return Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>", "endobj",
            "4 0 obj", $"<< /Length {content.Length.ToString(CultureInfo.InvariantCulture)} >>", "stream", content, "endstream", "endobj",
            "5 0 obj", "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>", "endobj",
            "trailer", "<< /Root 1 0 R /Size 6 >>", "%%EOF"
        }));
    }

    private static byte[] BuildTextVisualIdentityPdf(string paintOperators, string textMatrix) {
        string content = $"BT /F1 12 Tf {paintOperators} {textMatrix} Tm (Visible text) Tj ET";
        return Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>", "endobj",
            "4 0 obj", $"<< /Length {Encoding.ASCII.GetByteCount(content).ToString(CultureInfo.InvariantCulture)} >>", "stream", content, "endstream", "endobj",
            "5 0 obj", "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>", "endobj",
            "trailer", "<< /Root 1 0 R /Size 6 >>", "%%EOF"
        }));
    }

    private static byte[] BuildEmbeddedFontIdentityPdf(string fontProgram) {
        const string content = "BT /F1 12 Tf 20 100 Td (Visible text) Tj ET";
        return Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>", "endobj",
            "4 0 obj", $"<< /Length {Encoding.ASCII.GetByteCount(content).ToString(CultureInfo.InvariantCulture)} >>", "stream", content, "endstream", "endobj",
            "5 0 obj", "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /FontDescriptor 6 0 R >>", "endobj",
            "6 0 obj", "<< /Type /FontDescriptor /FontName /Helvetica /Flags 32 /FontBBox [0 -200 1000 900] /ItalicAngle 0 /Ascent 800 /Descent -200 /CapHeight 700 /StemV 80 /FontFile 7 0 R >>", "endobj",
            "7 0 obj", $"<< /Length {Encoding.ASCII.GetByteCount(fontProgram).ToString(CultureInfo.InvariantCulture)} >>", "stream", fontProgram, "endstream", "endobj",
            "trailer", "<< /Root 1 0 R /Size 8 >>", "%%EOF"
        }));
    }

    private static byte[] BuildTextOpacityIdentityPdf(string selectedState) {
        string content = $"/{selectedState} gs BT /F1 12 Tf 2 Tr 20 100 Td (Visible text) Tj ET";
        return Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Resources << /Font << /F1 5 0 R >> /ExtGState << /GS1 << /CA 0.25 >> /GS2 << /CA 0.75 >> >> >> /Contents 4 0 R >>", "endobj",
            "4 0 obj", $"<< /Length {Encoding.ASCII.GetByteCount(content).ToString(CultureInfo.InvariantCulture)} >>", "stream", content, "endstream", "endobj",
            "5 0 obj", "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>", "endobj",
            "trailer", "<< /Root 1 0 R /Size 6 >>", "%%EOF"
        }));
    }

    private static byte[] BuildRotatedTextIdentityPdf() {
        const string content = "BT /F1 12 Tf 0 1 -1 0 100 40 Tm (VERTICAL) Tj ET";
        return Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>", "endobj",
            "4 0 obj", $"<< /Length {content.Length.ToString(CultureInfo.InvariantCulture)} >>", "stream", content, "endstream", "endobj",
            "5 0 obj", "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>", "endobj",
            "trailer", "<< /Root 1 0 R /Size 6 >>", "%%EOF"
        }));
    }

    private static byte[] BuildVectorClipIdentityPdf(string clipOperators) {
        string content = $"q {clipOperators} 1 0 0 rg 20 20 80 60 re f Q";
        return Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>", "endobj",
            "4 0 obj", $"<< /Length {Encoding.ASCII.GetByteCount(content).ToString(CultureInfo.InvariantCulture)} >>", "stream", content, "endstream", "endobj",
            "trailer", "<< /Root 1 0 R /Size 5 >>", "%%EOF"
        }));
    }

    private static byte[] BuildNestedFormVectorPathPdf() {
        const string pageContent = "q 1 0 0 1 50 50 cm /Fm1 Do Q";
        const string formContent = "1 0 0 rg 0 0 40 40 re f";
        return Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Resources << /XObject << /Fm1 5 0 R >> >> /Contents 4 0 R >>", "endobj",
            "4 0 obj", $"<< /Length {Encoding.ASCII.GetByteCount(pageContent).ToString(CultureInfo.InvariantCulture)} >>", "stream", pageContent, "endstream", "endobj",
            "5 0 obj", $"<< /Type /XObject /Subtype /Form /BBox [0 0 40 40] /Length {Encoding.ASCII.GetByteCount(formContent).ToString(CultureInfo.InvariantCulture)} >>", "stream", formContent, "endstream", "endobj",
            "trailer", "<< /Root 1 0 R /Size 6 >>", "%%EOF"
        }));
    }

    private static byte[] BuildVectorStyleIdentityPdf(string styleOperators) {
        string content = $"q {styleOperators} 1 0 0 RG 4 w 20 20 80 60 re S Q";
        return Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>", "endobj",
            "4 0 obj", $"<< /Length {Encoding.ASCII.GetByteCount(content).ToString(CultureInfo.InvariantCulture)} >>", "stream", content, "endstream", "endobj",
            "trailer", "<< /Root 1 0 R /Size 5 >>", "%%EOF"
        }));
    }

    private static byte[] BuildVectorEffectIdentityPdf(string selectedState) {
        string content = $"q /{selectedState} gs 1 0 0 rg 20 20 80 60 re f Q";
        const string maskOne = "0 g 0 0 20 20 re f";
        const string maskTwo = "1 g 0 0 20 20 re f";
        return Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Resources << /ExtGState << /GS1 << /BM /Multiply >> /GS2 << /BM /Screen >> /GS3 << /SMask << /S /Alpha /G 5 0 R >> >> /GS4 << /SMask << /S /Alpha /G 6 0 R >> >> >> >> /Contents 4 0 R >>", "endobj",
            "4 0 obj", $"<< /Length {Encoding.ASCII.GetByteCount(content).ToString(CultureInfo.InvariantCulture)} >>", "stream", content, "endstream", "endobj",
            "5 0 obj", $"<< /Type /XObject /Subtype /Form /BBox [0 0 20 20] /Group << /S /Transparency >> /Length {Encoding.ASCII.GetByteCount(maskOne).ToString(CultureInfo.InvariantCulture)} >>", "stream", maskOne, "endstream", "endobj",
            "6 0 obj", $"<< /Type /XObject /Subtype /Form /BBox [0 0 20 20] /Group << /S /Transparency >> /Length {Encoding.ASCII.GetByteCount(maskTwo).ToString(CultureInfo.InvariantCulture)} >>", "stream", maskTwo, "endstream", "endobj",
            "trailer", "<< /Root 1 0 R /Size 7 >>", "%%EOF"
        }));
    }

    private static byte[] BuildTilingPatternIdentityPdf(string tileContent, double horizontalStep) {
        const string content = "/Pattern cs /P1 scn 20 20 80 60 re f";
        return Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Resources << /Pattern << /P1 5 0 R >> >> /Contents 4 0 R >>", "endobj",
            "4 0 obj", $"<< /Length {Encoding.ASCII.GetByteCount(content).ToString(CultureInfo.InvariantCulture)} >>", "stream", content, "endstream", "endobj",
            "5 0 obj", $"<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 10 10] /XStep {horizontalStep.ToString("R", CultureInfo.InvariantCulture)} /YStep 10 /Resources << >> /Length {Encoding.ASCII.GetByteCount(tileContent).ToString(CultureInfo.InvariantCulture)} >>", "stream", tileContent, "endstream", "endobj",
            "trailer", "<< /Root 1 0 R /Size 6 >>", "%%EOF"
        }));
    }

    private static byte[] BuildAnnotationAppearanceIdentityPdf(string colorOperator, string annotationStyle = "") {
        string appearance = $"q {colorOperator} 0 0 40 40 re f Q";
        return Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Annots [5 0 R] /Contents 4 0 R >>", "endobj",
            "4 0 obj", "<< /Length 0 >>", "stream", "", "endstream", "endobj",
            "5 0 obj", $"<< /Type /Annot /Subtype /Stamp /Rect [100 100 140 140] /NM (appearance-proof) {annotationStyle} /AP << /N 6 0 R >> >>", "endobj",
            "6 0 obj", $"<< /Type /XObject /Subtype /Form /BBox [0 0 40 40] /Length {appearance.Length.ToString(CultureInfo.InvariantCulture)} >>", "stream", appearance, "endstream", "endobj",
            "trailer", "<< /Root 1 0 R /Size 7 >>", "%%EOF"
        }));
    }

    private static byte[] BuildSparsePageObjectPdf() {
        const string content = "q Q";
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [50 0 R] >>",
            "endobj",
            "50 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 120] /Contents 75 0 R >>",
            "endobj",
            "75 0 obj",
            "<< /Length 3 >>",
            "stream",
            content,
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R >>",
            "%%EOF"
        }) + "\n";
        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildTextObjectWithTwoSpansPdf() {
        const string content = "BT /F1 12 Tf 10 100 Td (FIRST) Tj 90 0 Td (SECOND) Tj ET";
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 240 200] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length " + Encoding.ASCII.GetByteCount(content).ToString(CultureInfo.InvariantCulture) + " >>",
            "stream",
            content,
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R >>",
            "%%EOF"
        }) + "\n";
        return Encoding.ASCII.GetBytes(pdf);
    }

    private static void AssertPlanIdentityChanged(byte[] source, byte[] rewritten) {
        PdfRedactionPlan plan = PdfRedactionPlanner.Plan(source, [
            new PdfRedactionArea(1, 150D, 20D, 10D, 10D, "reviewed blank area")
        ]);

        PdfRedactionVerificationReport report = PdfRedactionVerification.VerifyAppliedPlan(
            rewritten,
            plan,
            new PdfRedactionVerificationOptions { RequireCompleteStreamInspection = true });

        Assert.False(report.IsVerified);
        Assert.Contains(report.Issues, static issue => issue.Feature == "RedactionPlanPageIdentityChanged");
    }

    private static byte[] BuildRepeatedSimpleFlateImageRedactionSource() {
        const string pageContent = "q\n40 0 0 20 20 30 cm\n/ImSimple Do\nQ\nq\n40 0 0 20 80 30 cm\n/ImSimple Do\nQ\n";
        return BuildSimpleFlateImagePdf(pageContent);
    }

    private static byte[] BuildRepeatedIdenticalSimpleFlateImageRedactionSource() {
        const string pageContent = "q\n40 0 0 20 20 30 cm\n/ImSimple Do\nQ\nq\n40 0 0 20 20 30 cm\n/ImSimple Do\nQ\n";
        return BuildSimpleFlateImagePdf(pageContent);
    }

    private static byte[] BuildManyIdenticalSimpleFlateImageRedactionSource() {
        string pageContent = string.Concat(Enumerable.Range(0, 64)
            .Select(static _ => "q\n40 0 0 20 20 30 cm\n/ImSimple Do\nQ\n"));
        byte[] pixels = CreateSimpleFlateImagePixels();
        byte[] compressed = Compress(pixels);
        string encodedPageContent = ToHex(Encoding.ASCII.GetBytes(pageContent.TrimEnd('\n'))) + ">";

        using var output = new MemoryStream();
        void WriteAscii(string text) {
            byte[] bytes = Encoding.ASCII.GetBytes(text);
            output.Write(bytes, 0, bytes.Length);
        }

        WriteAscii(string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 200 120] /Resources << /XObject << /ImSimple 5 0 R >> >> >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length " + encodedPageContent.Length.ToString(CultureInfo.InvariantCulture) + " /Filter 6 0 R >>",
            "stream",
            encodedPageContent,
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /XObject /Subtype /Image /Width 4 /Height 2 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /FlateDecode /Length " + compressed.Length.ToString(CultureInfo.InvariantCulture) + " >>",
            "stream"
        }) + "\n");
        output.Write(compressed, 0, compressed.Length);
        WriteAscii("\nendstream\nendobj\n6 0 obj\n/ASCIIHexDecode\nendobj\ntrailer\n<< /Root 1 0 R >>\n%%EOF\n");
        return output.ToArray();
    }

    private static byte[] BuildSoftMaskedSimpleFlateImageRedactionSource() {
        const string pageContent = "q\n40 0 0 20 20 30 cm\n/ImSoft Do\nQ\n";
        byte[] pixels = CreateSimpleFlateImagePixels();
        byte[] mask = new byte[] { 64, 64, 128, 128, 192, 192, 224, 224 };
        byte[] compressedPixels = Compress(pixels);
        byte[] compressedMask = Compress(mask);
        int pageStreamLength = Encoding.ASCII.GetByteCount(pageContent.TrimEnd('\n'));

        using var output = new MemoryStream();
        void WriteAscii(string text) {
            byte[] bytes = Encoding.ASCII.GetBytes(text);
            output.Write(bytes, 0, bytes.Length);
        }

        WriteAscii(string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 200 120] /Resources << /XObject << /ImSoft 5 0 R >> >> >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length " + pageStreamLength.ToString(CultureInfo.InvariantCulture) + " >>",
            "stream",
            pageContent.TrimEnd('\n'),
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /XObject /Subtype /Image /Width 4 /Height 2 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /FlateDecode /SMask 6 0 R /Length " + compressedPixels.Length.ToString(CultureInfo.InvariantCulture) + " >>",
            "stream"
        }) + "\n");
        output.Write(compressedPixels, 0, compressedPixels.Length);
        WriteAscii("\nendstream\nendobj\n6 0 obj\n<< /Type /XObject /Subtype /Image /Width 4 /Height 2 /ColorSpace /DeviceGray /BitsPerComponent 8 /Filter /FlateDecode /Length " + compressedMask.Length.ToString(CultureInfo.InvariantCulture) + " >>\nstream\n");
        output.Write(compressedMask, 0, compressedMask.Length);
        WriteAscii("\nendstream\nendobj\ntrailer\n<< /Root 1 0 R >>\n%%EOF\n");
        return output.ToArray();
    }

    private static byte[] BuildDecodeAwareSoftMaskedSimpleFlateImageRedactionSource() {
        const string pageContent = "q\n40 0 0 20 20 30 cm\n/ImDecoded Do\nQ\n";
        byte[] pixels = CreateSimpleFlateImagePixels();
        byte[] mask = new byte[] { 64, 64, 128, 128, 192, 192, 224, 224 };
        byte[] compressedPixels = Compress(pixels);
        byte[] compressedMask = Compress(mask);
        int pageStreamLength = Encoding.ASCII.GetByteCount(pageContent.TrimEnd('\n'));

        using var output = new MemoryStream();
        void WriteAscii(string text) {
            byte[] bytes = Encoding.ASCII.GetBytes(text);
            output.Write(bytes, 0, bytes.Length);
        }

        WriteAscii(string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 200 120] /Resources << /XObject << /ImDecoded 5 0 R >> >> >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length " + pageStreamLength.ToString(CultureInfo.InvariantCulture) + " >>",
            "stream",
            pageContent.TrimEnd('\n'),
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /XObject /Subtype /Image /Width 4 /Height 2 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Decode [1 0 1 0 1 0] /Filter /FlateDecode /SMask 6 0 R /Length " + compressedPixels.Length.ToString(CultureInfo.InvariantCulture) + " >>",
            "stream"
        }) + "\n");
        output.Write(compressedPixels, 0, compressedPixels.Length);
        WriteAscii("\nendstream\nendobj\n6 0 obj\n<< /Type /XObject /Subtype /Image /Width 4 /Height 2 /ColorSpace /DeviceGray /BitsPerComponent 8 /Decode [1 0] /Filter /FlateDecode /Length " + compressedMask.Length.ToString(CultureInfo.InvariantCulture) + " >>\nstream\n");
        output.Write(compressedMask, 0, compressedMask.Length);
        WriteAscii("\nendstream\nendobj\ntrailer\n<< /Root 1 0 R >>\n%%EOF\n");
        return output.ToArray();
    }

    private static byte[] BuildSimpleFlateImagePdf(string pageContent) {
        byte[] pixels = CreateSimpleFlateImagePixels();
        byte[] compressed = Compress(pixels);
        int pageStreamLength = Encoding.ASCII.GetByteCount(pageContent.TrimEnd('\n'));

        using var output = new MemoryStream();
        void WriteAscii(string text) {
            byte[] bytes = Encoding.ASCII.GetBytes(text);
            output.Write(bytes, 0, bytes.Length);
        }

        WriteAscii(string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 200 120] /Resources << /XObject << /ImSimple 5 0 R >> >> >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length " + pageStreamLength.ToString(CultureInfo.InvariantCulture) + " >>",
            "stream",
            pageContent.TrimEnd('\n'),
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /XObject /Subtype /Image /Width 4 /Height 2 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /FlateDecode /Length " + compressed.Length.ToString(CultureInfo.InvariantCulture) + " >>",
            "stream"
        }) + "\n");
        output.Write(compressed, 0, compressed.Length);
        WriteAscii("\nendstream\nendobj\ntrailer\n<< /Root 1 0 R >>\n%%EOF\n");
        return output.ToArray();
    }

    private static byte[] BuildJpegImageRedactionSource(string imageFilter = "/DCTDecode") {
        const string pageContent = "q\n20 0 0 20 20 30 cm\n/ImJpeg Do\nQ\nBT /F1 12 Tf 20 90 Td (REMOVE-JPEG-DOC-TEXT) Tj ET\n";
        byte[] jpegBytes = new byte[] { 0xFF, 0xD8, 0xFF, 0xD9 };
        int pageStreamLength = Encoding.ASCII.GetByteCount(pageContent.TrimEnd('\n'));
        using var output = new MemoryStream();
        void WriteAscii(string text) {
            byte[] bytes = Encoding.ASCII.GetBytes(text);
            output.Write(bytes, 0, bytes.Length);
        }

        WriteAscii(string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 200 120] /Resources << /XObject << /ImJpeg 5 0 R >> /Font << /F1 6 0 R >> >> >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length " + pageStreamLength.ToString(CultureInfo.InvariantCulture) + " >>",
            "stream",
            pageContent.TrimEnd('\n'),
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter " + imageFilter + " /Length " + jpegBytes.Length.ToString(CultureInfo.InvariantCulture) + " >>",
            "stream"
        }) + "\n");
        output.Write(jpegBytes, 0, jpegBytes.Length);
        WriteAscii("\nendstream\nendobj\n6 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj\ntrailer\n<< /Root 1 0 R >>\n%%EOF\n");
        return output.ToArray();
    }

    private static byte[] CreateSimpleFlateImagePixels() {
        return new byte[] {
            255, 0, 0, 255, 0, 0, 0, 255, 0, 0, 255, 0,
            0, 0, 255, 0, 0, 255, 255, 255, 255, 255, 255, 255
        };
    }

    private static PdfLogicalImage GetSingleImage(byte[] source) {
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(source, new PdfTextLayoutOptions {
            ForceSingleColumn = true
        });
        PdfLogicalImage image = Assert.Single(logical.Images);
        Assert.NotNull(image.PrimaryPlacement);
        return image;
    }

    private static PdfRedactionArea CreateImageIntersectionArea(PdfLogicalImage image) {
        PdfImagePlacement placement = image.PrimaryPlacement!;
        return new PdfRedactionArea(
            image.PageNumber,
            placement.X + 1D,
            placement.Y + 1D,
            Math.Max(1D, placement.Width / 2D),
            Math.Max(1D, placement.Height / 2D),
            "image-risk");
    }

    private static PdfRedactionArea CreateImageLeftHalfArea(PdfLogicalImage image, PdfImagePlacement placement) {
        return new PdfRedactionArea(
            image.PageNumber,
            placement.X,
            placement.Y,
            placement.Width / 2D,
            placement.Height,
            "image-pixel-redact");
    }

    private static PdfRedactionArea CreateImageCoveringArea(PdfLogicalImage image) {
        return CreateImageCoveringArea(image, image.PrimaryPlacement!);
    }

    private static PdfRedactionArea CreateImageCoveringArea(PdfLogicalImage image, PdfImagePlacement placement) {
        return new PdfRedactionArea(
            image.PageNumber,
            placement.X,
            placement.Y,
            placement.Width,
            placement.Height,
            "image-remove");
    }

    private static byte[] AppendPdfHexStringResidue(byte[] pdf, string marker) {
        byte[] suffix = Encoding.ASCII.GetBytes("\n999 0 obj\n<" + ToHex(Encoding.BigEndianUnicode.GetBytes(marker)) + ">\nendobj\n");
        return pdf.Concat(suffix).ToArray();
    }

    private static byte[] AppendPdfLiteralStringResidue(byte[] pdf, string escapedMarker) {
        byte[] suffix = Encoding.ASCII.GetBytes("\n998 0 obj\n(" + escapedMarker + ")\nendobj\n");
        return pdf.Concat(suffix).ToArray();
    }

    private static byte[] AppendFlateStreamResidue(byte[] pdf, string marker) {
        byte[] compressed = Compress(Encoding.UTF8.GetBytes("compressed residue " + marker));
        using var output = new MemoryStream();
        output.Write(pdf, 0, pdf.Length);
        byte[] header = Encoding.ASCII.GetBytes("\n997 0 obj\n<< /Length " + compressed.Length.ToString(CultureInfo.InvariantCulture) + " /Filter /FlateDecode >>\nstream\n");
        output.Write(header, 0, header.Length);
        output.Write(compressed, 0, compressed.Length);
        byte[] footer = Encoding.ASCII.GetBytes("\nendstream\nendobj\n");
        output.Write(footer, 0, footer.Length);
        return output.ToArray();
    }

    private static byte[] AppendUnsupportedFilteredStream(byte[] pdf) {
        byte[] data = new byte[] { 0xFF, 0xD8, 0xFF, 0xD9 };
        using var output = new MemoryStream();
        output.Write(pdf, 0, pdf.Length);
        byte[] header = Encoding.ASCII.GetBytes("\n996 0 obj\n<< /Length " + data.Length.ToString(CultureInfo.InvariantCulture) + " /Filter /DCTDecode >>\nstream\n");
        output.Write(header, 0, header.Length);
        output.Write(data, 0, data.Length);
        byte[] footer = Encoding.ASCII.GetBytes("\nendstream\nendobj\n");
        output.Write(footer, 0, footer.Length);
        return output.ToArray();
    }

    private static byte[] DecodeSingleImagePixels(byte[] pdf) {
        return Assert.Single(DecodeImagePixelStreams(pdf));
    }

    private static byte[][] DecodeImagePixelStreams(byte[] pdf) {
        IReadOnlyList<PdfExtractedImage> images = PdfImageExtractor.ExtractImages(pdf);
        var (objects, _) = PdfSyntax.ParseObjects(pdf);
        return images
            .Where(image => image.ObjectNumber > 0)
            .Select(image => objects.TryGetValue(image.ObjectNumber, out PdfIndirectObject? indirect) && indirect.Value is PdfStream stream
                ? StreamDecoder.Decode(stream.Dictionary, stream.Data, objects)
                : Array.Empty<byte>())
            .ToArray();
    }

    private static byte[] DecodeSoftMaskPixels(byte[] pdf, int imageObjectNumber) {
        var (objects, _) = PdfSyntax.ParseObjects(pdf);
        Assert.True(objects.TryGetValue(imageObjectNumber, out PdfIndirectObject? imageIndirect));
        PdfStream imageStream = Assert.IsType<PdfStream>(imageIndirect.Value);
        PdfReference softMaskReference = Assert.IsType<PdfReference>(imageStream.Dictionary.Items["SMask"]);
        Assert.True(objects.TryGetValue(softMaskReference.ObjectNumber, out PdfIndirectObject? softMaskIndirect));
        PdfStream softMaskStream = Assert.IsType<PdfStream>(softMaskIndirect.Value);
        return StreamDecoder.Decode(softMaskStream.Dictionary, softMaskStream.Data, objects);
    }

    private static void AssertRedactedLeftHalf(byte[] pixels, int width, int height, int components) {
        Assert.True(LeftHalfIsRedacted(pixels, width, height, components));
    }

    private static void AssertSoftMaskLeftHalfOpaque(byte[] pixels, int width, int height) {
        Assert.Equal(width * height, pixels.Length);
        for (int row = 0; row < height; row++) {
            for (int column = 0; column < width; column++) {
                byte value = pixels[row * width + column];
                if (column < width / 2) {
                    Assert.Equal((byte)255, value);
                } else {
                    Assert.NotEqual((byte)255, value);
                }
            }
        }
    }

    private static void AssertInvertedDecodeRedactedLeftHalf(byte[] pixels, int width, int height, int components) {
        Assert.Equal(width * height * components, pixels.Length);
        for (int row = 0; row < height; row++) {
            for (int column = 0; column < width; column++) {
                int offset = ((row * width) + column) * components;
                if (column < width / 2) {
                    Assert.Equal((byte)255, pixels[offset]);
                    Assert.Equal((byte)255, pixels[offset + 1]);
                    Assert.Equal((byte)255, pixels[offset + 2]);
                }
            }
        }
    }

    private static void AssertInvertedDecodeSoftMaskLeftHalfOpaque(byte[] pixels, int width, int height) {
        Assert.Equal(width * height, pixels.Length);
        for (int row = 0; row < height; row++) {
            for (int column = 0; column < width; column++) {
                byte value = pixels[row * width + column];
                if (column < width / 2) {
                    Assert.Equal((byte)0, value);
                } else {
                    Assert.NotEqual((byte)0, value);
                }
            }
        }
    }

    private static bool LeftHalfIsRedacted(byte[] pixels, int width, int height, int components) {
        for (int row = 0; row < height; row++) {
            for (int column = 0; column < width; column++) {
                int offset = ((row * width) + column) * components;
                if (column < width / 2) {
                    if (pixels[offset] != 0 || pixels[offset + 1] != 0 || pixels[offset + 2] != 0) {
                        return false;
                    }
                } else {
                    if (pixels[offset] == 0 && pixels[offset + 1] == 0 && pixels[offset + 2] == 0) {
                        return false;
                    }
                }
            }
        }

        return true;
    }

    private static bool PixelRowsMatch(byte[] left, byte[] right) {
        if (left.Length != right.Length) {
            return false;
        }

        for (int i = 0; i < left.Length; i++) {
            if (left[i] != right[i]) {
                return false;
            }
        }

        return true;
    }

    private static byte[] Compress(byte[] bytes) {
        using var output = new MemoryStream();
        using (var deflate = new DeflateStream(output, CompressionLevel.Optimal, leaveOpen: true)) {
            deflate.Write(bytes, 0, bytes.Length);
        }

        return output.ToArray();
    }

    private static string ToHex(byte[] bytes) {
        var builder = new StringBuilder(bytes.Length * 2);
        for (int i = 0; i < bytes.Length; i++) {
            builder.Append(bytes[i].ToString("X2", CultureInfo.InvariantCulture));
        }

        return builder.ToString();
    }
}
