using System.Globalization;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfRedactionApplierTests {
    [Fact]
    public void Apply_RemovesAreaIntersectingTextEvenWhenExtractionProducesNoMatch() {
        const string whitespaceTextObject = "BT\n/F1 12 Tf\n50 60 Td\n(   ) Tj\nET";
        var objects = new[] {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 300 300] /Resources << /Font << /F1 4 0 R >> >> >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Contents 5 0 R >>\nendobj",
            "4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>\nendobj",
            BuildStreamObject(5, Encoding.ASCII.GetBytes(whitespaceTextObject))
        };
        byte[] source = BuildPdf(objects, rootObjectNumber: 1);
        var area = new PdfRedactionArea(1, 45, 40, 80, 35, "unmatched-text");
        Assert.Empty(PdfRedactionPlanner.Plan(source, new[] { area }).Matches);
        Assert.Contains("(   ) Tj", PdfEncoding.Latin1GetString(source), StringComparison.Ordinal);

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });

        Assert.DoesNotContain("(   ) Tj", PdfEncoding.Latin1GetString(redacted), StringComparison.Ordinal);
    }

    [Fact]
    public void Apply_AreaRedactionFailsClosedForTextObjectsWithoutKnownGeometry() {
        const string content =
            "BT\n/F1 12 Tf\n50 60 Td\n(   ) Tj\nET\n" +
            "BT\n% unlocatable-outside-target\nET";
        var objects = new[] {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 300 300] /Resources << /Font << /F1 4 0 R >> >> >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Contents 5 0 R >>\nendobj",
            "4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>\nendobj",
            BuildStreamObject(5, Encoding.ASCII.GetBytes(content))
        };
        byte[] source = BuildPdf(objects, rootObjectNumber: 1);
        var area = new PdfRedactionArea(1, 45, 40, 80, 35, "unmatched-text");

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });
        string rewritten = PdfEncoding.Latin1GetString(redacted);

        Assert.DoesNotContain("(   ) Tj", rewritten, StringComparison.Ordinal);
        Assert.DoesNotContain("unlocatable-outside-target", rewritten, StringComparison.Ordinal);
    }

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
        Assert.DoesNotContain("Secret account", PdfEncoding.Latin1GetString(redacted), StringComparison.Ordinal);

        PdfRedactionPlan redactedPlan = PdfRedactionPlanner.Plan(redacted, new[] { area });
        Assert.DoesNotContain(redactedPlan.Matches, match => match.Text != null && match.Text.Contains("Secret account", StringComparison.Ordinal));
    }

    [Theory]
    [InlineData("(Alpha secret Omega) Tj")]
    [InlineData("[(Alpha ) -120 (secret) 40 ( Omega)] TJ")]
    [InlineData("[(Alpha ) <> -120 (secret) 40 ( Omega)] TJ")]
    [InlineData("(Alpha secret Omega) '")]
    [InlineData("0 0 (Alpha secret Omega) \"")]
    public void Apply_RewritesTextShowOperatorsAndPreservesUnmatchedEncodedGlyphs(string showOperation) {
        byte[] source = BuildSingleTextObjectRedactionSource(showOperation);
        PdfTextSpan span = Assert.Single(PdfReadDocument.Open(source).Pages[0].GetTextSpans(), static value => value.Text.Contains("secret", StringComparison.Ordinal));
        PdfRedactionArea area = BuildAreaForSubstring(span, "secret");

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });
        string text = PdfTextExtractor.ExtractAllText(redacted);
        string raw = PdfEncoding.Latin1GetString(redacted);

        Assert.Contains("Alpha", text, StringComparison.Ordinal);
        Assert.Contains("Omega", text, StringComparison.Ordinal);
        Assert.DoesNotContain("secret", text, StringComparison.Ordinal);
        Assert.Contains("416C706861", raw, StringComparison.Ordinal);
        Assert.Contains("4F6D656761", raw, StringComparison.Ordinal);
        Assert.Contains("TJ", raw, StringComparison.Ordinal);
        Assert.Contains("/F1 20 Tf", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void Apply_PreservesEmbeddedFontGlyphBytesOutsidePartialRedaction() {
        const string text = "Alpha secret Omega";
        byte[] source = BuildToUnicodeSingleTextObjectRedactionSource(text);
        PdfTextSpan span = Assert.Single(PdfReadDocument.Open(source).Pages[0].GetTextSpans());
        PdfRedactionArea area = BuildAreaForSubstring(span, "secret");

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });
        string extracted = PdfTextExtractor.ExtractAllText(redacted);
        string raw = PdfEncoding.Latin1GetString(redacted);

        Assert.Contains("Alpha", extracted, StringComparison.Ordinal);
        Assert.Contains("Omega", extracted, StringComparison.Ordinal);
        Assert.DoesNotContain("secret", extracted, StringComparison.Ordinal);
        Assert.Contains("/AAAAAA+Helvetica", raw, StringComparison.Ordinal);
        Assert.DoesNotContain("/BaseFont /Helvetica ", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void Apply_FallsBackToWholeTextObjectWhenActualTextCoversPartialRedaction() {
        byte[] source = BuildSingleTextObjectRedactionSource(
            "/Span << /ActualText (Alpha secret Omega) >> BDC (painted substitute) Tj EMC");
        PdfTextSpan span = Assert.Single(PdfReadDocument.Open(source).Pages[0].GetTextSpans());
        PdfTextSpanBounds bounds = PdfTextSpanGeometry.GetAxisAlignedBounds(span);
        var area = new PdfRedactionArea(
            1,
            bounds.Left + bounds.Width * 0.4D,
            bounds.Bottom + 0.05D,
            bounds.Width * 0.2D,
            bounds.Height - 0.1D,
            "ActualText partial");

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });
        string extracted = PdfTextExtractor.ExtractAllText(redacted);
        string raw = PdfEncoding.Latin1GetString(redacted);

        Assert.DoesNotContain("secret", extracted, StringComparison.Ordinal);
        Assert.DoesNotContain("painted substitute", raw, StringComparison.Ordinal);
        Assert.DoesNotContain("Tj", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void Apply_PreservesTwoByteEmbeddedType0GlyphCodesOutsidePartialRedaction() {
        const string text = "Alpha secret Omega";
        byte[] source = BuildType0ToUnicodeSingleTextObjectRedactionSource(text);
        PdfTextSpan span = Assert.Single(PdfReadDocument.Open(source).Pages[0].GetTextSpans());
        PdfRedactionArea area = BuildAreaForSubstring(span, "secret");

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });
        string extracted = PdfTextExtractor.ExtractAllText(redacted);
        string raw = PdfEncoding.Latin1GetString(redacted);

        Assert.Contains("Alpha", extracted, StringComparison.Ordinal);
        Assert.Contains("Omega", extracted, StringComparison.Ordinal);
        Assert.DoesNotContain("secret", extracted, StringComparison.Ordinal);
        Assert.Contains("00010002000300040005", raw, StringComparison.Ordinal);
        Assert.Contains("000E000F001000110012", raw, StringComparison.Ordinal);
        Assert.Contains("/Subtype /Type0", raw, StringComparison.Ordinal);
        Assert.Contains("/FontFile2 9 0 R", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void Apply_FailsClosedForVerticalWritingTextObjects() {
        byte[] source = PdfEncoding.Latin1GetBytes(
            PdfEncoding.Latin1GetString(BuildType0ToUnicodeSingleTextObjectRedactionSource("Alpha secret Omega"))
                .Replace("/Identity-H", "/Identity-V"));
        PdfTextSpan span = Assert.Single(PdfReadDocument.Open(source).Pages[0].GetTextSpans());
        PdfRedactionArea area = BuildAreaForSubstring(span, "secret");

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });

        Assert.DoesNotContain("secret", PdfTextExtractor.ExtractAllText(redacted), StringComparison.Ordinal);
        Assert.DoesNotContain("Alpha", PdfTextExtractor.ExtractAllText(redacted), StringComparison.Ordinal);
        Assert.DoesNotContain(" Tj", PdfEncoding.Latin1GetString(redacted), StringComparison.Ordinal);
    }

    [Fact]
    public void AppliedPlanVerificationTreatsMultiCharacterToUnicodeMappingAsOneGlyph() {
        byte[] source = BuildType0ToUnicodeLigatureRedactionSource();
        PdfTextSpan span = Assert.Single(PdfReadDocument.Open(source).Pages[0].GetTextSpans());
        Assert.Equal("AfiZ", span.Text);
        PdfRedactionArea area = BuildAreaForSubstring(span, "f");
        PdfRedactionPlan plan = PdfRedactionPlanner.Plan(source, new[] { area });

        byte[] redacted = PdfRedactionApplier.Apply(source, plan);
        PdfRedactionVerificationReport report = PdfRedactionVerification.VerifyAppliedPlan(
            redacted,
            plan,
            new PdfRedactionVerificationOptions { RequireCompleteStreamInspection = true });

        Assert.True(report.IsVerified, string.Join("; ", report.Issues.Select(static issue => issue.Message)));
        Assert.Equal("AZ", PdfTextExtractor.ExtractAllText(redacted).Trim());
    }

    [Fact]
    public void Apply_PreservesGeneratedEmbeddedType0FontOutsidePartialRedaction() {
        const string text = "Alpha secret Omega";
        string fontPath = Assert.IsType<string>(PdfComplianceTestFonts.FindBundledOpenTypeCffFont());
        byte[] source = PdfDocument.Create(new PdfOptions { CompressContentStreams = false })
            .UseFontFamily("Redaction Embedded", fontPath)
            .Paragraph(paragraph => paragraph.Text(text))
            .ToBytes();
        PdfTextSpan span = Assert.Single(PdfReadDocument.Open(source).Pages[0].GetTextSpans(), static value => value.Text == "secret");
        PdfRedactionArea area = BuildAreaForSubstring(span, span.Text);

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });
        string extracted = PdfTextExtractor.ExtractAllText(redacted);
        string raw = PdfEncoding.Latin1GetString(redacted);

        Assert.Contains("Alpha", extracted, StringComparison.Ordinal);
        Assert.Contains("Omega", extracted, StringComparison.Ordinal);
        Assert.DoesNotContain("secret", extracted, StringComparison.Ordinal);
        Assert.Contains("/Subtype /Type0", raw, StringComparison.Ordinal);
        Assert.Contains("/FontFile3", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void Apply_PreservesTextObjectWhenAreaFallsOnlyBetweenItsGlyphRuns() {
        byte[] source = BuildSingleTextObjectRedactionSource("[(Alpha) -1000 (Omega)] TJ");
        PdfTextSpan[] spans = PdfReadDocument.Open(source).Pages[0].GetTextSpans().ToArray();
        Assert.Equal(2, spans.Length);
        double gapLeft = spans[0].X + spans[0].Advance;
        double gapWidth = spans[1].X - gapLeft;
        Assert.True(gapWidth > 2D);
        var area = new PdfRedactionArea(1, gapLeft + 0.5D, spans[0].Y - spans[0].FontSize + 0.05D, gapWidth - 1D, spans[0].FontSize * 1.5D - 0.1D, "gap");

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });
        string text = PdfTextExtractor.ExtractAllText(redacted);

        Assert.Contains("Alpha", text, StringComparison.Ordinal);
        Assert.Contains("Omega", text, StringComparison.Ordinal);
    }

    [Fact]
    public void Apply_DoesNotTreatCharacterSpacingAsPaintedGlyphArea() {
        byte[] source = BuildTextContentRedactionSource("BT /F1 20 Tf 100 Tc 72 700 Td (AB) Tj ET");
        var area = new PdfRedactionArea(1, 90D, 680D, 5D, 30D, "character-spacing gap");
        PdfRedactionPlan plan = PdfRedactionPlanner.Plan(source, new[] { area });

        byte[] redacted = PdfRedactionApplier.Apply(source, plan);
        PdfRedactionVerificationReport verification = PdfRedactionVerification.VerifyAppliedPlan(
            redacted,
            plan,
            new PdfRedactionVerificationOptions());

        Assert.DoesNotContain(plan.Matches, match => match.Kind == PdfRedactionMatchKind.TextBlock);
        Assert.Contains("AB", PdfTextExtractor.ExtractAllText(redacted), StringComparison.Ordinal);
        Assert.True(verification.IsVerified, string.Join(Environment.NewLine, verification.Issues.Select(issue => issue.Message)));
    }

    [Fact]
    public void Apply_DoesNotTreatWordSpacingAsPaintedGlyphArea() {
        byte[] source = BuildTextContentRedactionSource("BT /F1 20 Tf 100 Tw 72 700 Td (A B) Tj ET");
        var area = new PdfRedactionArea(1, 100D, 680D, 5D, 30D, "word-spacing gap");
        PdfRedactionPlan plan = PdfRedactionPlanner.Plan(source, new[] { area });

        byte[] redacted = PdfRedactionApplier.Apply(source, plan);
        PdfRedactionVerificationReport verification = PdfRedactionVerification.VerifyAppliedPlan(
            redacted,
            plan,
            new PdfRedactionVerificationOptions());

        Assert.DoesNotContain(plan.Matches, match => match.Kind == PdfRedactionMatchKind.TextBlock);
        Assert.Contains("A B", PdfTextExtractor.ExtractAllText(redacted), StringComparison.Ordinal);
        Assert.True(verification.IsVerified, string.Join(Environment.NewLine, verification.Issues.Select(issue => issue.Message)));
    }

    [Fact]
    public void Apply_UsesPaintedTextDirectionWhenNegativeSpacingReversesTheRunEndpoint() {
        byte[] source = BuildTextContentRedactionSource("BT /F1 20 Tf -30 Tc 100 700 Td (AB) Tj ET");
        var area = new PdfRedactionArea(1, 85D, 680D, 5D, 30D, "painted B");
        PdfRedactionPlan plan = PdfRedactionPlanner.Plan(source, new[] { area });

        byte[] redacted = PdfRedactionApplier.Apply(source, plan);
        PdfRedactionVerificationReport verification = PdfRedactionVerification.VerifyAppliedPlan(
            redacted,
            plan,
            new PdfRedactionVerificationOptions());

        Assert.Contains(plan.Matches, match => match.Kind == PdfRedactionMatchKind.TextBlock);
        Assert.Equal("A", PdfTextExtractor.ExtractAllText(redacted).Trim());
        Assert.True(verification.IsVerified, string.Join(Environment.NewLine, verification.Issues.Select(issue => issue.Message)));
    }

    [Fact]
    public void ApplyWithEvidenceIgnoresReviewedPathRemovedBeforePartialTextSurvivor() {
        byte[] source = BuildTextContentRedactionSource(string.Join("\n", new[] {
            "0 0 0 rg 72 695 13 25 re f",
            "BT /F1 20 Tf 72 700 Td (AB) Tj ET"
        }));
        var area = new PdfRedactionArea(1, 74D, 680D, 5D, 30D, "painted path and A");
        PdfDocument document = PdfDocument.Load(source);
        PdfRedactionPlan plan = document.Redactions.Plan(new[] { area });

        PdfRedactionApplyResult result = document.Redactions.ApplyWithEvidence(
            plan,
            verificationOptions: new PdfRedactionVerificationOptions {
                RequireCompleteStreamInspection = true
            });

        Assert.Equal("B", PdfTextExtractor.ExtractAllText(result.Pdf).Trim());
        Assert.True(result.IsVerified, result.Evidence.Summary);
    }

    [Fact]
    public void Apply_IgnoresInlineImagePayloadWhileDiscoveringFormText() {
        byte[] source = BuildInlineImagePayloadBeforeFormRedactionSource();
        PdfRedactionArea area = FindAreaForText(source, "Visible secret");

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });
        string raw = PdfEncoding.Latin1GetString(redacted);

        Assert.DoesNotContain("Visible secret", PdfTextExtractor.ExtractAllText(redacted), StringComparison.Ordinal);
        Assert.Contains("Dormant secret", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void TextExtractionUsesZeroDefaultLeading() {
        byte[] source = BuildTextContentRedactionSource("BT /F1 12 Tf 50 100 Td (A) Tj T* (B) Tj ET");

        PdfTextSpan[] spans = PdfReadDocument.Open(source).Pages[0].GetTextSpans().ToArray();

        Assert.Equal(2, spans.Length);
        Assert.Equal(spans[0].Y, spans[1].Y, precision: 6);
    }

    [Fact]
    public void Type0UnicodeSpaceDoesNotReceiveWordSpacing() {
        byte[] source = BuildType0UnicodeSpaceWordSpacingPdf();

        PdfTextSpan span = Assert.Single(PdfReadDocument.Open(source).Pages[0].GetTextSpans());

        Assert.Equal(" ", span.Text);
        Assert.Equal(12D, span.Advance, precision: 6);
    }

    [Fact]
    public void Apply_UsesTextStateInheritedAcrossTextObjects() {
        byte[] source = BuildTextContentRedactionSource(string.Join("\n", new[] {
            "/F1 20 Tf",
            "10 Tc",
            "BT 72 750 Td (Prelude) Tj ET",
            "BT 72 700 Td (Alpha secret Omega) Tj ET"
        }));
        PdfTextSpan span = Assert.Single(
            PdfReadDocument.Open(source).Pages[0].GetTextSpans(),
            static value => value.Text.Contains("secret", StringComparison.Ordinal));
        PdfRedactionArea area = BuildAreaForSubstring(span, "secret");

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });
        string text = PdfTextExtractor.ExtractAllText(redacted);

        Assert.Contains("Alpha", text, StringComparison.Ordinal);
        Assert.Contains("Omega", text, StringComparison.Ordinal);
        Assert.DoesNotContain("secret", text, StringComparison.Ordinal);
    }

    [Fact]
    public void Apply_RestoresTextStateAcrossGraphicsStateScopes() {
        byte[] source = BuildTextContentRedactionSource(string.Join("\n", new[] {
            "/F1 20 Tf",
            "10 Tc",
            "q /F1 8 Tf 0 Tc Q",
            "BT 72 700 Td (Alpha secret Omega) Tj ET"
        }));
        PdfTextSpan span = Assert.Single(PdfReadDocument.Open(source).Pages[0].GetTextSpans());
        PdfRedactionArea area = BuildAreaForSubstring(span, "secret");

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });
        string text = PdfTextExtractor.ExtractAllText(redacted);

        Assert.Contains("Alpha", text, StringComparison.Ordinal);
        Assert.Contains("Omega", text, StringComparison.Ordinal);
        Assert.DoesNotContain("secret", text, StringComparison.Ordinal);
    }

    [Fact]
    public void Apply_NormalizesMirroredCharacterAdvancesBeforeHitTesting() {
        byte[] source = BuildTextContentRedactionSource("/F1 20 Tf -100 Tz BT 300 700 Td (Alpha secret Omega) Tj ET");
        PdfTextSpan span = Assert.Single(PdfReadDocument.Open(source).Pages[0].GetTextSpans());
        PdfRedactionArea area = BuildAreaForSubstring(span, "secret");

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });
        string text = PdfTextExtractor.ExtractAllText(redacted);

        Assert.Contains("Alpha", text, StringComparison.Ordinal);
        Assert.Contains("Omega", text, StringComparison.Ordinal);
        Assert.DoesNotContain("secret", text, StringComparison.Ordinal);
    }

    [Fact]
    public void AppliedPlanVerificationRejectsChangedStateOnPreservedGlyphs() {
        byte[] source = BuildSingleTextObjectRedactionSource("(Alpha secret Omega) Tj");
        PdfTextSpan span = Assert.Single(PdfReadDocument.Open(source).Pages[0].GetTextSpans());
        PdfRedactionArea area = BuildAreaForSubstring(span, "secret");
        PdfRedactionPlan plan = PdfRedactionPlanner.Plan(source, new[] { area });
        byte[] redacted = PdfRedactionApplier.Apply(source, plan);

        PdfRedactionVerificationReport accepted = PdfRedactionVerification.VerifyAppliedPlan(
            redacted,
            plan,
            new PdfRedactionVerificationOptions { RequireCompleteStreamInspection = true });
        string mutatedRaw = PdfEncoding.Latin1GetString(redacted)
            .Replace("/F1 20 Tf", "/F1 21 Tf");
        PdfRedactionVerificationReport rejected = PdfRedactionVerification.VerifyAppliedPlan(
            PdfEncoding.Latin1GetBytes(mutatedRaw),
            plan,
            new PdfRedactionVerificationOptions { RequireCompleteStreamInspection = true });

        Assert.True(accepted.IsVerified, string.Join("; ", accepted.Issues.Select(static issue => issue.Message)));
        Assert.False(rejected.IsVerified);
        Assert.Contains(rejected.Issues, static issue => issue.Feature == "RedactionPlanPageIdentityChanged");
    }

    [Theory]
    [InlineData("1 w", "2 w")]
    [InlineData("0 J", "1 J")]
    [InlineData("0 j", "1 j")]
    [InlineData("9 M", "8 M")]
    [InlineData("[1] 0 d", "[2] 0 d")]
    public void AppliedPlanVerificationRejectsChangedStrokeParametersOnPreservedGlyphs(string originalState, string changedState) {
        byte[] source = BuildTextContentRedactionSource("2 Tr\n1 w\n0 J\n0 j\n9 M\n[1] 0 d\nBT\n/F1 20 Tf\n72 720 Td\n(Alpha secret Omega) Tj\nET");
        PdfTextSpan span = Assert.Single(PdfReadDocument.Open(source).Pages[0].GetTextSpans());
        PdfRedactionArea area = BuildAreaForSubstring(span, "secret");
        PdfRedactionPlan plan = PdfRedactionPlanner.Plan(source, new[] { area });
        byte[] redacted = PdfRedactionApplier.Apply(source, plan);

        PdfRedactionVerificationReport accepted = PdfRedactionVerification.VerifyAppliedPlan(
            redacted,
            plan,
            new PdfRedactionVerificationOptions { RequireCompleteStreamInspection = true });
        byte[] mutated = PdfEncoding.Latin1GetBytes(
            PdfEncoding.Latin1GetString(redacted).Replace(originalState, changedState));
        PdfRedactionVerificationReport rejected = PdfRedactionVerification.VerifyAppliedPlan(
            mutated,
            plan,
            new PdfRedactionVerificationOptions { RequireCompleteStreamInspection = true });

        Assert.True(accepted.IsVerified, string.Join("; ", accepted.Issues.Select(static issue => issue.Message)));
        Assert.False(rejected.IsVerified);
        Assert.Contains(rejected.Issues, static issue => issue.Feature == "RedactionPlanPageIdentityChanged");
    }

    [Fact]
    public void AppliedPlanVerificationRejectsUnreviewedTextMovedAcrossPartialSurvivor() {
        const string backdrop = "BT /F1 10 Tf 72 700 Td (Backdrop) Tj ET";
        const string reviewed = "BT /F1 20 Tf 72 700 Td (Alpha secret Omega) Tj ET";
        byte[] source = BuildTextContentRedactionSource(backdrop + "\n" + reviewed);
        PdfTextSpan span = Assert.Single(
            PdfReadDocument.Open(source).Pages[0].GetTextSpans(),
            static value => value.Text.Contains("secret", StringComparison.Ordinal));
        PdfRedactionArea area = BuildAreaForSubstring(span, "secret");
        PdfRedactionPlan plan = PdfRedactionPlanner.Plan(source, new[] { area });
        byte[] redacted = PdfRedactionApplier.Apply(source, plan);
        string raw = PdfEncoding.Latin1GetString(redacted);
        int backdropStart = raw.IndexOf(backdrop, StringComparison.Ordinal);
        int reviewedStart = raw.IndexOf("BT /F1 20 Tf", StringComparison.Ordinal);
        int reviewedEnd = raw.IndexOf("ET", reviewedStart, StringComparison.Ordinal) + 2;
        Assert.True(backdropStart >= 0 && reviewedStart > backdropStart && reviewedEnd > reviewedStart);
        string reordered = raw.Substring(0, backdropStart) +
            raw.Substring(reviewedStart, reviewedEnd - reviewedStart) + "\n" +
            backdrop +
            raw.Substring(reviewedEnd);

        PdfRedactionVerificationReport report = PdfRedactionVerification.VerifyAppliedPlan(
            PdfEncoding.Latin1GetBytes(reordered),
            plan,
            new PdfRedactionVerificationOptions { RequireCompleteStreamInspection = true });

        Assert.False(report.IsVerified);
        Assert.Contains(report.Issues, static issue => issue.Feature == "RedactionPlanPageIdentityChanged");
    }

    [Fact]
    public void AppliedPlanVerificationRejectsMissingTextObjectWithExpectedSurvivors() {
        const string textObject = "BT\n/F1 20 Tf\n72 720 Td\n(Alpha secret Omega) Tj\nET";
        byte[] source = BuildTextContentRedactionSource(textObject);
        PdfTextSpan span = Assert.Single(PdfReadDocument.Open(source).Pages[0].GetTextSpans());
        PdfRedactionArea area = BuildAreaForSubstring(span, "secret");
        PdfRedactionPlan plan = PdfRedactionPlanner.Plan(source, new[] { area });
        string removedRaw = PdfEncoding.Latin1GetString(source)
            .Replace(textObject, new string(' ', textObject.Length));

        PdfRedactionVerificationReport report = PdfRedactionVerification.VerifyAppliedPlan(
            PdfEncoding.Latin1GetBytes(removedRaw),
            plan,
            new PdfRedactionVerificationOptions { RequireCompleteStreamInspection = true });

        Assert.False(report.IsVerified);
        Assert.Contains(report.Issues, static issue => issue.Feature == "RedactionPlanPageIdentityChanged");
    }

    [Fact]
    public void ApplyRedactions_FacadeReturnsRedactedDocumentAndTryResult() {
        byte[] source = BuildRedactionSource();
        PdfRedactionArea area = FindAreaForText(source, "Secret account 123-45");
        PdfDocument document = PdfDocument.Load(source);

        PdfDocument redacted = document.ApplyRedactions(new[] { area });
        PdfOperationResult<PdfDocument> result = document.TryApplyRedactions(new[] { area });

        Assert.DoesNotContain("Secret account", redacted.Reader.Text(), StringComparison.Ordinal);
        Assert.True(result.Succeeded);
        Assert.DoesNotContain("Secret account", result.RequireValue().Reader.Text(), StringComparison.Ordinal);
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

    [Fact]
    public void Apply_DecodesNestedLiteralStringsBeforeMatchingTextObjects() {
        byte[] source = BuildNestedLiteralRedactionSource();
        PdfRedactionArea area = FindAreaForText(source, "Account (secret)");
        Assert.Contains("Account (secret)", PdfTextExtractor.ExtractAllText(source), StringComparison.Ordinal);

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });
        string text = PdfTextExtractor.ExtractAllText(redacted);

        Assert.DoesNotContain("Account (secret)", text, StringComparison.Ordinal);
        Assert.DoesNotContain("secret", text, StringComparison.Ordinal);
        Assert.Contains("Visible before", text, StringComparison.Ordinal);
        Assert.Contains("Visible after", text, StringComparison.Ordinal);
    }

    [Fact]
    public void Apply_ScrubsTextPositionedByOuterGraphicsTransform() {
        byte[] source = BuildTransformedTextRedactionSource();
        PdfRedactionArea area = FindAreaForText(source, "Transformed secret");

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });

        Assert.DoesNotContain("Transformed secret", PdfTextExtractor.ExtractAllText(redacted), StringComparison.Ordinal);
        Assert.DoesNotContain("Transformed secret", PdfEncoding.Latin1GetString(redacted), StringComparison.Ordinal);
    }

    [Fact]
    public void Apply_ScrubsTextPositionedByTransformFromPriorContentStream() {
        byte[] source = BuildSplitTransformedTextRedactionSource();
        PdfRedactionArea area = FindAreaForText(source, "Split transformed secret");

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });

        Assert.DoesNotContain("Split transformed secret", PdfTextExtractor.ExtractAllText(redacted), StringComparison.Ordinal);
        Assert.DoesNotContain("Split transformed secret", PdfEncoding.Latin1GetString(redacted), StringComparison.Ordinal);
    }

    [Fact]
    public void Apply_ScrubsTextPositionedByTransformOperandsSplitAcrossContentStreams() {
        byte[] source = BuildSplitTransformOperandRedactionSource();
        PdfRedactionArea area = FindAreaForText(source, "Split operand secret");

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });

        Assert.DoesNotContain("Split operand secret", PdfTextExtractor.ExtractAllText(redacted), StringComparison.Ordinal);
        Assert.DoesNotContain("Split operand secret", PdfEncoding.Latin1GetString(redacted), StringComparison.Ordinal);
    }

    [Fact]
    public void Apply_IgnoresEndTextOperatorInsideLiteralStringsWhenScrubbingTextObjects() {
        byte[] source = BuildLiteralEndTextOperatorRedactionSource();
        PdfRedactionArea area = FindAreaForText(source, "SSN ET 123");
        Assert.Contains("SSN ET 123", PdfTextExtractor.ExtractAllText(source), StringComparison.Ordinal);

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });
        string text = PdfTextExtractor.ExtractAllText(redacted);
        string raw = PdfEncoding.Latin1GetString(redacted);

        Assert.DoesNotContain("SSN ET 123", text, StringComparison.Ordinal);
        Assert.DoesNotContain("SSN ET 123", raw, StringComparison.Ordinal);
        Assert.Contains("Visible before", text, StringComparison.Ordinal);
        Assert.Contains("Visible after", text, StringComparison.Ordinal);
    }

    [Fact]
    public void Apply_UsesFontDecoderForLiteralRedactionText() {
        byte[] source = BuildToUnicodeLiteralRedactionSource();
        PdfRedactionArea area = FindAreaForText(source, "Secret account 123-45");
        Assert.Contains("Secret account 123-45", ExtractLogicalText(source), StringComparison.Ordinal);

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });
        string text = ExtractLogicalText(redacted);

        Assert.DoesNotContain("Secret account", text, StringComparison.Ordinal);
        Assert.DoesNotContain("123-45", text, StringComparison.Ordinal);
    }

    [Fact]
    public void Apply_ScrubsMatchedTextInsideFormXObjects() {
        byte[] source = BuildFormXObjectRedactionSource();
        PdfRedactionArea area = FindAreaForText(source, "Secret account 123-45");
        Assert.Contains("Secret account 123-45", PdfTextExtractor.ExtractAllText(source), StringComparison.Ordinal);

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });
        string text = PdfTextExtractor.ExtractAllText(redacted);

        Assert.DoesNotContain("Secret account", text, StringComparison.Ordinal);
        Assert.DoesNotContain("123-45", text, StringComparison.Ordinal);
    }

    [Fact]
    public void Apply_UsesTextStateInheritedByFormInvocation() {
        byte[] source = BuildInheritedTextStateFormRedactionSource();
        PdfTextSpan span = Assert.Single(
            PdfReadDocument.Open(source).Pages[0].GetTextSpans(),
            static value => value.Text.Contains("secret", StringComparison.Ordinal));
        PdfRedactionArea area = BuildAreaForSubstring(span, "secret");

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });
        string text = PdfTextExtractor.ExtractAllText(redacted);

        Assert.Contains("Alpha", text, StringComparison.Ordinal);
        Assert.Contains("Omega", text, StringComparison.Ordinal);
        Assert.DoesNotContain("secret", text, StringComparison.Ordinal);
    }

    [Fact]
    public void Apply_UsesFontSelectedByExtendedGraphicsStateBeforeFormInvocation() {
        byte[] source = BuildExtGStateFontFormRedactionSource();
        var area = new PdfRedactionArea(1, 182D, 680D, 100D, 35D, "ExtGState-selected secret");

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });
        string text = PdfTextExtractor.ExtractAllText(redacted);

        Assert.Contains("Alpha", text, StringComparison.Ordinal);
        Assert.Contains("Omega", text, StringComparison.Ordinal);
        Assert.DoesNotContain("secret", text, StringComparison.Ordinal);
    }

    [Fact]
    public void Apply_UsesSpacingSetByDoubleQuoteBeforeFormInvocation() {
        byte[] source = BuildQuotedTextStateFormRedactionSource();
        PdfTextSpan span = Assert.Single(
            PdfReadDocument.Open(source).Pages[0].GetTextSpans(),
            static value => value.Text.Contains("secret", StringComparison.Ordinal));
        PdfRedactionArea area = BuildAreaForSubstring(span, "secret");

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });
        string text = PdfTextExtractor.ExtractAllText(redacted);

        Assert.Contains("Alpha", text, StringComparison.Ordinal);
        Assert.Contains("Omega", text, StringComparison.Ordinal);
        Assert.DoesNotContain("secret", text, StringComparison.Ordinal);
    }

    [Fact]
    public void Apply_PreservesInheritedFontWhenFormShadowsItsResourceName() {
        byte[] source = BuildCollidingInheritedFontFormRedactionSource();
        PdfTextSpan span = Assert.Single(
            PdfReadDocument.Open(source).Pages[0].GetTextSpans(),
            static value => value.Text.Contains("secret", StringComparison.Ordinal));
        Assert.Equal("Helvetica", span.BaseFont);
        PdfRedactionArea area = BuildAreaForSubstring(span, "secret");

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });
        string text = PdfTextExtractor.ExtractAllText(redacted);

        Assert.Contains("Alpha", text, StringComparison.Ordinal);
        Assert.Contains("Omega", text, StringComparison.Ordinal);
        Assert.DoesNotContain("secret", text, StringComparison.Ordinal);
    }

    [Fact]
    public void Apply_UsesTextStateInheritedAcrossNestedFormInvocations() {
        byte[] source = BuildNestedInheritedTextStateFormRedactionSource();
        PdfTextSpan span = Assert.Single(
            PdfReadDocument.Open(source).Pages[0].GetTextSpans(),
            static value => value.Text.Contains("secret", StringComparison.Ordinal));
        PdfRedactionArea area = BuildAreaForSubstring(span, "secret");

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });
        string text = PdfTextExtractor.ExtractAllText(redacted);

        Assert.Contains("Alpha", text, StringComparison.Ordinal);
        Assert.Contains("Omega", text, StringComparison.Ordinal);
        Assert.DoesNotContain("secret", text, StringComparison.Ordinal);
    }

    [Fact]
    public void Apply_PreservesTokensSplitAcrossPageContentStreamsWhenLocatingForms() {
        byte[] source = BuildSplitFormTransformOperandRedactionSource();
        PdfRedactionArea area = FindAreaForText(source, "Split form secret");

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });

        Assert.DoesNotContain("Split form secret", PdfTextExtractor.ExtractAllText(redacted), StringComparison.Ordinal);
        Assert.DoesNotContain("Split form secret", PdfEncoding.Latin1GetString(redacted), StringComparison.Ordinal);
    }

    [Fact]
    public void Apply_ScrubsMatchedTextInsideNestedFormXObjects() {
        byte[] source = BuildNestedFormXObjectRedactionSource();
        PdfRedactionArea area = FindAreaForText(source, "Nested secret account 123-45");
        Assert.Contains("Nested secret account 123-45", PdfTextExtractor.ExtractAllText(source), StringComparison.Ordinal);

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });
        string text = PdfTextExtractor.ExtractAllText(redacted);
        string raw = PdfEncoding.Latin1GetString(redacted);

        Assert.DoesNotContain("Nested secret", text, StringComparison.Ordinal);
        Assert.DoesNotContain("123-45", text, StringComparison.Ordinal);
        Assert.DoesNotContain("Nested secret account 123-45", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void Apply_ClonesSharedPageContentBeforeScrubbingMatchedText() {
        byte[] source = BuildSharedPageContentPdf();
        PdfRedactionArea area = FindAreasForText(source, "Shared page secret").Single(redaction => redaction.PageNumber == 1);
        Assert.Equal(2, CountOccurrences(PdfTextExtractor.ExtractAllText(source), "Shared page secret"));

        PdfRedactionPlan plan = PdfRedactionPlanner.Plan(source, new[] { area });
        byte[] redacted = PdfRedactionApplier.Apply(source, plan);
        string text = PdfTextExtractor.ExtractAllText(redacted);
        PdfRedactionVerificationReport report = PdfRedactionVerification.VerifyAppliedPlan(
            redacted,
            plan,
            new PdfRedactionVerificationOptions { RequireCompleteStreamInspection = true });

        Assert.Equal(1, CountOccurrences(text, "Shared page secret"));
        Assert.True(report.IsVerified, string.Join("; ", report.Issues.Select(static issue => issue.Message)));
    }

    [Fact]
    public void ReplacePageContentReferenceAtIndex_ReplacesOnlyTheSelectedRepeatedOccurrence() {
        var first = new PdfReference(5, 0);
        var middle = new PdfReference(6, 0);
        var repeated = new PdfReference(5, 0);
        var replacement = new PdfReference(7, 0);
        var contents = new PdfArray();
        contents.Items.Add(first);
        contents.Items.Add(middle);
        contents.Items.Add(repeated);
        var page = new PdfDictionary();
        page.Items["Contents"] = contents;

        PdfRedactionApplier.ReplacePageContentReferenceAtIndex(
            new Dictionary<int, PdfIndirectObject>(),
            page,
            contents,
            contentIndex: 2,
            replacement);

        Assert.Same(first, contents.Items[0]);
        Assert.Same(middle, contents.Items[1]);
        Assert.Same(replacement, contents.Items[2]);
    }

    [Fact]
    public void Apply_ClonesSharedFormXObjectBeforeScrubbingMatchedText() {
        byte[] source = BuildSharedFormXObjectTextPdf();
        PdfRedactionArea area = FindAreasForText(source, "Shared form secret").Single(redaction => redaction.PageNumber == 1);
        Assert.Equal(2, CountOccurrences(PdfTextExtractor.ExtractAllText(source), "Shared form secret"));

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });
        string text = PdfTextExtractor.ExtractAllText(redacted);

        Assert.Equal(1, CountOccurrences(text, "Shared form secret"));
    }

    [Fact]
    public void Apply_ClonesRepeatedFormInvocationBeforeScrubbingIntersectingInstance() {
        byte[] source = BuildRepeatedFormXObjectTextPdf();
        PdfRedactionArea area = FindAreaForTextOccurrence(source, "Repeated form secret", occurrenceFromTop: 1);
        Assert.Equal(2, CountOccurrences(PdfTextExtractor.ExtractAllText(source), "Repeated form secret"));

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });
        string text = PdfTextExtractor.ExtractAllText(redacted);

        Assert.Equal(1, CountOccurrences(text, "Repeated form secret"));
    }

    [Fact]
    public void Apply_ClonesIndirectNestedFormResourcesBeforeScrubbingIntersectingInstance() {
        byte[] source = BuildRepeatedNestedFormWithIndirectResourcesPdf();
        PdfRedactionArea area = FindAreaForTextOccurrence(source, "Indirect nested secret", occurrenceFromTop: 1);
        Assert.Equal(2, CountOccurrences(PdfTextExtractor.ExtractAllText(source), "Indirect nested secret"));

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });
        string text = PdfTextExtractor.ExtractAllText(redacted);

        Assert.Equal(1, CountOccurrences(text, "Indirect nested secret"));
    }

    [Fact]
    public void Apply_IsolatesExistingContentBeforePaintingRedactionOverlay() {
        byte[] source = BuildLeakingGraphicsStateRedactionSource();
        var area = new PdfRedactionArea(1, 40, 40, 80, 24, "manual");

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });
        string raw = PdfEncoding.Latin1GetString(redacted);

        Assert.Contains("\nq\n", raw, StringComparison.Ordinal);
        Assert.Contains("0 0 1 1 re W n", raw, StringComparison.Ordinal);
        Assert.Contains("\nQ\n", raw, StringComparison.Ordinal);
        Assert.Contains("40 40 80 24 re", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void Apply_IsolatesContentArrayAsSingleSequenceBeforeOverlay() {
        byte[] source = BuildSplitContentStateRedactionSource();
        var area = new PdfRedactionArea(1, 40, 40, 80, 24, "manual");

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });
        string raw = PdfEncoding.Latin1GetString(redacted);

        Assert.Contains("\nq\n", raw, StringComparison.Ordinal);
        Assert.Contains("\nQ\n", raw, StringComparison.Ordinal);
        Assert.Contains("/F1 12 Tf", raw, StringComparison.Ordinal);
        Assert.Contains("(Visible split text) Tj", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void Apply_ScopesDuplicateTextRedactionToIntersectingInstance() {
        byte[] source = BuildDuplicateRedactionSource();
        PdfRedactionArea area = FindAreaForTextOccurrence(source, "Repeat secret", occurrenceFromTop: 1);
        string originalText = PdfTextExtractor.ExtractAllText(source);
        Assert.Equal(2, CountOccurrences(originalText, "Repeat secret"));

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });
        string text = PdfTextExtractor.ExtractAllText(redacted);

        Assert.Equal(1, CountOccurrences(text, "Repeat secret"));
        Assert.Contains("Visible before", text, StringComparison.Ordinal);
        Assert.Contains("Visible after", text, StringComparison.Ordinal);
    }

    [Fact]
    public void Apply_PrunesRemovedAnnotationAppearanceStreams() {
        byte[] source = BuildAnnotationAppearanceRedactionSource();
        var area = new PdfRedactionArea(1, 20, 20, 40, 40, "annotation");

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });
        string raw = PdfEncoding.Latin1GetString(redacted);

        Assert.Contains("Sensitive annotation", PdfEncoding.Latin1GetString(source), StringComparison.Ordinal);
        Assert.DoesNotContain("Sensitive annotation", raw, StringComparison.Ordinal);
        Assert.DoesNotContain("Old sensitive appearance", raw, StringComparison.Ordinal);
        Assert.Empty(PdfInspector.Inspect(redacted).GetAnnotationsBySubtype("FreeText"));
    }

    [Fact]
    public void Apply_ClearsParentPopupReferenceWhenRedactingPopupAnnotation() {
        byte[] source = BuildIndirectAnnotationWithPopupPdf();
        var area = new PdfRedactionArea(1, 100, 100, 60, 60, "popup");

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });
        string raw = PdfEncoding.Latin1GetString(redacted);

        Assert.DoesNotContain("/Popup", raw, StringComparison.Ordinal);
        Assert.Single(PdfInspector.Inspect(redacted).GetAnnotationsBySubtype("Text"));
        Assert.Empty(PdfInspector.Inspect(redacted).GetAnnotationsBySubtype("Popup"));
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

    private static byte[] BuildNestedLiteralRedactionSource() {
        string streamContent = string.Join("\n", new[] {
            "BT",
            "/F1 12 Tf",
            "72 720 Td",
            "(Visible before) Tj",
            "ET",
            "BT",
            "/F1 12 Tf",
            "72 702 Td",
            "(Account (secret)) Tj",
            "ET",
            "BT",
            "/F1 12 Tf",
            "72 650 Td",
            "(Visible after) Tj",
            "ET"
        });
        var objects = new[] {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >> >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Contents 5 0 R >>\nendobj",
            "4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>\nendobj",
            BuildStreamObject(5, Encoding.ASCII.GetBytes(streamContent))
        };

        return BuildPdf(objects, rootObjectNumber: 1);
    }

    private static byte[] BuildTransformedTextRedactionSource() {
        const string streamContent = "q\n2 0 0 2 100 100 cm\nBT\n/F1 12 Tf\n0 0 Td\n(Transformed secret) Tj\nET\nQ\n";
        var objects = new[] {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >> >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Contents 5 0 R >>\nendobj",
            "4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>\nendobj",
            BuildStreamObject(5, Encoding.ASCII.GetBytes(streamContent))
        };

        return BuildPdf(objects, rootObjectNumber: 1);
    }

    private static byte[] BuildSplitTransformedTextRedactionSource() {
        var objects = new[] {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >> >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Contents [5 0 R 6 0 R] >>\nendobj",
            "4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>\nendobj",
            BuildStreamObject(5, Encoding.ASCII.GetBytes("q\n2 0 0 2 100 100 cm\n")),
            BuildStreamObject(6, Encoding.ASCII.GetBytes("BT\n/F1 12 Tf\n0 0 Td\n(Split transformed secret) Tj\nET\nQ\n"))
        };

        return BuildPdf(objects, rootObjectNumber: 1);
    }

    private static byte[] BuildSplitTransformOperandRedactionSource() {
        var objects = new[] {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >> >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Contents [5 0 R 6 0 R] >>\nendobj",
            "4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>\nendobj",
            BuildStreamObject(5, Encoding.ASCII.GetBytes("q\n2 0 0 2 1")),
            BuildStreamObject(6, Encoding.ASCII.GetBytes("00 100 cm\nBT\n/F1 12 Tf\n0 0 Td\n(Split operand secret) Tj\nET\nQ\n"))
        };

        return BuildPdf(objects, rootObjectNumber: 1);
    }

    private static byte[] BuildLiteralEndTextOperatorRedactionSource() {
        string streamContent = string.Join("\n", new[] {
            "BT",
            "/F1 12 Tf",
            "72 720 Td",
            "(Visible before) Tj",
            "ET",
            "BT",
            "/F1 12 Tf",
            "72 702 Td",
            "(SSN ET 123) Tj",
            "ET",
            "BT",
            "/F1 12 Tf",
            "72 650 Td",
            "(Visible after) Tj",
            "ET"
        });
        var objects = new[] {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >> >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Contents 5 0 R >>\nendobj",
            "4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>\nendobj",
            BuildStreamObject(5, Encoding.ASCII.GetBytes(streamContent))
        };

        return BuildPdf(objects, rootObjectNumber: 1);
    }

    private static byte[] BuildToUnicodeLiteralRedactionSource() {
        string secret = "Secret account 123-45";
        string streamContent = string.Join("\n", new[] {
            "BT",
            "/F1 12 Tf",
            "72 720 Td",
            "(Visible before) Tj",
            "0 -18 Td",
            "(" + EncodeLiteralGlyphBytes(secret) + ") Tj",
            "0 -18 Td",
            "(Visible after) Tj",
            "ET"
        });
        string cmap = BuildSingleByteToUnicodeCMap(secret);
        var objects = new[] {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >> >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Contents 5 0 R >>\nendobj",
            "4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /AAAAAA+Helvetica /Encoding /WinAnsiEncoding /ToUnicode 6 0 R >>\nendobj",
            BuildStreamObject(5, Encoding.ASCII.GetBytes(streamContent)),
            BuildStreamObject(6, Encoding.ASCII.GetBytes(cmap))
        };

        return BuildPdf(objects, rootObjectNumber: 1);
    }

    private static byte[] BuildSingleTextObjectRedactionSource(string showOperation) {
        string content = "BT\n/F1 20 Tf\n72 720 Td\n" + showOperation + "\nET\n";
        return BuildTextContentRedactionSource(content);
    }

    private static byte[] BuildTextContentRedactionSource(string content) {
        var objects = new[] {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >> >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Contents 5 0 R >>\nendobj",
            "4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>\nendobj",
            BuildStreamObject(5, Encoding.ASCII.GetBytes(content))
        };
        return BuildPdf(objects, rootObjectNumber: 1);
    }

    private static byte[] BuildToUnicodeSingleTextObjectRedactionSource(string text) {
        string content = "BT\n/F1 20 Tf\n72 720 Td\n(" + EncodeLiteralGlyphBytes(text) + ") Tj\nET\n";
        string cmap = BuildSingleByteToUnicodeCMap(text);
        var objects = new[] {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >> >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Contents 5 0 R >>\nendobj",
            "4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /AAAAAA+Helvetica /Encoding /WinAnsiEncoding /ToUnicode 6 0 R >>\nendobj",
            BuildStreamObject(5, Encoding.ASCII.GetBytes(content)),
            BuildStreamObject(6, Encoding.ASCII.GetBytes(cmap))
        };
        return BuildPdf(objects, rootObjectNumber: 1);
    }

    private static byte[] BuildType0ToUnicodeSingleTextObjectRedactionSource(string text) {
        string content = "BT\n/F1 20 Tf\n72 720 Td\n<" + EncodeTwoByteGlyphHex(text) + "> Tj\nET\n";
        string cmap = BuildTwoByteToUnicodeCMap(text);
        string widths = string.Join(" ", Enumerable.Repeat("600", text.Length));
        var objects = new[] {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >> >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Contents 5 0 R >>\nendobj",
            "4 0 obj\n<< /Type /Font /Subtype /Type0 /BaseFont /AAAAAA+Composite /Encoding /Identity-H /DescendantFonts [7 0 R] /ToUnicode 6 0 R >>\nendobj",
            BuildStreamObject(5, Encoding.ASCII.GetBytes(content)),
            BuildStreamObject(6, Encoding.ASCII.GetBytes(cmap)),
            "7 0 obj\n<< /Type /Font /Subtype /CIDFontType2 /BaseFont /AAAAAA+Composite /CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> /FontDescriptor 8 0 R /DW 600 /W [1 [" + widths + "]] >>\nendobj",
            "8 0 obj\n<< /Type /FontDescriptor /FontName /AAAAAA+Composite /Flags 32 /FontBBox [0 -200 1000 900] /ItalicAngle 0 /Ascent 800 /Descent -200 /CapHeight 700 /StemV 80 /FontFile2 9 0 R >>\nendobj",
            BuildStreamObject(9, Encoding.ASCII.GetBytes("embedded-font-program"))
        };
        return BuildPdf(objects, rootObjectNumber: 1);
    }

    private static byte[] BuildType0ToUnicodeLigatureRedactionSource() {
        const string content = "BT\n/F1 20 Tf\n72 720 Td\n<000100020003> Tj\nET\n";
        const string cmap = "/CIDInit /ProcSet findresource begin\n" +
            "12 dict begin\n" +
            "begincmap\n" +
            "3 beginbfchar\n" +
            "<0001> <0041>\n" +
            "<0002> <00660069>\n" +
            "<0003> <005A>\n" +
            "endbfchar\n" +
            "endcmap\n" +
            "CMapName currentdict /CMap defineresource pop\n" +
            "end\n" +
            "end\n";
        var objects = new[] {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >> >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Contents 5 0 R >>\nendobj",
            "4 0 obj\n<< /Type /Font /Subtype /Type0 /BaseFont /AAAAAA+Composite /Encoding /Identity-H /DescendantFonts [7 0 R] /ToUnicode 6 0 R >>\nendobj",
            BuildStreamObject(5, Encoding.ASCII.GetBytes(content)),
            BuildStreamObject(6, Encoding.ASCII.GetBytes(cmap)),
            "7 0 obj\n<< /Type /Font /Subtype /CIDFontType2 /BaseFont /AAAAAA+Composite /CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> /FontDescriptor 8 0 R /DW 600 /W [1 [600 600 600]] >>\nendobj",
            "8 0 obj\n<< /Type /FontDescriptor /FontName /AAAAAA+Composite /Flags 32 /FontBBox [0 -200 1000 900] /ItalicAngle 0 /Ascent 800 /Descent -200 /CapHeight 700 /StemV 80 /FontFile2 9 0 R >>\nendobj",
            BuildStreamObject(9, Encoding.ASCII.GetBytes("embedded-font-program"))
        };
        return BuildPdf(objects, rootObjectNumber: 1);
    }

    private static byte[] BuildType0UnicodeSpaceWordSpacingPdf() {
        const string content = "BT /F1 20 Tf 100 Tw 72 720 Td <0001> Tj ET";
        const string cmap = "/CIDInit /ProcSet findresource begin\n12 dict begin\nbegincmap\n1 beginbfchar\n<0001> <0020>\nendbfchar\nendcmap\nCMapName currentdict /CMap defineresource pop\nend\nend\n";
        var objects = new[] {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >> >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Contents 5 0 R >>\nendobj",
            "4 0 obj\n<< /Type /Font /Subtype /Type0 /BaseFont /AAAAAA+Composite /Encoding /Identity-H /DescendantFonts [7 0 R] /ToUnicode 6 0 R >>\nendobj",
            BuildStreamObject(5, Encoding.ASCII.GetBytes(content)),
            BuildStreamObject(6, Encoding.ASCII.GetBytes(cmap)),
            "7 0 obj\n<< /Type /Font /Subtype /CIDFontType2 /BaseFont /AAAAAA+Composite /CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> /DW 600 /W [1 [600]] >>\nendobj"
        };
        return BuildPdf(objects, rootObjectNumber: 1);
    }

    private static byte[] BuildFormXObjectRedactionSource() {
        string pageContent = "q\n1 0 0 1 72 700 cm\n/Fm1 Do\nQ\n";
        string formContent = "BT\n/F1 12 Tf\n0 0 Td\n(Secret account 123-45) Tj\nET\n";
        var objects = new[] {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Resources << /XObject << /Fm1 6 0 R >> >> /Contents 5 0 R >>\nendobj",
            "4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>\nendobj",
            BuildStreamObject(5, Encoding.ASCII.GetBytes(pageContent)),
            BuildStreamObject(6, Encoding.ASCII.GetBytes(formContent), "/Type /XObject /Subtype /Form /BBox [0 0 220 40] /Resources << /Font << /F1 4 0 R >> >>")
        };

        return BuildPdf(objects, rootObjectNumber: 1);
    }

    private static byte[] BuildInheritedTextStateFormRedactionSource() {
        const string pageContent = "/F1 20 Tf\n12 Tc\nq\n1 0 0 1 72 700 cm\n/Fm1 Do\nQ\n";
        const string formContent = "BT\n0 0 Td\n(Alpha secret Omega) Tj\nET\n";
        var objects = new[] {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Resources << /Font << /F1 4 0 R >> /XObject << /Fm1 6 0 R >> >> /Contents 5 0 R >>\nendobj",
            "4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>\nendobj",
            BuildStreamObject(5, Encoding.ASCII.GetBytes(pageContent)),
            BuildStreamObject(6, Encoding.ASCII.GetBytes(formContent), "/Type /XObject /Subtype /Form /BBox [0 0 320 60]")
        };
        return BuildPdf(objects, rootObjectNumber: 1);
    }

    private static byte[] BuildExtGStateFontFormRedactionSource() {
        const string pageContent = "/F1 4 Tf\n/GSFont gs\nq\n1 0 0 1 72 700 cm\n/Fm1 Do\nQ\n";
        const string formContent = "BT\n0 0 Td\n(Alpha secret Omega) Tj\nET\n";
        var objects = new[] {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Resources << /Font << /F1 4 0 R /F2 7 0 R >> /ExtGState << /GSFont 8 0 R >> /XObject << /Fm1 6 0 R >> >> /Contents 5 0 R >>\nendobj",
            "4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>\nendobj",
            BuildStreamObject(5, Encoding.ASCII.GetBytes(pageContent)),
            BuildStreamObject(6, Encoding.ASCII.GetBytes(formContent), "/Type /XObject /Subtype /Form /BBox [0 0 420 60]"),
            "7 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Courier /Encoding /WinAnsiEncoding >>\nendobj",
            "8 0 obj\n<< /Type /ExtGState /Font [/F2 30] >>\nendobj"
        };
        return BuildPdf(objects, rootObjectNumber: 1);
    }

    private static byte[] BuildInlineImagePayloadBeforeFormRedactionSource() {
        const string inlinePayload = "q 1 0 0 1 72 700 cm /FmDormant Do Q";
        string pageContent = $"BI /W {inlinePayload.Length} /H 1 /BPC 8 /CS /G ID {inlinePayload} EI\n" +
            "q 1 0 0 1 72 700 cm /FmVisible Do Q\n";
        const string dormantFormContent = "BT /F1 12 Tf 0 0 Td (Dormant secret) Tj ET";
        const string visibleFormContent = "BT /F1 12 Tf 0 0 Td (Visible secret) Tj ET";
        var objects = new[] {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Resources << /Font << /F1 4 0 R >> /XObject << /FmDormant 6 0 R /FmVisible 7 0 R >> >> /Contents 5 0 R >>\nendobj",
            "4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>\nendobj",
            BuildStreamObject(5, Encoding.ASCII.GetBytes(pageContent)),
            BuildStreamObject(6, Encoding.ASCII.GetBytes(dormantFormContent), "/Type /XObject /Subtype /Form /BBox [0 0 220 40] /Resources << /Font << /F1 4 0 R >> >>"),
            BuildStreamObject(7, Encoding.ASCII.GetBytes(visibleFormContent), "/Type /XObject /Subtype /Form /BBox [0 0 220 40] /Resources << /Font << /F1 4 0 R >> >>")
        };
        return BuildPdf(objects, rootObjectNumber: 1);
    }

    private static byte[] BuildQuotedTextStateFormRedactionSource() {
        const string pageContent = "/F1 20 Tf\nBT 72 750 Td 30 12 (Prelude) \" ET\nq\n1 0 0 1 72 700 cm\n/Fm1 Do\nQ\n";
        const string formContent = "BT\n0 0 Td\n(Alpha secret Omega) Tj\nET\n";
        var objects = new[] {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Resources << /Font << /F1 4 0 R >> /XObject << /Fm1 6 0 R >> >> /Contents 5 0 R >>\nendobj",
            "4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>\nendobj",
            BuildStreamObject(5, Encoding.ASCII.GetBytes(pageContent)),
            BuildStreamObject(6, Encoding.ASCII.GetBytes(formContent), "/Type /XObject /Subtype /Form /BBox [0 0 420 60]")
        };
        return BuildPdf(objects, rootObjectNumber: 1);
    }

    private static byte[] BuildCollidingInheritedFontFormRedactionSource() {
        const string pageContent = "/F1 20 Tf\nq\n1 0 0 1 72 700 cm\n/Fm1 Do\nQ\n";
        const string formContent = "BT\n0 0 Td\n(Alpha secret Omega) Tj\nET\n";
        var objects = new[] {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Resources << /Font << /F1 4 0 R >> /XObject << /Fm1 6 0 R >> >> /Contents 5 0 R >>\nendobj",
            "4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>\nendobj",
            BuildStreamObject(5, Encoding.ASCII.GetBytes(pageContent)),
            BuildStreamObject(6, Encoding.ASCII.GetBytes(formContent), "/Type /XObject /Subtype /Form /BBox [0 0 320 60] /Resources << /Font << /F1 7 0 R >> >>"),
            "7 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Symbol >>\nendobj"
        };
        return BuildPdf(objects, rootObjectNumber: 1);
    }

    private static byte[] BuildNestedInheritedTextStateFormRedactionSource() {
        const string pageContent = "/F1 20 Tf\nq\n1 0 0 1 72 700 cm\n/FmOuter Do\nQ\n";
        const string outerFormContent = "12 Tc\n/FmInner Do\n";
        const string innerFormContent = "BT\n0 0 Td\n(Alpha secret Omega) Tj\nET\n";
        var objects = new[] {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Resources << /Font << /F1 4 0 R >> /XObject << /FmOuter 6 0 R >> >> /Contents 5 0 R >>\nendobj",
            "4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>\nendobj",
            BuildStreamObject(5, Encoding.ASCII.GetBytes(pageContent)),
            BuildStreamObject(6, Encoding.ASCII.GetBytes(outerFormContent), "/Type /XObject /Subtype /Form /BBox [0 0 320 60] /Resources << /Font << /F1 4 0 R >> /XObject << /FmInner 7 0 R >> >>"),
            BuildStreamObject(7, Encoding.ASCII.GetBytes(innerFormContent), "/Type /XObject /Subtype /Form /BBox [0 0 320 60]")
        };
        return BuildPdf(objects, rootObjectNumber: 1);
    }

    private static byte[] BuildSplitFormTransformOperandRedactionSource() {
        const string formContent = "BT\n/F1 12 Tf\n0 0 Td\n(Split form secret) Tj\nET\n";
        var objects = new[] {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Resources << /XObject << /Fm1 7 0 R >> >> /Contents [5 0 R 6 0 R] >>\nendobj",
            "4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>\nendobj",
            BuildStreamObject(5, Encoding.ASCII.GetBytes("q\n1 0 0 1 1")),
            BuildStreamObject(6, Encoding.ASCII.GetBytes("00 100 cm\n/Fm1 Do\nQ\n")),
            BuildStreamObject(7, Encoding.ASCII.GetBytes(formContent), "/Type /XObject /Subtype /Form /BBox [0 0 220 40] /Resources << /Font << /F1 4 0 R >> >>")
        };

        return BuildPdf(objects, rootObjectNumber: 1);
    }

    private static byte[] BuildNestedFormXObjectRedactionSource() {
        string pageContent = "q\n1 0 0 1 72 700 cm\n/FmOuter Do\nQ\n";
        string outerFormContent = "q\n1 0 0 1 0 0 cm\n/FmInner Do\nQ\n";
        string innerFormContent = "BT\n/F1 12 Tf\n0 0 Td\n(Nested secret account 123-45) Tj\nET\n";
        var objects = new[] {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Resources << /XObject << /FmOuter 6 0 R >> >> /Contents 5 0 R >>\nendobj",
            "4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>\nendobj",
            BuildStreamObject(5, Encoding.ASCII.GetBytes(pageContent)),
            BuildStreamObject(6, Encoding.ASCII.GetBytes(outerFormContent), "/Type /XObject /Subtype /Form /BBox [0 0 220 40] /Resources << /XObject << /FmInner 7 0 R >> >>"),
            BuildStreamObject(7, Encoding.ASCII.GetBytes(innerFormContent), "/Type /XObject /Subtype /Form /BBox [0 0 220 40] /Resources << /Font << /F1 4 0 R >> >>")
        };

        return BuildPdf(objects, rootObjectNumber: 1);
    }

    private static byte[] BuildLeakingGraphicsStateRedactionSource() {
        string streamContent = string.Join("\n", new[] {
            "q",
            "0 0 1 1 re W n",
            "BT",
            "/F1 12 Tf",
            "72 720 Td",
            "(Visible page text) Tj",
            "ET"
        });
        var objects = new[] {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >> >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Contents 5 0 R >>\nendobj",
            "4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>\nendobj",
            BuildStreamObject(5, Encoding.ASCII.GetBytes(streamContent))
        };

        return BuildPdf(objects, rootObjectNumber: 1);
    }

    private static byte[] BuildSplitContentStateRedactionSource() {
        var objects = new[] {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >> >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Contents [5 0 R 6 0 R] >>\nendobj",
            "4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>\nendobj",
            BuildStreamObject(5, Encoding.ASCII.GetBytes("/F1 12 Tf")),
            BuildStreamObject(6, Encoding.ASCII.GetBytes("72 720 Td (Visible split text) Tj"))
        };

        return BuildPdf(objects, rootObjectNumber: 1);
    }

    private static byte[] BuildDuplicateRedactionSource() {
        string streamContent = string.Join("\n", new[] {
            "BT",
            "/F1 12 Tf",
            "72 740 Td",
            "(Visible before) Tj",
            "ET",
            "BT",
            "/F1 12 Tf",
            "72 700 Td",
            "(Repeat secret) Tj",
            "ET",
            "BT",
            "/F1 12 Tf",
            "72 660 Td",
            "(Repeat secret) Tj",
            "ET",
            "BT",
            "/F1 12 Tf",
            "72 620 Td",
            "(Visible after) Tj",
            "ET"
        });
        var objects = new[] {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >> >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Contents 5 0 R >>\nendobj",
            "4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>\nendobj",
            BuildStreamObject(5, Encoding.ASCII.GetBytes(streamContent))
        };

        return BuildPdf(objects, rootObjectNumber: 1);
    }

    private static byte[] BuildAnnotationAppearanceRedactionSource() {
        string pageContent = "BT\n/F1 12 Tf\n72 720 Td\n(Visible page text) Tj\nET";
        string appearanceContent = "BT /F1 12 Tf 0 0 Td (Old sensitive appearance Sensitive annotation) Tj ET";
        var objects = new[] {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >> >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Resources << /Font << /F1 4 0 R >> >> /Annots [6 0 R] /Contents 5 0 R >>\nendobj",
            "4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>\nendobj",
            BuildStreamObject(5, Encoding.ASCII.GetBytes(pageContent)),
            "6 0 obj\n<< /Type /Annot /Subtype /FreeText /Rect [20 20 60 60] /Contents (Sensitive annotation) /AP << /N 7 0 R >> >>\nendobj",
            BuildStreamObject(7, Encoding.ASCII.GetBytes(appearanceContent), "/Type /XObject /Subtype /Form /BBox [0 0 40 40] /Resources << /Font << /F1 4 0 R >> >>")
        };

        return BuildPdf(objects, rootObjectNumber: 1);
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

    private static byte[] BuildInheritedFormXObjectTextPdf() {
        var objects = new List<string> {
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] /Resources << /XObject << /Fm1 5 0 R >> >> >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Contents 6 0 R >>",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            BuildStream("BT\n/F1 12 Tf\n0 0 Td\n(Inherited form secret) Tj\nET", "/Type /XObject /Subtype /Form /BBox [0 0 200 50] /Resources << /Font << /F1 4 0 R >> >>"),
            BuildStream("q\n1 0 0 1 100 100 cm\n/Fm1 Do\nQ")
        };

        return Encoding.ASCII.GetBytes(BuildPdf(objects));
    }

    private static byte[] BuildSharedPageContentPdf() {
        var objects = new List<string> {
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Count 2 /Kids [3 0 R 4 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Resources << /Font << /F1 5 0 R >> >> /Contents 6 0 R >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Resources << /Font << /F1 5 0 R >> >> /Contents 6 0 R >>",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            BuildStream("BT\n/F1 12 Tf\n72 120 Td\n(Shared page secret) Tj\nET")
        };

        return Encoding.ASCII.GetBytes(BuildPdf(objects));
    }

    private static byte[] BuildSharedFormXObjectTextPdf() {
        var objects = new List<string> {
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Count 2 /Kids [3 0 R 4 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Resources << /XObject << /Fm1 6 0 R >> >> /Contents 7 0 R >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Resources << /XObject << /Fm1 6 0 R >> >> /Contents 8 0 R >>",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            BuildStream("BT\n/F1 12 Tf\n0 0 Td\n(Shared form secret) Tj\nET", "/Type /XObject /Subtype /Form /BBox [0 0 200 50] /Resources << /Font << /F1 5 0 R >> >>"),
            BuildStream("q\n1 0 0 1 100 100 cm\n/Fm1 Do\nQ"),
            BuildStream("q\n1 0 0 1 100 100 cm\n/Fm1 Do\nQ")
        };

        return Encoding.ASCII.GetBytes(BuildPdf(objects));
    }

    private static byte[] BuildRepeatedFormXObjectTextPdf() {
        var objects = new List<string> {
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Resources << /XObject << /Fm1 6 0 R >> >> /Contents [7 0 R 8 0 R] >>",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            BuildStream("BT\n/F1 12 Tf\n0 0 Td\n(Repeated form secret) Tj\nET", "/Type /XObject /Subtype /Form /BBox [0 0 200 50] /Resources << /Font << /F1 4 0 R >> >>"),
            BuildStream("q\n1 0 0 1 0 0 cm\n/FmInner Do\nQ", "/Type /XObject /Subtype /Form /BBox [0 0 200 50] /Resources << /XObject << /FmInner 5 0 R >> >>"),
            BuildStream("q\n1 0 0 1 30 220 cm\n/Fm1 Do\nQ\nq\n1 0 0 1 3"),
            BuildStream("0 80 cm\n/Fm1 Do\nQ")
        };

        return Encoding.ASCII.GetBytes(BuildPdf(objects));
    }

    private static byte[] BuildRepeatedNestedFormWithIndirectResourcesPdf() {
        var objects = new List<string> {
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Resources << /XObject << /FmOuter 6 0 R >> >> /Contents 8 0 R >>",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            BuildStream("BT\n/F1 12 Tf\n0 0 Td\n(Indirect nested secret) Tj\nET", "/Type /XObject /Subtype /Form /BBox [0 0 200 50] /Resources << /Font << /F1 4 0 R >> >>"),
            BuildStream("q\n/FmInner Do\nQ", "/Type /XObject /Subtype /Form /BBox [0 0 200 50] /Resources 7 0 R"),
            "<< /XObject << /FmInner 5 0 R >> >>",
            BuildStream("q\n1 0 0 1 30 220 cm\n/FmOuter Do\nQ\nq\n1 0 0 1 30 80 cm\n/FmOuter Do\nQ")
        };

        return Encoding.ASCII.GetBytes(BuildPdf(objects));
    }

    private static byte[] BuildAliasedFormXObjectTextPdf() {
        var objects = new List<string> {
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Resources << /XObject << /FmUnused 5 0 R /FmPainted 5 0 R >> >> /Contents 6 0 R >>",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            BuildStream("BT\n/F1 12 Tf\n0 0 Td\n(Aliased form secret) Tj\nET", "/Type /XObject /Subtype /Form /BBox [0 0 200 50] /Resources << /Font << /F1 4 0 R >> >>"),
            BuildStream("q\n1 0 0 1 100 100 cm\n/FmPainted Do\nQ")
        };

        return Encoding.ASCII.GetBytes(BuildPdf(objects));
    }

    private static byte[] BuildLargeTextPdf() {
        var objects = new List<string> {
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 400 300] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            BuildStream("BT\n/F1 48 Tf\n72 100 Td\n(Large secret heading) Tj\nET")
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

    private static byte[] BuildIndirectAnnotationWithPopupPdf() {
        var objects = new List<string> {
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Annots [4 0 R 5 0 R] /Contents 6 0 R >>",
            "<< /Type /Annot /Subtype /Text /Rect [20 20 40 40] /Contents (Keep parent note) /Popup 5 0 R >>",
            "<< /Type /Annot /Subtype /Popup /Rect [100 100 160 160] /Parent 4 0 R >>",
            BuildStream("BT\n/F1 12 Tf\n72 720 Td\n(Annotation carrier) Tj\nET")
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

    private static PdfRedactionArea BuildAreaForSubstring(PdfTextSpan span, string text) {
        int start = span.Text.IndexOf(text, StringComparison.Ordinal);
        Assert.True(start >= 0);
        Assert.True(PdfTextAdvanceProjection.TryGetResolvedBoundaries(span, out double[] boundaries));
        double offset = boundaries[start];
        double end = boundaries[start + text.Length];
        PdfTextSpanBounds bounds = PdfTextSpanGeometry.GetAxisAlignedBounds(span, Math.Min(offset, end), Math.Abs(end - offset));
        return new PdfRedactionArea(1, bounds.Left + 0.05D, bounds.Bottom + 0.05D, bounds.Width - 0.1D, bounds.Height - 0.1D, "partial");
    }

    private static PdfRedactionArea[] FindAreasForText(byte[] pdf, string text) {
        return PdfDocumentReadResult.Load(pdf)
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

    private static PdfRedactionArea FindAreaForTextOccurrence(byte[] pdf, string text, int occurrenceFromTop) {
        PdfLogicalTextBlock block = PdfDocumentReadResult.Load(pdf)
            .TextBlocks
            .Where(item => item.Text.Contains(text, StringComparison.Ordinal))
            .OrderByDescending(item => item.BaselineY)
            .ElementAt(occurrenceFromTop);

        double x = Math.Min(block.XStart, block.XEnd) - 2D;
        double width = Math.Abs(block.XEnd - block.XStart) + 4D;
        return new PdfRedactionArea(block.PageNumber, x, block.BaselineY - 14D, width, 20D, "secret");
    }

    private static string BuildSingleByteToUnicodeCMap(string text) {
        var builder = new StringBuilder();
        builder.Append("/CIDInit /ProcSet findresource begin\n");
        builder.Append("12 dict begin\n");
        builder.Append("begincmap\n");
        builder.Append(text.Length.ToString(System.Globalization.CultureInfo.InvariantCulture)).Append(" beginbfchar\n");
        for (int i = 0; i < text.Length; i++) {
            builder.Append('<')
                .Append((i + 1).ToString("X2", System.Globalization.CultureInfo.InvariantCulture))
                .Append("> <")
                .Append(((int)text[i]).ToString("X4", System.Globalization.CultureInfo.InvariantCulture))
                .Append(">\n");
        }

        builder.Append("endbfchar\n");
        builder.Append("endcmap\n");
        builder.Append("CMapName currentdict /CMap defineresource pop\n");
        builder.Append("end\n");
        builder.Append("end\n");
        return builder.ToString();
    }

    private static string BuildTwoByteToUnicodeCMap(string text) {
        var builder = new StringBuilder();
        builder.Append("/CIDInit /ProcSet findresource begin\n");
        builder.Append("12 dict begin\n");
        builder.Append("begincmap\n");
        builder.Append(text.Length.ToString(System.Globalization.CultureInfo.InvariantCulture)).Append(" beginbfchar\n");
        for (int i = 0; i < text.Length; i++) {
            builder.Append('<')
                .Append((i + 1).ToString("X4", System.Globalization.CultureInfo.InvariantCulture))
                .Append("> <")
                .Append(((int)text[i]).ToString("X4", System.Globalization.CultureInfo.InvariantCulture))
                .Append(">\n");
        }

        builder.Append("endbfchar\n");
        builder.Append("endcmap\n");
        builder.Append("CMapName currentdict /CMap defineresource pop\n");
        builder.Append("end\n");
        builder.Append("end\n");
        return builder.ToString();
    }

    private static string EncodeTwoByteGlyphHex(string text) {
        var builder = new StringBuilder(text.Length * 4);
        for (int i = 0; i < text.Length; i++) {
            builder.Append((i + 1).ToString("X4", System.Globalization.CultureInfo.InvariantCulture));
        }

        return builder.ToString();
    }

    private static string EncodeLiteralGlyphBytes(string text) {
        var builder = new StringBuilder(text.Length * 4);
        for (int i = 0; i < text.Length; i++) {
            builder.Append('\\')
                .Append(Convert.ToString(i + 1, 8).PadLeft(3, '0'));
        }

        return builder.ToString();
    }

    private static string BuildStreamObject(int objectNumber, byte[] streamBytes, string extraDictionary = "") {
        string dictionary = "<< /Length " + streamBytes.Length.ToString(System.Globalization.CultureInfo.InvariantCulture);
        if (!string.IsNullOrWhiteSpace(extraDictionary)) {
            dictionary += " " + extraDictionary.Trim();
        }

        dictionary += " >>";
        return objectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) + " 0 obj\n" +
            dictionary + "\nstream\n" +
            Encoding.ASCII.GetString(streamBytes) + "\nendstream\nendobj";
    }

    private static byte[] BuildPdf(IReadOnlyList<string> objects, int rootObjectNumber) {
        var offsets = new Dictionary<int, int>();
        using var stream = new MemoryStream();
        using var writer = new StreamWriter(stream, Encoding.ASCII, 1024, leaveOpen: true);

        writer.WriteLine("%PDF-1.4");
        writer.Flush();
        int maxObjectNumber = 0;
        foreach (string obj in objects) {
            int objectNumber = ReadObjectNumber(obj);
            offsets[objectNumber] = (int)stream.Position;
            maxObjectNumber = Math.Max(maxObjectNumber, objectNumber);
            writer.WriteLine(obj);
            writer.Flush();
        }

        int xrefOffset = (int)stream.Position;
        writer.WriteLine("xref");
        writer.WriteLine("0 " + (maxObjectNumber + 1).ToString(System.Globalization.CultureInfo.InvariantCulture));
        writer.WriteLine("0000000000 65535 f ");
        for (int i = 1; i <= maxObjectNumber; i++) {
            if (offsets.TryGetValue(i, out int offset)) {
                writer.WriteLine(offset.ToString("D10", System.Globalization.CultureInfo.InvariantCulture) + " 00000 n ");
            } else {
                writer.WriteLine("0000000000 65535 f ");
            }
        }

        writer.WriteLine("trailer");
        writer.WriteLine("<< /Size " + (maxObjectNumber + 1).ToString(System.Globalization.CultureInfo.InvariantCulture) + " /Root " + rootObjectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) + " 0 R >>");
        writer.WriteLine("startxref");
        writer.WriteLine(xrefOffset.ToString(System.Globalization.CultureInfo.InvariantCulture));
        writer.WriteLine("%%EOF");
        writer.Flush();
        return stream.ToArray();
    }

    private static int ReadObjectNumber(string obj) {
        int space = obj.IndexOf(' ');
        return int.Parse(obj.Substring(0, space), System.Globalization.CultureInfo.InvariantCulture);
    }

    private static string ExtractLogicalText(byte[] pdf) {
        return string.Join(
            Environment.NewLine,
            PdfDocumentReadResult.Load(pdf).TextBlocks.Select(item => item.Text));
    }
}
