using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfRedactionEvidenceTests {
    [Fact]
    public void ApplyWithEvidenceReturnsSourceBoundActualVersusPlannedProof() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Remove evidence marker"))
            .Paragraph(paragraph => paragraph.Text("Retain ordinary content"))
            .ToBytes();
        PdfDocument document = PdfDocument.Load(source);
        PdfRedactionPlan plan = document.Redactions.Search(
            new PdfRedactionSearchOptions().AddLiteral("Remove evidence marker"));
        var verification = new PdfRedactionVerificationOptions {
            RequireCompleteStreamInspection = true,
            CheckManagedRendering = true
        }.RequireRemovedText("Remove evidence marker")
            .RequireRetainedText("Retain ordinary content");

        PdfRedactionApplyResult result = document.Redactions.ApplyWithEvidence(
            plan,
            verificationOptions: verification);
        string outputText = result.ToDocument().Read().Text;

        Assert.True(result.IsVerified, result.Evidence.Summary);
        Assert.Equal(PdfMutationOperation.Redact, result.MutationPlan.Operation);
        Assert.Equal(PdfMutationExecutionMode.FullRewrite, result.MutationPlan.ExecutionMode);
        Assert.Equal(plan.SourceSha256, result.Evidence.SourceSha256);
        Assert.NotEqual(result.Evidence.SourceSha256, result.Evidence.OutputSha256);
        Assert.NotEmpty(result.Evidence.Items);
        Assert.All(result.Evidence.Items, item => Assert.Equal(PdfRedactionEvidenceStatus.VerifiedAbsent, item.Status));
        Assert.Equal(result.Evidence.Items.Count, result.Evidence.VerifiedAbsentCount);
        Assert.Empty(result.Evidence.ResidualMatches);
        Assert.True(result.Evidence.Verification.CompleteStreamInspectionRequired);
        Assert.True(result.Evidence.Verification.ManagedRenderingChecked);
        Assert.DoesNotContain("Remove evidence marker", outputText, StringComparison.Ordinal);
        Assert.Contains("Retain ordinary content", outputText, StringComparison.Ordinal);
    }

    [Fact]
    public void ApplyWithEvidenceHandlesUnknownOperatorsAndAnUnbalancedTopLevelTransform() {
        const string content = "2 0 0 2 10 10 cm\n99 /Noise madeUpOperator\nBT /F1 12 Tf 10 10 Td (HOSTILE-MARKER) Tj ET\nBT /F1 12 Tf 10 40 Td (RETAIN-MARKER) Tj ET";
        byte[] source = BuildPdf(
            pageResources: "<< /Font << /F1 5 0 R >> >>",
            pageContent: content,
            additionalObjects: new[] { "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj" });
        PdfDocument document = PdfDocument.Load(source);
        PdfRedactionPlan plan = document.Redactions.Search(new PdfRedactionSearchOptions().AddLiteral("HOSTILE-MARKER"));
        var verification = new PdfRedactionVerificationOptions { RequireCompleteStreamInspection = true }
            .RequireRemovedText("HOSTILE-MARKER")
            .RequireRetainedText("RETAIN-MARKER");

        PdfRedactionApplyResult result = document.Redactions.ApplyWithEvidence(plan, verificationOptions: verification);

        Assert.True(result.IsVerified, result.Evidence.Summary);
        Assert.Contains("RETAIN-MARKER", result.ToDocument().Read().Text, StringComparison.Ordinal);
        Assert.DoesNotContain("HOSTILE-MARKER", result.ToDocument().Read().Text, StringComparison.Ordinal);
    }

    [Fact]
    public void ApplyWithEvidenceMarksItemsInconclusiveWhenConfiguredProofFails() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Remove this marker"))
            .ToBytes();
        PdfDocument document = PdfDocument.Load(source);
        PdfRedactionPlan plan = document.Redactions.Search(
            new PdfRedactionSearchOptions().AddLiteral("Remove this marker"));
        var verification = new PdfRedactionVerificationOptions()
            .RequireRemovedText("Remove this marker")
            .RequireRetainedText("Required retained marker is absent");

        PdfRedactionApplyResult result = document.Redactions.ApplyWithEvidence(
            plan,
            verificationOptions: verification);

        Assert.False(result.IsVerified);
        Assert.All(result.Evidence.Items, item => Assert.Equal(PdfRedactionEvidenceStatus.Inconclusive, item.Status));
        Assert.Equal(result.Evidence.Items.Count, result.Evidence.InconclusiveCount);
        Assert.Throws<InvalidOperationException>(() => result.ThrowIfUnverified());
    }

    [Fact]
    public void ApplyWithEvidenceReportsReviewedAnnotationRemoval() {
        byte[] source = BuildAnnotationPdf();
        PdfDocument document = PdfDocument.Load(source);
        var area = new PdfRedactionArea(1, 20, 20, 40, 40, "reviewed annotation");
        PdfRedactionPlan plan = document.Redactions.Plan([area]);

        PdfRedactionApplyResult result = document.Redactions.ApplyWithEvidence(plan);

        PdfRedactionEvidenceItem item = Assert.Single(result.Evidence.Items);
        Assert.True(result.IsVerified, result.Evidence.Summary);
        Assert.Equal(PdfRedactionMatchKind.Annotation, item.ReviewedMatch.Kind);
        Assert.Equal(PdfRedactionEvidenceStatus.VerifiedAbsent, item.Status);
        Assert.Empty(PdfInspector.Inspect(result.Pdf).GetAnnotationsBySubtype("Text"));
    }

    [Fact]
    public void ApplyWithEvidenceDoesNotVerifyWidgetWhileItsParentFieldValueRemains() {
        byte[] source = BuildParentOwnedWidgetValuePdf();
        PdfDocument document = PdfDocument.Load(source);
        PdfRedactionPlan plan = document.Redactions.Plan([
            new PdfRedactionArea(1, 40, 20, 140, 30, "reviewed widget")
        ]);

        PdfRedactionApplyResult result = document.Redactions.ApplyWithEvidence(plan);

        PdfRedactionEvidenceItem item = Assert.Single(
            result.Evidence.Items,
            evidence => evidence.ReviewedMatch.Kind == PdfRedactionMatchKind.Annotation);
        Assert.False(result.IsVerified);
        Assert.Equal("Widget", item.ReviewedMatch.Subtype);
        Assert.Equal(PdfRedactionEvidenceStatus.Inconclusive, item.Status);
        Assert.Contains(
            PdfReadDocument.Open(result.Pdf).FormFields,
            field => field.Name == "SensitiveField" && field.Values.Contains("PARENT-FIELD-SECRET"));
    }

    [Fact]
    public void ApplyWithEvidenceDoesNotDeleteTerminalFieldWidgetOrVerifyItsRetainedValue() {
        byte[] source = BuildTerminalFieldWidgetValuePdf();
        PdfDocument document = PdfDocument.Load(source);
        PdfRedactionPlan plan = document.Redactions.Plan([
            new PdfRedactionArea(1, 40, 20, 140, 30, "reviewed terminal widget")
        ]);

        PdfRedactionApplyResult result = document.Redactions.ApplyWithEvidence(plan);

        PdfRedactionEvidenceItem item = Assert.Single(
            result.Evidence.Items,
            evidence => evidence.ReviewedMatch.Kind == PdfRedactionMatchKind.Annotation);
        Assert.False(result.IsVerified);
        Assert.Equal(PdfRedactionEvidenceStatus.Inconclusive, item.Status);
        Assert.Contains(
            PdfReadDocument.Open(result.Pdf).FormFields,
            field => field.Name == "TerminalSensitiveField" && field.Values.Contains("TERMINAL-FIELD-SECRET"));
    }

    [Fact]
    public void ApplyWithEvidenceVerifiesWidgetWhenTheReviewedFieldOwnerIsRemoved() {
        byte[] source = BuildParentOwnedWidgetValuePdf();
        PdfDocument document = PdfDocument.Load(source);
        PdfRedactionPlan plan = document.Redactions.Search(
            new PdfRedactionSearchOptions().AddFormField("SensitiveField"));

        PdfRedactionApplyResult result = document.Redactions.ApplyWithEvidence(plan);

        Assert.True(result.IsVerified, result.Evidence.Summary);
        Assert.All(result.Evidence.Items, item => Assert.Equal(PdfRedactionEvidenceStatus.VerifiedAbsent, item.Status));
        Assert.DoesNotContain(PdfReadDocument.Open(result.Pdf).FormFields, field => field.Name == "SensitiveField");
        Assert.DoesNotContain("PARENT-FIELD-SECRET", PdfEncoding.Latin1GetString(result.Pdf), StringComparison.Ordinal);
    }

    [Fact]
    public void AreaRedactionUsesDeclaredFontWidthsWhenDecidingTextIntersection() {
        const string retained = "AAAAAAAAAA";
        byte[] source = BuildPdf(
            pageResources: "<< /Font << /F1 5 0 R >> >>",
            pageContent: "BT /F1 12 Tf 10 100 Td (" + retained + ") Tj ET",
            additionalObjects: new[] {
                "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /FirstChar 65 /LastChar 65 /Widths [100] >>\nendobj"
            });
        PdfDocument document = PdfDocument.Load(source);
        PdfRedactionPlan plan = document.Redactions.Plan([
            new PdfRedactionArea(1, 40, 95, 5, 10, "outside narrow glyph run")
        ]);

        PdfRedactionApplyResult result = document.Redactions.ApplyWithEvidence(plan);

        Assert.DoesNotContain(plan.Matches, match => match.Kind == PdfRedactionMatchKind.TextBlock);
        Assert.Contains(retained, result.ToDocument().Read().Text, StringComparison.Ordinal);
    }

    [Fact]
    public void AreaRedactionUsesNestedFormFontWidthsWhenDecidingTextIntersection() {
        const string retained = "AAAAAAAAAA";
        byte[] source = BuildPdf(
            pageResources: "<< /XObject << /Fm 6 0 R >> >>",
            pageContent: "/Fm Do",
            additionalObjects: new[] {
                "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /FirstChar 65 /LastChar 65 /Widths [100] >>\nendobj",
                StreamObject(6, "/Type /XObject /Subtype /Form /BBox [0 0 240 180] /Resources << /Font << /F1 5 0 R >> >>", "BT /F1 12 Tf 10 100 Td (" + retained + ") Tj ET")
            });
        PdfDocument document = PdfDocument.Load(source);
        PdfRedactionPlan plan = document.Redactions.Plan([
            new PdfRedactionArea(1, 40, 95, 5, 10, "outside narrow nested glyph run")
        ]);

        PdfRedactionApplyResult result = document.Redactions.ApplyWithEvidence(plan);

        Assert.DoesNotContain(plan.Matches, match => match.Kind == PdfRedactionMatchKind.TextBlock);
        Assert.Contains(retained, result.ToDocument().Read().Text, StringComparison.Ordinal);
    }

    [Fact]
    public void ApplyWithEvidenceExpandsSourceTightLimitsForGeneratedOutput() {
        byte[] source = PdfDocument.Create(pdf => pdf.Content(content => content
                .Paragraph(paragraph => paragraph.Text("Retained source content"))),
                new PdfOptions { CompressContentStreams = false })
            .ToBytes();
        PdfStream[] sourceStreams = PdfSyntax.ParseObjects(source).Map.Values
            .Select(static item => item.Value)
            .OfType<PdfStream>()
            .ToArray();
        int maximumSourceStreamBytes = sourceStreams.Max(static stream => stream.Data.Length);
        long totalSourceStreamBytes = sourceStreams.Sum(static stream => (long)stream.Data.Length);
        var readOptions = new PdfLoadOptions {
            Limits = new PdfReadLimits {
                MaxInputBytes = source.LongLength,
                MaxRawStreamBytes = maximumSourceStreamBytes,
                MaxDecodedStreamBytes = maximumSourceStreamBytes,
                MaxTotalDecodedStreamBytes = totalSourceStreamBytes,
                MaxPageContentBytes = maximumSourceStreamBytes,
                MaxRetainedContentBytes = totalSourceStreamBytes,
                MaxContentOperations = 64,
                MaxContentOperands = 128
            }
        };
        PdfRedactionArea[] areas = Enumerable.Range(0, 64)
            .Select(index => new PdfRedactionArea(
                1,
                20D + ((index % 8) * 24D),
                20D + ((index / 8) * 24D),
                8D,
                8D,
                "reviewed area " + index.ToString(System.Globalization.CultureInfo.InvariantCulture)))
            .ToArray();
        PdfDocument document = PdfDocument.Load(source, readOptions);
        PdfRedactionPlan plan = document.Redactions.Plan(areas);
        PdfRedactionApplyResult result = document.Redactions.ApplyWithEvidence(
            plan,
            verificationOptions: new PdfRedactionVerificationOptions {
                RequireCompleteStreamInspection = true,
                CheckManagedRendering = false
            });

        Assert.True(result.Pdf.LongLength > source.LongLength);
        Assert.True(result.IsVerified, result.Evidence.Summary);
        Assert.Contains("Retained source content", result.ToDocument().Read().Text, StringComparison.Ordinal);
    }

    [Fact]
    public void RedactionPlanningStopsAtSelfReferencingFormXObject() {
        const string formContent = "BT /F1 12 Tf 20 40 Td (CYCLE-MARKER) Tj ET\n/Loop Do";
        byte[] source = BuildPdf(
            pageResources: "<< /Font << /F1 5 0 R >> /XObject << /Loop 6 0 R >> >>",
            pageContent: "/Loop Do",
            additionalObjects: new[] {
                "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj",
                StreamObject(6, "/Type /XObject /Subtype /Form /BBox [0 0 200 200] /Resources << /Font << /F1 5 0 R >> /XObject << /Loop 6 0 R >> >>", formContent)
            });

        PdfRedactionPlan plan = PdfDocument.Load(source).Redactions.Search(
            new PdfRedactionSearchOptions().AddLiteral("CYCLE-MARKER"));

        Assert.True(plan.IsReviewable);
        Assert.Single(plan.Matches, match => match.Kind == PdfRedactionMatchKind.TextBlock);
    }

    private static byte[] BuildPdf(string pageResources, string pageContent, IReadOnlyList<string> additionalObjects) {
        var objects = new List<string> {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 240 180] /Resources " + pageResources + " /Contents 4 0 R >>\nendobj",
            StreamObject(4, string.Empty, pageContent)
        };
        objects.AddRange(additionalObjects);
        string pdf = "%PDF-1.7\n" + string.Join("\n", objects) + "\ntrailer\n<< /Root 1 0 R /Size " + (objects.Count + 1) + " >>\n%%EOF";
        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildAnnotationPdf() {
        const string content = "BT /F1 12 Tf 20 150 Td (VISIBLE-CONTENT) Tj ET";
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 240 180] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R /Annots [6 0 R] >>\nendobj",
            StreamObject(4, string.Empty, content),
            "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj",
            "6 0 obj\n<< /Type /Annot /Subtype /Text /Rect [20 20 60 60] /Contents (SENSITIVE-NOTE) /F 4 >>\nendobj",
            "trailer\n<< /Root 1 0 R /Size 7 >>",
            "%%EOF"
        });
        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildParentOwnedWidgetValuePdf() {
        const string content = "BT /F1 12 Tf 20 150 Td (VISIBLE-CONTENT) Tj ET";
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R /AcroForm << /Fields [6 0 R] >> >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 240 180] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R /Annots [7 0 R] >>\nendobj",
            StreamObject(4, string.Empty, content),
            "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj",
            "6 0 obj\n<< /FT /Tx /T (SensitiveField) /V (PARENT-FIELD-SECRET) /Kids [7 0 R] >>\nendobj",
            "7 0 obj\n<< /Type /Annot /Subtype /Widget /Parent 6 0 R /Rect [40 20 180 50] /P 3 0 R /F 4 >>\nendobj",
            "trailer\n<< /Root 1 0 R /Size 8 >>",
            "%%EOF"
        });
        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildTerminalFieldWidgetValuePdf() {
        const string content = "BT /F1 12 Tf 20 150 Td (VISIBLE-CONTENT) Tj ET";
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R /AcroForm << /Fields [6 0 R] >> >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 240 180] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R /Annots [6 0 R] >>\nendobj",
            StreamObject(4, string.Empty, content),
            "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj",
            "6 0 obj\n<< /Type /Annot /Subtype /Widget /FT /Tx /T (TerminalSensitiveField) /V (TERMINAL-FIELD-SECRET) /Rect [40 20 180 50] /P 3 0 R /F 4 >>\nendobj",
            "trailer\n<< /Root 1 0 R /Size 7 >>",
            "%%EOF"
        });
        return Encoding.ASCII.GetBytes(pdf);
    }

    private static string StreamObject(int objectNumber, string dictionaryEntries, string content) {
        int length = Encoding.ASCII.GetByteCount(content);
        string entries = string.IsNullOrWhiteSpace(dictionaryEntries) ? string.Empty : dictionaryEntries + " ";
        return objectNumber + " 0 obj\n<< " + entries + "/Length " + length + " >>\nstream\n" + content + "\nendstream\nendobj";
    }
}
