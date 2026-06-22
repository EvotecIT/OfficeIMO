using System.Globalization;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfITextInspiredCoverageTests {
    [Fact]
    public void Inspect_ReportsFormsAnnotationsPageBoxesTaggedFontsAndAppendPlan() {
        byte[] pdf = BuildCoveragePdf();

        byte[] appendablePdf = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Append-only metadata plan"))
            .ToBytes();
        PdfAppendOnlyMutationReport appendPlan = PdfIncrementalUpdater.AnalyzeAppendOnlyMutation(appendablePdf);
        Assert.True(appendPlan.CanAppendMetadata);
        Assert.True(appendPlan.CanAppendFormFields);
        Assert.Contains("Metadata", appendPlan.SupportedActions);
        Assert.Contains("FormFill", appendPlan.SupportedActions);

        PdfDocumentInfo info = PdfInspector.Inspect(pdf);
        Assert.Equal(1, info.TextFormFieldCount);
        Assert.Equal(1, info.RequiredFormFieldCount);
        Assert.Equal(1, info.ReadOnlyFormFieldCount);
        Assert.Equal(1, info.FormWidgetCount);
        Assert.True(info.HasProductionPageBoxes);
        Assert.Equal(1, info.TrimBoxPageCount);
        Assert.Equal(1, info.BleedBoxPageCount);
        Assert.Equal(1, info.ArtBoxPageCount);
        Assert.Equal(1, info.ActiveAnnotationCount);
        Assert.Equal(1, info.RiskyAnnotationActionCount);
        Assert.True(info.AnnotationSubtypeCounts["Text"] >= 1);

        PdfAnnotation note = Assert.Single(info.GetAnnotationsBySubtype("Text"));
        Assert.Equal("Review note", note.Contents);
        Assert.Equal("Note-1", note.Name);
        Assert.Equal("Reviewer", note.Title);
        Assert.True(note.IsLocked);
        Assert.True(note.HasColor);
        Assert.Equal("Launch", Assert.Single(note.AdditionalActions).ActionType);

        PdfTaggedContentInfo tagged = Assert.IsType<PdfTaggedContentInfo>(info.TaggedContent);
        Assert.True(tagged.HasRoleMap);
        Assert.True(tagged.HasDeepTaggedPdfEvidence);
        Assert.Equal(1, tagged.LanguageElementCount);
        Assert.Equal(0, tagged.AlternateTextElementCount);
        Assert.Equal(1, tagged.FigureWithoutAlternateTextCount);

        PdfDiagnosticReport diagnostics = PdfDiagnostics.Analyze(pdf);
        Assert.True(diagnostics.FontCount >= 2);
        Assert.Contains(diagnostics.Fonts, font => font.ObjectNumber == 4 && font.IsStandardBase14Font);
        Assert.Contains(diagnostics.Fonts, font => font.ObjectNumber == 14 && font.HasEmbeddedFontFile && font.EmbeddedFontFileKind == "FontFile2");
        Assert.Contains(diagnostics.Fonts, font => font.ObjectNumber == 14 && font.RepairReadiness == "Ready");

        PdfOptimizationReport optimization = PdfDiagnostics.AnalyzeOptimization(pdf);
        Assert.True(optimization.StreamCount > 0);
        Assert.True(optimization.TotalStreamBytes > 0);
        Assert.True(optimization.LargestStreamBytes > 0);
        Assert.True(optimization.FindingCount >= 0);
    }

    [Fact]
    public void AssessProof_ReportsMissingExternalValidationStatus() {
        var options = new PdfOptions {
            ComplianceProfile = PdfComplianceProfile.PdfA3B
        };

        PdfComplianceProofReport proof = PdfComplianceAnalyzer.AssessProof(PdfComplianceProfile.PdfA3B, options);

        Assert.True(proof.RequiresExternalValidation);
        Assert.True(proof.RequiredExternalValidatorCount > 0);
        Assert.Equal("InternalGaps", proof.ProofStatus);
        Assert.Contains("Missing external validation", proof.ExternalProofSummary, StringComparison.Ordinal);
        Assert.False(proof.CanClaimConformance);
    }

    [Fact]
    public void IncrementalUpdater_AppendsSimpleFormFieldRevision() {
        byte[] pdf = BuildCoveragePdf();

        byte[] updated = PdfIncrementalUpdater.UpdateFormFields(pdf, new Dictionary<string, string> {
            ["Name"] = "Grace"
        });

        Assert.True(updated.Length > pdf.Length);
        PdfDocumentInfo info = PdfInspector.Inspect(updated);
        PdfFormField field = Assert.Single(info.FormFields);
        Assert.Equal("Grace", field.Value);
        Assert.True(info.AcroFormNeedAppearances);
        Assert.True(info.Security.HasIncrementalUpdates);
    }

    [Fact]
    public void PageEditor_SetsProductionBoundaryBoxes() {
        byte[] pdf = PdfPageGeometrySupport.BuildPageGeometryPdf();

        byte[] updated = PdfPageEditor.SetPageBox(pdf, "TrimBox", 12, 14, 222, 244);

        PdfPageGeometry geometry = PdfInspector.Inspect(updated).Pages[0].Geometry!;
        Assert.NotNull(geometry.TrimBox);
        Assert.Equal(12, geometry.TrimBox!.Left);
        Assert.Equal(14, geometry.TrimBox.Bottom);
        Assert.Equal(222, geometry.TrimBox.Right);
        Assert.Equal(244, geometry.TrimBox.Top);
    }

    [Fact]
    public void SecurityInfo_ReportsObjectStreamRewriteReadiness() {
        byte[] pdf = Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 0 /Kids [] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /ObjStm /N 0 /First 0 /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 4 >>",
            "startxref",
            "123",
            "%%EOF"
        }));

        PdfDocumentSecurityInfo security = PdfInspector.Probe(pdf).Security;

        Assert.True(security.HasObjectStreams);
        Assert.True(security.BlocksOfficeIMOFullRewriteMutation);
    }

    [Fact]
    public void AnnotationEditor_UpdatesAndRemovesAnnotations() {
        byte[] pdf = BuildCoveragePdf();

        PdfAnnotationEditResult updated = PdfAnnotationEditor.UpdateAnnotation(pdf, 6, new PdfAnnotationUpdateOptions {
            Contents = "Updated note",
            Title = "Lead reviewer",
            Name = "Note-2",
            Flags = 4,
            Color = new[] { 0D, 0.5D, 1D },
            RemoveActions = true
        });

        Assert.True(updated.Applied);
        PdfAnnotation annotation = Assert.Single(PdfInspector.Inspect(updated.Bytes).GetAnnotationsBySubtype("Text"));
        Assert.Equal("Updated note", annotation.Contents);
        Assert.Equal("Lead reviewer", annotation.Title);
        Assert.Equal("Note-2", annotation.Name);
        Assert.False(annotation.HasActions);
        Assert.Equal(4, annotation.Flags);

        PdfAnnotationEditResult removed = PdfAnnotationEditor.RemoveAnnotations(updated.Bytes, new PdfAnnotationRemovalOptions {
            Subtype = "Text"
        });

        Assert.True(removed.Applied);
        PdfDocumentInfo info = PdfInspector.Inspect(removed.Bytes);
        Assert.Empty(info.GetAnnotationsBySubtype("Text"));
        Assert.Equal(1, info.FormWidgetCount);
    }

    [Fact]
    public void ExternalValidationResult_FromExitCodeImportsProcessOutcome() {
        PdfExternalValidationResult passed = PdfExternalValidationResult.FromExitCode(
            PdfExternalValidatorKind.VeraPdf,
            0,
            "veraPDF",
            "passed",
            profile: "PDF/A-3b",
            executablePath: "verapdf",
            arguments: "--format text file.pdf");

        PdfExternalValidationResult failed = PdfExternalValidationResult.FromExitCode(
            PdfExternalValidatorKind.PdfUaValidator,
            2,
            "pdfua",
            "failed",
            profile: "PDF/UA-1");

        Assert.Equal(PdfExternalValidationStatus.Passed, passed.Status);
        Assert.Equal(0, passed.ExitCode);
        Assert.Equal("verapdf", passed.ExecutablePath);
        Assert.Equal("--format text file.pdf", passed.Arguments);
        Assert.Equal(PdfExternalValidationStatus.Failed, failed.Status);
        Assert.Equal(2, failed.ExitCode);
    }

    private static byte[] BuildCoveragePdf() {
        string longText = new string('A', 512);
        byte[] contentBytes = Encoding.ASCII.GetBytes("BT\n/F1 12 Tf\n72 720 Td\n(" + longText + ") Tj\nET\n");
        byte[] fontBytes = { 1, 2, 3, 4 };
        var objects = new List<string> {
            "<< /Type /Catalog /Pages 2 0 R /AcroForm 8 0 R /MarkInfo << /Marked true >> /StructTreeRoot 10 0 R >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /CropBox [0 0 290 290] /BleedBox [5 5 295 295] /TrimBox [10 10 290 290] /ArtBox [20 20 280 280] /Resources << /Font << /F1 4 0 R /F2 14 0 R >> >> /Annots [6 0 R 7 0 R] /Contents 5 0 R >>",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            BuildStream(contentBytes),
            "<< /Type /Annot /Subtype /Text /Rect [20 20 40 40] /Contents (Review note) /F 132 /NM (Note-1) /T (Reviewer) /M (D:20260622090000Z) /C [1 0 0] /AA << /E << /S /Launch /F (tool.exe) >> >> >>",
            "<< /Type /Annot /Subtype /Widget /FT /Tx /T (Name) /V (Ada) /Ff 3 /Rect [50 50 180 70] /F 4 >>",
            "<< /Fields [7 0 R] /SigFlags 2 >>",
            "<< /Type /FontDescriptor /FontName /EmbeddedSans /FontFile2 15 0 R >>",
            "<< /Type /StructTreeRoot /K [11 0 R] /ParentTree 13 0 R /ParentTreeNextKey 1 /RoleMap << /Custom /P >> >>",
            "<< /Type /StructElem /S /Document /P 10 0 R /K [12 0 R] /Lang (en-US) >>",
            "<< /Type /StructElem /S /Figure /P 11 0 R /K 0 >>",
            "<< /Nums [0 12 0 R] >>",
            "<< /Type /Font /Subtype /TrueType /BaseFont /EmbeddedSans /FontDescriptor 9 0 R >>",
            BuildStream(fontBytes)
        };

        return Encoding.ASCII.GetBytes(BuildPdf(objects));
    }

    private static string BuildStream(byte[] data) =>
        "<< /Length " + data.Length.ToString(CultureInfo.InvariantCulture) + " >>\nstream\n" +
        Encoding.ASCII.GetString(data) +
        "\nendstream";

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
}
