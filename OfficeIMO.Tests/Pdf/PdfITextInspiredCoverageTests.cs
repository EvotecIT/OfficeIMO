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
    public void IncrementalUpdater_AppendsFormAppearanceStreams() {
        byte[] pdf = BuildCoveragePdf();

        byte[] updated = PdfIncrementalUpdater.UpdateFormFields(pdf, new Dictionary<string, string> {
            ["Name"] = "Grace"
        }, new PdfIncrementalFormFieldUpdateOptions {
            GenerateAppearanceStreams = true,
            KeepNeedAppearances = false
        });

        string text = PdfEncoding.Latin1GetString(updated);
        PdfDocumentInfo info = PdfInspector.Inspect(updated);
        PdfFormField field = Assert.Single(info.FormFields);
        Assert.Equal("Grace", field.Value);
        Assert.Equal(false, info.AcroFormNeedAppearances);
        Assert.Contains("/AP", text, StringComparison.Ordinal);
        Assert.Contains("/Subtype /Form", text, StringComparison.Ordinal);
        Assert.Contains("/BaseFont /Helvetica", text, StringComparison.Ordinal);
    }

    [Fact]
    public void IncrementalUpdater_PreservesChangedObjectGeneration() {
        byte[] pdf = BuildNonZeroGenerationFormPdf();

        byte[] updated = PdfIncrementalUpdater.UpdateFormFields(pdf, new Dictionary<string, string> {
            ["Name"] = "Grace"
        });

        string text = PdfEncoding.Latin1GetString(updated);
        Assert.True(updated.Length > pdf.Length);
        Assert.Contains("5 2 obj", text, StringComparison.Ordinal);
        Assert.Contains("/V <4772616365>", text, StringComparison.Ordinal);
        Assert.Contains("/Fields [ 5 2 R ]", text, StringComparison.Ordinal);
        Assert.Contains("00002 n", text, StringComparison.Ordinal);
    }

    [Fact]
    public void IncrementalUpdater_AppendsParentObjectForDirectChildFields() {
        byte[] pdf = BuildDirectChildFormPdf();

        byte[] updated = PdfIncrementalUpdater.UpdateFormFields(pdf, new Dictionary<string, string> {
            ["Parent.Child"] = "Grace"
        });

        PdfDocumentInfo info = PdfInspector.Inspect(updated);
        PdfFormField field = Assert.Single(info.FormFields);
        Assert.Equal("Parent.Child", field.Name);
        Assert.Equal("Grace", field.Value);
        Assert.Contains("5 0 obj", PdfEncoding.Latin1GetString(updated), StringComparison.Ordinal);
    }

    [Fact]
    public void IncrementalUpdater_AllowsDocMDPCertifiedFormFillWhenPermissionPermits() {
        byte[] pdf = BuildDocMdpFormPdf(permissionLevel: 2);

        PdfAppendOnlyMutationReport plan = PdfIncrementalUpdater.AnalyzeAppendOnlyMutation(pdf);
        Assert.True(plan.CanAppendFormFields);
        Assert.False(plan.CanAppendMetadata);
        Assert.Contains("SignedDocMDPFormFill", plan.Warnings);

        byte[] updated = PdfIncrementalUpdater.UpdateFormFields(pdf, new Dictionary<string, string> {
            ["Name"] = "Grace"
        }, new PdfIncrementalFormFieldUpdateOptions {
            GenerateAppearanceStreams = true,
            KeepNeedAppearances = false
        });

        PdfDocumentInfo info = PdfInspector.Inspect(updated);
        PdfFormField textField = Assert.Single(info.FormFields, static field => field.Name == "Name");
        Assert.Equal("Grace", textField.Value);
        Assert.True(info.Security.HasSignatures);
        Assert.True(info.Security.HasDocMDPPermissions);
        Assert.True(info.Security.HasIncrementalUpdates);
        Assert.Contains("/Prev", PdfEncoding.Latin1GetString(updated), StringComparison.Ordinal);
    }

    [Fact]
    public void IncrementalUpdater_BlocksDocMDPCertifiedFormFillWhenPermissionForbidsChanges() {
        byte[] pdf = BuildDocMdpFormPdf(permissionLevel: 1);

        PdfAppendOnlyMutationReport plan = PdfIncrementalUpdater.AnalyzeAppendOnlyMutation(pdf);

        Assert.False(plan.CanAppendFormFields);
        Assert.Contains("DocMDP", plan.Blockers);
        Assert.Throws<NotSupportedException>(() => PdfIncrementalUpdater.UpdateFormFields(pdf, new Dictionary<string, string> {
            ["Name"] = "Grace"
        }, new PdfIncrementalFormFieldUpdateOptions {
            GenerateAppearanceStreams = true,
            KeepNeedAppearances = false
        }));
    }

    [Fact]
    public void IncrementalUpdater_BlocksDocMDPCertifiedFormFillWhenSignatureFieldLockIncludesRequestedField() {
        byte[] pdf = BuildDocMdpFormPdf(
            permissionLevel: 2,
            signatureFieldLock: " /Lock << /Type /SigFieldLock /Action /Include /Fields [(Name)] >>");

        PdfAppendOnlyMutationReport plan = PdfIncrementalUpdater.AnalyzeAppendOnlyMutation(pdf);
        Assert.False(plan.CanAppendFormFields);
        Assert.Contains("SignatureFieldLock", plan.Blockers);

        var exception = Assert.Throws<NotSupportedException>(() => PdfIncrementalUpdater.UpdateFormFields(pdf, new Dictionary<string, string> {
            ["Name"] = "Grace"
        }, new PdfIncrementalFormFieldUpdateOptions {
            GenerateAppearanceStreams = true,
            KeepNeedAppearances = false
        }));
        Assert.Contains("SignatureFieldLock", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void IncrementalUpdater_BlocksDocMDPCertifiedFormFillWhenSignatureFieldLockIncludesParentField() {
        byte[] pdf = BuildDocMdpFormPdf(
            permissionLevel: 2,
            signatureFieldLock: " /Lock << /Type /SigFieldLock /Action /Include /Fields [(Billing)] >>",
            textFieldName: "Billing.Name");

        var exception = Assert.Throws<NotSupportedException>(() => PdfIncrementalUpdater.UpdateFormFields(pdf, new Dictionary<string, string> {
            ["Billing.Name"] = "Grace"
        }, new PdfIncrementalFormFieldUpdateOptions {
            GenerateAppearanceStreams = true,
            KeepNeedAppearances = false
        }));
        Assert.Contains("SignatureFieldLock", exception.Message, StringComparison.Ordinal);
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
        byte[] pdf = BuildAnnotationEditPdf();

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
        Assert.False(annotation.HasAction);
        Assert.False(annotation.HasAdditionalActions);
        Assert.False(annotation.HasChainedActions);
        Assert.Equal(4, annotation.Flags);

        PdfAnnotationEditResult removed = PdfAnnotationEditor.RemoveAnnotations(updated.Bytes, new PdfAnnotationRemovalOptions {
            Subtype = "Text"
        });

        Assert.True(removed.Applied);
        PdfDocumentInfo info = PdfInspector.Inspect(removed.Bytes);
        Assert.Empty(info.GetAnnotationsBySubtype("Text"));
        Assert.Equal(0, info.AnnotationCount);
    }

    [Fact]
    public void AnnotationEditor_UpdateRejectsNonAnnotationSubtypeObjects() {
        byte[] pdf = BuildAnnotationEditPdf();

        var exception = Assert.Throws<ArgumentException>(() => PdfAnnotationEditor.UpdateAnnotation(pdf, 4, new PdfAnnotationUpdateOptions {
            Contents = "Do not write annotation keys into a font"
        }));

        Assert.Contains("PDF annotation object was not found", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void AnnotationEditor_UpdateRejectsNonFiniteColorComponents() {
        byte[] pdf = BuildAnnotationEditPdf();

        var exception = Assert.Throws<ArgumentException>(() => PdfAnnotationEditor.UpdateAnnotation(pdf, 6, new PdfAnnotationUpdateOptions {
            Color = new[] { 0D, double.NaN, 1D }
        }));

        Assert.Contains("finite RGB", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void AnnotationEditor_RemovesDirectAnnotationDictionaries() {
        byte[] pdf = BuildDirectAnnotationPdf();

        PdfAnnotationEditResult removed = PdfAnnotationEditor.RemoveAnnotations(pdf, new PdfAnnotationRemovalOptions {
            Subtype = "Text"
        });

        string text = PdfEncoding.Latin1GetString(removed.Bytes);
        Assert.True(removed.Applied);
        Assert.Equal(1, removed.AffectedAnnotationCount);
        Assert.DoesNotContain("Direct note", text, StringComparison.Ordinal);
        Assert.Equal(0, PdfInspector.Inspect(removed.Bytes).AnnotationCount);
    }

    [Fact]
    public void AnnotationEditor_RemovesPopupReferencesWhenRemovingParentAnnotation() {
        byte[] pdf = BuildAnnotationWithLinkedPopupPdf();

        PdfAnnotationEditResult removed = PdfAnnotationEditor.RemoveAnnotations(pdf, new PdfAnnotationRemovalOptions {
            Subtype = "Text"
        });

        PdfDocumentInfo info = PdfInspector.Inspect(removed.Bytes);
        string text = PdfEncoding.Latin1GetString(removed.Bytes);
        Assert.True(removed.Applied);
        Assert.Equal(2, removed.AffectedAnnotationCount);
        Assert.Equal(0, info.AnnotationCount);
        Assert.DoesNotContain("/Subtype /Popup", text, StringComparison.Ordinal);
        Assert.DoesNotContain("/Popup 7 0 R", text, StringComparison.Ordinal);
    }

    [Fact]
    public void AnnotationEditor_ClearsParentPopupReferenceWhenRemovingPopupAnnotation() {
        byte[] pdf = BuildAnnotationWithLinkedPopupPdf();

        PdfAnnotationEditResult removed = PdfAnnotationEditor.RemoveAnnotations(pdf, new PdfAnnotationRemovalOptions {
            Subtype = "Popup"
        });

        PdfDocumentInfo info = PdfInspector.Inspect(removed.Bytes);
        string text = PdfEncoding.Latin1GetString(removed.Bytes);
        Assert.True(removed.Applied);
        Assert.Equal(1, removed.AffectedAnnotationCount);
        Assert.Single(info.GetAnnotationsBySubtype("Text"));
        Assert.Empty(info.GetAnnotationsBySubtype("Popup"));
        Assert.DoesNotContain("/Subtype /Popup", text, StringComparison.Ordinal);
        Assert.DoesNotContain("/Popup 7 0 R", text, StringComparison.Ordinal);
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

    private static byte[] BuildAnnotationEditPdf() {
        byte[] contentBytes = Encoding.ASCII.GetBytes("BT\n/F1 12 Tf\n72 720 Td\n(Annotation editing) Tj\nET\n");
        var objects = new List<string> {
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Resources << /Font << /F1 4 0 R >> >> /Annots [6 0 R] /Contents 5 0 R >>",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            BuildStream(contentBytes),
            "<< /Type /Annot /Subtype /Text /Rect [20 20 40 40] /Contents (Review note) /F 132 /NM (Note-1) /T (Reviewer) /M (D:20260622090000Z) /C [1 0 0] /A << /S /URI /URI (https://example.com) >> >>"
        };

        return Encoding.ASCII.GetBytes(BuildPdf(objects));
    }

    private static byte[] BuildDirectAnnotationPdf() {
        byte[] contentBytes = Encoding.ASCII.GetBytes("BT\n/F1 12 Tf\n72 720 Td\n(Direct annotation) Tj\nET\n");
        var objects = new List<string> {
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Resources << /Font << /F1 4 0 R >> >> /Annots [<< /Type /Annot /Subtype /Text /Rect [20 20 40 40] /Contents (Direct note) >>] /Contents 5 0 R >>",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            BuildStream(contentBytes)
        };

        return Encoding.ASCII.GetBytes(BuildPdf(objects));
    }

    private static byte[] BuildAnnotationWithLinkedPopupPdf() {
        byte[] contentBytes = Encoding.ASCII.GetBytes("BT\n/F1 12 Tf\n72 720 Td\n(Popup annotation) Tj\nET\n");
        var objects = new List<string> {
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Resources << /Font << /F1 4 0 R >> >> /Annots [6 0 R 7 0 R] /Contents 5 0 R >>",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            BuildStream(contentBytes),
            "<< /Type /Annot /Subtype /Text /Rect [20 20 40 40] /Contents (Parent note) /Popup 7 0 R >>",
            "<< /Type /Annot /Subtype /Popup /Rect [45 20 120 80] /Parent 6 0 R >>"
        };

        return Encoding.ASCII.GetBytes(BuildPdf(objects));
    }

    private static byte[] BuildNonZeroGenerationFormPdf() {
        byte[] contentBytes = Encoding.ASCII.GetBytes("BT /F1 12 Tf 72 720 Td (Generated form) Tj ET");
        var objects = new List<(int Number, int Generation, string Body)> {
            (1, 0, "<< /Type /Catalog /Pages 2 0 R /AcroForm 6 0 R >>"),
            (2, 0, "<< /Type /Pages /Count 1 /Kids [3 0 R] >>"),
            (3, 0, "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Resources << /Font << /F1 4 0 R >> >> /Annots [5 2 R] /Contents 7 0 R >>"),
            (4, 0, "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>"),
            (5, 2, "<< /Type /Annot /Subtype /Widget /FT /Tx /T (Name) /V (Ada) /Rect [50 50 180 70] /F 4 >>"),
            (6, 0, "<< /Fields [5 2 R] >>"),
            (7, 0, BuildStream(contentBytes))
        };

        var builder = new StringBuilder();
        builder.AppendLine("%PDF-1.7");
        foreach ((int number, int generation, string body) in objects) {
            builder.Append(number.ToString(CultureInfo.InvariantCulture))
                .Append(' ')
                .Append(generation.ToString(CultureInfo.InvariantCulture))
                .AppendLine(" obj");
            builder.AppendLine(body);
            builder.AppendLine("endobj");
        }

        builder.AppendLine("trailer");
        builder.AppendLine("<< /Root 1 0 R /Size 8 >>");
        builder.AppendLine("startxref");
        builder.AppendLine("123");
        builder.AppendLine("%%EOF");
        return Encoding.ASCII.GetBytes(builder.ToString());
    }

    private static byte[] BuildDirectChildFormPdf() {
        var objects = new List<string> {
            "<< /Type /Catalog /Pages 2 0 R /AcroForm 6 0 R >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Contents 4 0 R >>",
            BuildStream(Encoding.ASCII.GetBytes("BT /F1 12 Tf 72 720 Td (Direct child form) Tj ET")),
            "<< /T (Parent) /Kids [<< /Type /Annot /Subtype /Widget /FT /Tx /T (Child) /V (Ada) /Rect [50 50 180 70] /F 4 >>] >>",
            "<< /Fields [5 0 R] >>"
        };

        return Encoding.ASCII.GetBytes(BuildPdf(objects));
    }

    private static byte[] BuildDocMdpFormPdf(int permissionLevel, string signatureFieldLock = "", string textFieldName = "Name") {
        var objects = new List<string> {
            "<< /Type /Catalog /Pages 2 0 R /AcroForm 8 0 R /Perms << /DocMDP 7 0 R >> >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Resources << /Font << /F1 4 0 R >> >> /Annots [5 0 R 6 0 R] /Contents 9 0 R >>",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            "<< /Type /Annot /Subtype /Widget /FT /Tx /T (" + textFieldName + ") /V (Ada) /Rect [50 50 180 70] /F 4 >>",
            "<< /FT /Sig /T (Approval) /V 7 0 R /Subtype /Widget /Rect [10 10 120 40]" + signatureFieldLock + " >>",
            "<< /Type /Sig /Filter /Adobe.PPKLite /SubFilter /adbe.pkcs7.detached /Name (Alice) /ByteRange [0 10 20 30] /Contents <001122> /Reference [<< /TransformMethod /DocMDP /TransformParams << /Type /TransformParams /V /1.2 /P " + permissionLevel.ToString(CultureInfo.InvariantCulture) + " >> >>] >>",
            "<< /Fields [5 0 R 6 0 R] /SigFlags 3 >>",
            BuildStream(Encoding.ASCII.GetBytes("BT /F1 12 Tf 72 720 Td (Signed form) Tj ET"))
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
