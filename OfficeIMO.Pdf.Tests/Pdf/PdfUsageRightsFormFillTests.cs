using System.Globalization;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfUsageRightsFormFillTests {
    private static readonly Dictionary<string, string> AfterValues = new(StringComparer.Ordinal) { ["Name"] = "after" };

    [Fact]
    public void RepeatedUsageRightsTransformRetainsDistinctPermissionsInOrder() {
        string rights = "[/FillIn " + string.Join(" ", Enumerable.Range(0, 512).Select(index => "/Grant" + index)) + "]";
        byte[] source = BuildUsageRightsFormPdf(rights, documentSignature: false);
        source = PdfDocumentObjectGraphRewriter.Rewrite(source, null, null, (objects, security) => {
            var signature = (PdfDictionary)objects[6].Value;
            var references = (PdfArray)signature.Items["Reference"];
            int number = objects.Keys.Max() + 1;
            objects.Add(number, new PdfIndirectObject(number, 0, references.Items[0]));
            references.Items.Clear();
            for (int index = 0; index < 128; index++) references.Items.Add(new PdfReference(number, 0));
            return security.InfoObjectNumber;
        });

        PdfDocumentSecurityInfo securityInfo = PdfSyntax.ReadDocumentSecurityInfo(source);
        Assert.Equal(513, securityInfo.UsageRightsFormPermissions.Count);
        Assert.Equal("FillIn", securityInfo.UsageRightsFormPermissions[0]);
        Assert.Equal("Grant511", securityInfo.UsageRightsFormPermissions[512]);
        Assert.True(securityInfo.UsageRightsAllowFormFillIn);
    }

    [Fact]
    public void ExcessiveUsageRightsPermissionsCannotGrantFormFillIn() {
        string rights = "[/FillIn " + string.Join(" ", Enumerable.Range(0, 8_192).Select(index => "/Grant" + index)) + "]";
        byte[] source = BuildUsageRightsFormPdf(rights, documentSignature: false);

        PdfDocumentSecurityInfo securityInfo = PdfSyntax.ReadDocumentSecurityInfo(source);
        Assert.Empty(securityInfo.UsageRightsFormPermissions);
        Assert.False(securityInfo.UsageRightsAllowFormFillIn);
        Assert.False(PdfInspector.Preflight(source).CanAppendFormFieldRevision);
    }

    [Fact]
    public void UsageRightsFillInGrant_AppendsFormRevisionAndKeepsSignedBytes() {
        byte[] original = BuildUsageRightsFormPdf("[/FillIn /Import]", documentSignature: false);

        PdfDocumentPreflight preflight = PdfInspector.Preflight(original);
        PdfDocumentSecurityInfo security = preflight.Probe.Security;
        Assert.Equal(new[] { "FillIn", "Import" }, security.UsageRightsFormPermissions);
        Assert.True(security.UsageRightsAllowFormFillIn);
        Assert.True(security.HasOnlyUsageRightsSignatures);
        Assert.False(preflight.CanFillSimpleFormFields);
        Assert.True(preflight.CanAppendFormFieldRevision);
        Assert.True(preflight.CanAppendOnlyMutate);
        Assert.False(preflight.CanAppendMetadataRevision);
        Assert.False(preflight.CanPrepareExternalSignatureRevision);
        Assert.Contains("The usage-rights signature grants form fill-in (/Form /FillIn) and is the only signature; form values can be appended as a new revision that keeps the signed bytes.", preflight.SecurityDiagnostics);
        Assert.Contains("Usage-rights signatures permit only append-only form fill-in revisions; other append-only changes would exceed the granted rights.", preflight.SecurityDiagnostics);
        Assert.Empty(preflight.AppendOnlyMutationDiagnostics);
        Assert.Contains("A full rewrite would invalidate the existing signature; fill the form through an append-only form-field revision instead.",
            preflight.GetCapabilityDiagnostics(PdfPreflightCapability.FillSimpleFormFields));

        PdfAppendOnlyMutationReport report = preflight.AppendOnlyMutationReport;
        Assert.Equal(new[] { "FormFill" }, report.SupportedActions);
        Assert.Contains("UsageRights", report.Blockers);
        Assert.Contains("Signed", report.Blockers);
        Assert.Contains("UsageRightsFormFillIn", report.Warnings);
        Assert.DoesNotContain("SignedApprovalAnnotationChange", report.Warnings);
        Assert.DoesNotContain("SignedAdditionalSignature", report.Warnings);

        PdfMutationPlan plan = PdfMutationPlanner.Plan(original, PdfMutationOperation.FillFormFields, fieldNames: AfterValues.Keys);
        Assert.Equal(PdfMutationExecutionMode.AppendOnly, plan.ExecutionMode);
        Assert.False(plan.FullRewriteAvailable);
        Assert.Contains("AppendOnly.UsageRightsFormFillIn", plan.Warnings);
        Assert.Equal(PdfMutationExecutionMode.Blocked, PdfMutationPlanner.Plan(
            original,
            PdfMutationOperation.FillFormFields,
            fieldNames: AfterValues.Keys,
            executionPreference: PdfMutationExecutionPreference.RequireFullRewrite).ExecutionMode);
        Assert.False(PdfMutationPlanner.Plan(original, PdfMutationOperation.UpdateMetadata).CanExecute);

        Assert.Throws<PdfMutationBlockedException>(() => PdfDocument.Load(original).Forms.Fill(AfterValues));
        PdfOperationResult<PdfDocument> filled = PdfDocument.Load(original).Forms.FillResult(AfterValues);
        Assert.True(filled.Succeeded, string.Join(" ", filled.Diagnostics));
        byte[] updated = filled.Value!.ToBytes();

        Assert.True(updated.Length > original.Length);
        Assert.True(updated.AsSpan(0, original.Length).SequenceEqual(original));
        PdfFormField field = Assert.Single(PdfInspector.Inspect(updated).FormFields);
        Assert.Equal("after", field.Value);
        PdfDocumentSecurityInfo after = PdfSyntax.ReadDocumentSecurityInfo(updated);
        Assert.True(after.HasIncrementalUpdates);
        Assert.Equal(security.UsageRightsObjectNumbers, after.UsageRightsObjectNumbers);
        PdfSignatureInfo usageRights = Assert.Single(after.Signatures);
        Assert.Equal(Assert.Single(security.Signatures).ByteRangeValues, usageRights.ByteRangeValues);
        Assert.True(after.HasOnlyUsageRightsSignatures);
    }

    [Theory]
    [InlineData("[/Import /Export]")]
    [InlineData(null)]
    public void UsageRightsWithoutFillInGrant_KeepsFormFillBlocked(string? formRights) {
        byte[] original = BuildUsageRightsFormPdf(formRights, documentSignature: false);

        PdfDocumentPreflight preflight = PdfInspector.Preflight(original);
        Assert.False(preflight.Probe.Security.UsageRightsAllowFormFillIn);
        Assert.False(preflight.CanAppendFormFieldRevision);
        Assert.False(preflight.AppendOnlyMutationReport.CanAppendFormFields);
        Assert.Contains("UsageRights", preflight.AppendOnlyMutationReport.Blockers);
        Assert.Contains("Usage-rights entries must be preserved before append-only mutation.", preflight.AppendOnlyMutationDiagnostics);
        Assert.False(PdfMutationPlanner.Plan(original, PdfMutationOperation.FillFormFields, fieldNames: AfterValues.Keys).CanExecute);
        Assert.False(PdfDocument.Load(original).Forms.FillResult(AfterValues).Succeeded);
    }

    [Theory]
    [InlineData("/FillIn")]
    [InlineData("(FillIn)")]
    [InlineData("[(FillIn)]")]
    [InlineData("[/Import (FillIn)]")]
    [InlineData("[/FillIn (Import)]")]
    public void MalformedFormRights_DoNotAuthorizeARevision(string formRights) {
        byte[] original = BuildUsageRightsFormPdf(formRights, documentSignature: false);
        PdfDocumentPreflight preflight = PdfInspector.Preflight(original);

        Assert.Empty(preflight.Probe.Security.UsageRightsFormPermissions);
        Assert.False(preflight.Probe.Security.UsageRightsAllowFormFillIn);
        Assert.False(PdfSyntax.ReadDocumentSecurityInfo(original).UsageRightsAllowFormFillIn);
        Assert.False(preflight.CanAppendFormFieldRevision);
        Assert.False(preflight.CanAppendOnlyMutate);
        Assert.False(PdfDocument.Load(original).Forms.FillResult(AfterValues).Succeeded);
        Assert.Throws<PdfMutationBlockedException>(() => PdfDocument.Load(original).Forms.AppendRevision(AfterValues));
    }

    [Fact]
    public void DocumentSignatureAlongsideUsageRightsFillIn_KeepsFormFillBlocked() {
        byte[] original = BuildUsageRightsFormPdf("[/FillIn]", documentSignature: true);

        PdfDocumentPreflight preflight = PdfInspector.Preflight(original);
        PdfDocumentSecurityInfo security = preflight.Probe.Security;
        Assert.True(security.UsageRightsAllowFormFillIn);
        Assert.False(security.HasOnlyUsageRightsSignatures);
        Assert.False(preflight.CanAppendFormFieldRevision);
        Assert.Contains("Signed", preflight.AppendOnlyMutationReport.Blockers);
        Assert.Contains("UsageRights", preflight.AppendOnlyMutationReport.Blockers);
        Assert.DoesNotContain("UsageRightsFormFillIn", preflight.AppendOnlyMutationReport.Warnings);
        Assert.Contains("Append-only signature preservation is not implemented by OfficeIMO.Pdf yet.", preflight.AppendOnlyMutationDiagnostics);
        Assert.False(PdfMutationPlanner.Plan(original, PdfMutationOperation.FillFormFields, fieldNames: AfterValues.Keys).CanExecute);
        Assert.False(PdfDocument.Load(original).Forms.FillResult(AfterValues).Succeeded);
    }

    [Theory]
    [InlineData("[/Import]", "[/FillIn]", false)]
    [InlineData(null, "[/FillIn]", false)]
    [InlineData("[/FillIn]", "[/Import]", true)]
    public void Ur3Grant_TakesPrecedenceOverLegacyUr(string? formRights, string legacyFormRights, bool canFill) {
        byte[] original = BuildUsageRightsFormPdf(formRights, documentSignature: false, legacyFormRights: legacyFormRights);

        PdfDocumentPreflight preflight = PdfInspector.Preflight(original);
        PdfDocumentSecurityInfo security = preflight.Probe.Security;
        Assert.Equal(canFill, security.UsageRightsAllowFormFillIn);
        Assert.Equal(canFill, PdfSyntax.ReadDocumentSecurityInfo(original).UsageRightsAllowFormFillIn);
        Assert.True(security.HasOnlyUsageRightsSignatures);
        Assert.Equal(canFill, preflight.CanAppendFormFieldRevision);
        Assert.Equal(canFill, PdfDocument.Load(original).Forms.FillResult(AfterValues).Succeeded);
        Assert.False(PdfMutationPlanner.Plan(original, PdfMutationOperation.UpdateMetadata).CanExecute);
    }

    [Fact]
    public void LegacyUrGrant_AllowsFillInWhenUr3IsAbsent() {
        byte[] original = BuildUsageRightsFormPdf("[/FillIn]", documentSignature: false, usageRightsKey: "UR");

        Assert.True(PdfInspector.Preflight(original).CanAppendFormFieldRevision);
        Assert.True(PdfSyntax.ReadDocumentSecurityInfo(original).UsageRightsAllowFormFillIn);
        PdfOperationResult<PdfDocument> filled = PdfDocument.Load(original).Forms.FillResult(AfterValues);
        Assert.True(filled.Succeeded, string.Join(" ", filled.Diagnostics));
        byte[] updated = filled.Value!.ToBytes();
        Assert.True(updated.AsSpan(0, original.Length).SequenceEqual(original));
        Assert.Equal("after", Assert.Single(PdfInspector.Inspect(updated).FormFields).Value);
    }

    [Theory]
    [InlineData("fw9.pdf")]
    [InlineData("fw9sp.pdf")]
    [InlineData("fw8ben.pdf")]
    [InlineData("sf1034.pdf")]
    public void RightsEnabledHybridGovernmentForms_FillByAppendKeepingScriptsXfaAndSignedBytes(string name) {
        byte[] original = File.ReadAllBytes(GovernmentForm(name));
        PdfDocument document = PdfDocument.Load(original);
        PdfFormField target = document.Inspect().FormFields.First(static field => field.IsTextField && !field.IsReadOnly && !string.IsNullOrEmpty(field.Name));
        var values = new Dictionary<string, string>(StringComparer.Ordinal) { [target.Name!] = "OfficeIMO filled" };

        PdfDocumentPreflight preflight = document.Preflight();
        Assert.True(preflight.Probe.Security.UsageRightsAllowFormFillIn);
        Assert.True(preflight.CanAppendFormFieldRevision);
        Assert.True(preflight.CanAppendOnlyMutate);
        PdfMutationPlan plan = document.PlanMutation(PdfMutationOperation.FillFormFields, values.Keys);
        Assert.Equal(PdfMutationExecutionMode.AppendOnly, plan.ExecutionMode);
        Assert.Contains("AppendOnly.UsageRightsFormFillIn", plan.Warnings);
        Assert.Contains("Input.ActiveContentPreserved", plan.Warnings);
        Assert.Contains("Input.XfaDataNotSynchronized", plan.Warnings);

        PdfOperationResult<PdfDocument> filled = document.Forms.FillResult(values);

        Assert.True(filled.Succeeded, string.Join(" ", filled.Diagnostics));
        byte[] updated = filled.Value!.ToBytes();
        Assert.True(updated.AsSpan(0, original.Length).SequenceEqual(original));
        PdfDocumentInfo reopened = PdfInspector.Inspect(updated);
        Assert.Equal("OfficeIMO filled", reopened.FormFieldsByName[target.Name!].Value);
        Assert.NotNull(reopened.AcroFormXfa);
        Assert.Equal(
            Assert.Single(preflight.Probe.Security.Signatures).ByteRangeValues,
            Assert.Single(reopened.Security.Signatures).ByteRangeValues);
    }

    [Fact]
    public void RemovingXfa_DropsItInTheAppendedRevisionOnly() {
        byte[] original = File.ReadAllBytes(GovernmentForm("fw9.pdf"));
        PdfDocument document = PdfDocument.Load(original);
        string name = document.Inspect().FormFields.First(static field => field.IsTextField && !string.IsNullOrEmpty(field.Name)).Name!;

        byte[] updated = document.Forms.AppendRevision(
            new Dictionary<string, PdfFormFieldValue>(StringComparer.Ordinal) { [name] = PdfFormFieldValue.From("Visible in Acrobat") },
            new PdfIncrementalFormFieldUpdateOptions { GenerateAppearanceStreams = true, KeepNeedAppearances = false, XfaHandling = PdfFormXfaHandling.Remove })
            .ToBytes();

        Assert.True(updated.AsSpan(0, original.Length).SequenceEqual(original));
        PdfDocumentInfo reopened = PdfInspector.Inspect(updated);
        Assert.Null(reopened.AcroFormXfa);
        Assert.Equal("Visible in Acrobat", reopened.FormFieldsByName[name].Value);
        Assert.True(reopened.Security.HasUsageRights);
    }

    [Theory]
    [InlineData(2)]
    [InlineData(3)]
    public void CertifiedHybridForms_PreserveXfaAndRefuseItsRemoval(int permissionLevel) {
        byte[] original = PdfITextInspiredCoverageTests.BuildDocMdpFormPdf(permissionLevel, includeXfa: true);
        PdfDocument document = PdfDocument.Load(original);
        Assert.NotNull(document.Inspect().AcroFormXfa);
        var values = new Dictionary<string, PdfFormFieldValue> { ["Name"] = PdfFormFieldValue.From("Grace") };

        byte[] preserved = document.Forms.AppendRevision(values).ToBytes();
        Assert.True(preserved.AsSpan(0, original.Length).SequenceEqual(original));
        Assert.NotNull(PdfInspector.Inspect(preserved).AcroFormXfa);
        Assert.Equal("Grace", PdfInspector.Inspect(preserved).FormFieldsByName["Name"].Value);

        Assert.Throws<NotSupportedException>(() => document.Forms.AppendRevision(values,
            new PdfIncrementalFormFieldUpdateOptions { XfaHandling = PdfFormXfaHandling.Remove }));
    }

    [Theory]
    [InlineData("UR3", "/9.9", false)]
    [InlineData("UR", "/9.9", false)]
    [InlineData("UR3", "2.2", false)]
    [InlineData("UR3", "(2.2)", false)]
    [InlineData("UR3", null, true)]
    [InlineData("UR", null, true)]
    public void UsageRightsTransformVersion_EnablesOnlyTheSupportedOrDefaultVersion(string key, string? version, bool canFill) {
        byte[] original = BuildUsageRightsFormPdf("[/FillIn]", documentSignature: false, usageRightsKey: key, transformVersion: version);
        PdfDocumentPreflight preflight = PdfInspector.Preflight(original);

        Assert.Equal(canFill, preflight.Probe.Security.UsageRightsAllowFormFillIn);
        Assert.Equal(canFill, PdfSyntax.ReadDocumentSecurityInfo(original).UsageRightsAllowFormFillIn);
        Assert.Equal(canFill, preflight.CanAppendFormFieldRevision);
        Assert.Equal(canFill, PdfDocument.Load(original).Forms.FillResult(AfterValues).Succeeded);
        if (!canFill) {
            Assert.Empty(preflight.Probe.Security.UsageRightsFormPermissions);
            Assert.Equal(PdfMutationExecutionMode.Blocked,
                PdfMutationPlanner.Plan(original, PdfMutationOperation.FillFormFields, fieldNames: AfterValues.Keys).ExecutionMode);
            Assert.Throws<PdfMutationBlockedException>(() => PdfDocument.Load(original).Forms.AppendRevision(AfterValues));
        }
    }

    private static string GovernmentForm(string name) =>
        Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.TestAssets", "OcrQuality", name);

    private static byte[] BuildUsageRightsFormPdf(string? formRights, bool documentSignature,
        string? legacyFormRights = null, string usageRightsKey = "UR3", string? transformVersion = "/2.2") {
        string transformParams = "<< /Type /TransformParams" + (transformVersion is null ? string.Empty : " /V " + transformVersion) +
            (formRights is null ? string.Empty : " /Form " + formRights) +
            " /Document [/FullSave] >>";
        string fields = documentSignature ? "[4 0 R 8 0 R]" : "[4 0 R]";
        string annots = documentSignature ? "[4 0 R 8 0 R]" : "[4 0 R]";
        var objects = new List<string> {
            "<< /Type /Catalog /Pages 2 0 R /AcroForm 5 0 R /Perms << /" + usageRightsKey + " 6 0 R" +
                (legacyFormRights is null ? string.Empty : " /UR 8 0 R") + " >> >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Resources << >> /Annots " + annots + " >>",
            "<< /FT /Tx /T (Name) /V (before) /Type /Annot /Subtype /Widget /Rect [20 200 220 222] /P 3 0 R /DA (/Helv 10 Tf 0 g) >>",
            "<< /Fields " + fields + " /DA (/Helv 10 Tf 0 g) /DR << /Font << /Helv 7 0 R >> >> /SigFlags 2 >>",
            "<< /Type /Sig /Filter /Adobe.PPKLite /SubFilter /adbe.pkcs7.detached /ByteRange [0 100 200 300] /Contents <0011> " +
                "/Reference [<< /Type /SigRef /TransformMethod /UR3 /TransformParams " + transformParams + " >>] >>",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>"
        };
        if (documentSignature) {
            objects.Add("<< /FT /Sig /T (Approval) /V 9 0 R /Type /Annot /Subtype /Widget /Rect [20 100 220 140] /P 3 0 R >>");
            objects.Add("<< /Type /Sig /Filter /Adobe.PPKLite /SubFilter /adbe.pkcs7.detached /Name (Alice) /ByteRange [0 10 20 30] /Contents <2233> >>");
        }
        if (legacyFormRights is not null) {
            objects.Add("<< /Type /Sig /Filter /Adobe.PPKLite /SubFilter /adbe.pkcs7.detached /ByteRange [0 100 200 300] /Contents <4455> " +
                "/Reference [<< /Type /SigRef /TransformMethod /UR /TransformParams << /Form " + legacyFormRights + " >> >>] >>");
        }

        var builder = new StringBuilder("%PDF-1.7\n");
        var offsets = new List<int>();
        for (int index = 0; index < objects.Count; index++) {
            offsets.Add(builder.Length);
            builder.Append(index + 1).Append(" 0 obj\n").Append(objects[index]).Append("\nendobj\n");
        }

        int xref = builder.Length;
        builder.Append("xref\n0 ").Append(objects.Count + 1).Append("\n0000000000 65535 f \n");
        foreach (int offset in offsets) builder.Append(offset.ToString("D10", CultureInfo.InvariantCulture)).Append(" 00000 n \n");
        builder.Append("trailer\n<< /Size ").Append(objects.Count + 1).Append(" /Root 1 0 R >>\nstartxref\n").Append(xref).Append("\n%%EOF\n");
        return Encoding.ASCII.GetBytes(builder.ToString());
    }
}
