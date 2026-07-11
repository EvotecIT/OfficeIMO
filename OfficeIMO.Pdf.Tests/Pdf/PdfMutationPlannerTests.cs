using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfMutationPlannerTests {
    [Fact]
    public void Plan_ChoosesFullRewriteForOrdinaryMetadataMutation() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Planner metadata source"))
            .ToBytes();

        PdfMutationPlan plan = PdfMutationPlanner.Plan(source, PdfMutationOperation.UpdateMetadata);

        Assert.True(plan.CanExecute);
        Assert.Equal(PdfMutationExecutionMode.FullRewrite, plan.ExecutionMode);
        Assert.True(plan.FullRewriteAvailable);
        Assert.True(plan.AppendOnlyAvailable);
        Assert.Contains(PdfMutationStructure.InfoDictionary, plan.AffectedStructures);
        Assert.Contains(PdfMutationPermissionCheck.ModifyDocument, plan.PermissionChecks);
        Assert.Contains(PdfMutationProof.RewritePreservation, plan.RequiredProofs);
        Assert.Contains(PdfMutationProof.MetadataReadback, plan.RequiredProofs);
        Assert.Empty(plan.BlockerCodes);
    }

    [Fact]
    public void TryUpdateMetadata_ChoosesAppendOnlyAndPreservesTaggedSourceBytes() {
        byte[] source = PdfRewritePreservationTestSupport.BuildTaggedPreservationProofPdf();
        int originalRevisionCount = PdfInspector.Inspect(source).Security.RevisionCount;

        PdfOperationResult<PdfDocument> result = PdfDocument.Open(source)
            .TryUpdateMetadata(title: "Planner tagged update");

        Assert.True(result.Succeeded, string.Join(" ", result.Diagnostics));
        PdfMutationPlan plan = Assert.IsType<PdfMutationPlan>(result.MutationPlan);
        Assert.Equal(PdfMutationExecutionMode.AppendOnly, plan.ExecutionMode);
        Assert.False(plan.FullRewriteAvailable);
        Assert.True(plan.AppendOnlyAvailable);
        Assert.Contains(PdfMutationProof.BytePrefixPreservation, plan.RequiredProofs);
        Assert.Contains(PdfMutationProof.RevisionChain, plan.RequiredProofs);

        byte[] updated = result.RequireValue().ToBytes();
        Assert.True(updated.AsSpan(0, source.Length).SequenceEqual(source));
        PdfDocumentInfo updatedInfo = PdfInspector.Inspect(updated);
        Assert.Equal("Planner tagged update", updatedInfo.Metadata.Title);
        Assert.True(updatedInfo.Security.RevisionCount > originalRevisionCount);
        Assert.True(updatedInfo.HasReadableTaggedContent);
        Assert.Contains("Document", updatedInfo.TaggedContent!.StructureTypes);
    }

    [Fact]
    public void Plan_BlocksPageTreeMutationForSignedIncrementalInput() {
        byte[] source = PdfRewritePreservationTestSupport.BuildSignedIncrementalProofPdf();

        PdfMutationPlan plan = PdfMutationPlanner.Plan(source, PdfMutationOperation.ModifyPageTree);

        Assert.False(plan.CanExecute);
        Assert.Equal(PdfMutationExecutionMode.Blocked, plan.ExecutionMode);
        Assert.Contains(plan.BlockerCodes, code => code.StartsWith("FullRewrite.", StringComparison.Ordinal));
        Assert.Contains("FullRewrite.AppendOnlyRequired", plan.BlockerCodes);
        Assert.Contains("AppendOnly.NotImplemented.ModifyPageTree", plan.BlockerCodes);
        Assert.Empty(plan.RequiredProofs);
    }

    [Fact]
    public void RequireFullRewrite_ThrowsTypedExceptionWithCompleteBlockedPlan() {
        byte[] source = PdfRewritePreservationTestSupport.BuildSignedIncrementalProofPdf();

        PdfMutationBlockedException exception = Assert.Throws<PdfMutationBlockedException>(() =>
            PdfMutationPlanner.RequireFullRewrite(source, PdfMutationOperation.ModifyPageTree));

        Assert.StartsWith(exception.Plan.Summary, exception.Message, StringComparison.Ordinal);
        Assert.Equal(PdfMutationOperation.ModifyPageTree, exception.Plan.Operation);
        Assert.Equal(PdfMutationExecutionPreference.RequireFullRewrite, exception.Plan.ExecutionPreference);
        Assert.Equal(PdfMutationExecutionMode.Blocked, exception.Plan.ExecutionMode);
        Assert.Contains(PdfMutationStructure.PageTree, exception.Plan.AffectedStructures);
        Assert.Contains(PdfMutationStructure.Catalog, exception.Plan.AffectedStructures);
        Assert.Contains(PdfMutationPermissionCheck.AssembleDocument, exception.Plan.PermissionChecks);
        Assert.Contains("FullRewrite.AppendOnlyRequired", exception.Plan.BlockerCodes);
    }

    [Fact]
    public void StaticPageEditor_ExposesPlannerDecisionWhenFullRewriteIsBlocked() {
        byte[] source = PdfRewritePreservationTestSupport.BuildSignedIncrementalProofPdf();

        PdfMutationBlockedException exception = Assert.Throws<PdfMutationBlockedException>(() =>
            PdfPageEditor.DeletePages(source, 1));

        Assert.Equal(PdfMutationOperation.ModifyPageTree, exception.Plan.Operation);
        Assert.Equal(PdfMutationExecutionMode.Blocked, exception.Plan.ExecutionMode);
        PdfMutationCapabilityRecord pageTree = Assert.Single(
            exception.Plan.CapabilityRecords,
            record => record.Kind == PdfMutationCapabilityKind.PageTreeChanges);
        Assert.True(pageTree.FullRewriteImplemented);
        Assert.False(pageTree.FullRewriteAllowed);
        Assert.Contains("FullRewrite.AppendOnlyRequired", pageTree.BlockerCodes);
    }

    [Fact]
    public void StaticMetadataEditor_RequiresFullRewriteEvenWhenAppendOnlyIsAvailable() {
        byte[] source = PdfRewritePreservationTestSupport.BuildTaggedPreservationProofPdf();

        PdfMutationBlockedException exception = Assert.Throws<PdfMutationBlockedException>(() =>
            PdfMetadataEditor.UpdateMetadata(source, title: "Blocked full rewrite"));

        Assert.Equal(PdfMutationOperation.UpdateMetadata, exception.Plan.Operation);
        Assert.Equal(PdfMutationExecutionPreference.RequireFullRewrite, exception.Plan.ExecutionPreference);
        Assert.True(exception.Plan.AppendOnlyAvailable);
        Assert.False(exception.Plan.FullRewriteAvailable);
        Assert.Contains(PdfMutationStructure.InfoDictionary, exception.Plan.AffectedStructures);
        Assert.Contains("FullRewrite.TaggedContent", exception.Plan.BlockerCodes);
    }

    [Fact]
    public void Plan_BlocksEncryptedMutationEvenWithValidPassword() {
        byte[] source = PdfDocument.Create(new PdfOptions().SetEncryption("open", "owner"))
            .Paragraph(paragraph => paragraph.Text("Encrypted planner source"))
            .ToBytes();
        var readOptions = new PdfReadOptions { Password = "open" };

        PdfMutationPlan plan = PdfMutationPlanner.Plan(source, PdfMutationOperation.UpdateMetadata, readOptions);

        Assert.False(plan.CanExecute);
        Assert.True(plan.Preflight.CanRead);
        Assert.Equal(PdfMutationExecutionMode.Blocked, plan.ExecutionMode);
        Assert.Contains("FullRewrite.Encryption", plan.BlockerCodes);
        Assert.Contains("AppendOnly.Encrypted", plan.BlockerCodes);
    }

    [Fact]
    public void Plan_BlocksPageTreeMutationWhenSourceStructureCannotBePreserved() {
        byte[] source = PdfRewritePreservationTestSupport.BuildSourceStructurePreservationProofPdf();

        PdfMutationPlan plan = PdfMutationPlanner.Plan(source, PdfMutationOperation.ModifyPageTree);

        Assert.False(plan.CanExecute);
        Assert.Equal(PdfMutationExecutionMode.Blocked, plan.ExecutionMode);
        Assert.Contains("FullRewrite.XrefStreamPreservation", plan.BlockerCodes);
        Assert.Contains("FullRewrite.ObjectStreamPreservation", plan.BlockerCodes);
    }

    [Fact]
    public void TryFill_ChoosesAppendOnlyForPermittedDocMdpFieldAndPreservesPrefix() {
        byte[] source = PdfITextInspiredCoverageTests.BuildDocMdpFormPdf(permissionLevel: 2);
        var values = new Dictionary<string, string> { ["Name"] = "Grace" };

        PdfOperationResult<PdfDocument> result = PdfDocument.Open(source).Forms.TryFill(values);

        Assert.True(result.Succeeded, string.Join(" ", result.Diagnostics));
        PdfMutationPlan plan = Assert.IsType<PdfMutationPlan>(result.MutationPlan);
        Assert.Equal(PdfMutationExecutionMode.AppendOnly, plan.ExecutionMode);
        Assert.Contains(PdfMutationPermissionCheck.DocMdp, plan.PermissionChecks);
        Assert.Contains(PdfMutationPermissionCheck.FieldMdp, plan.PermissionChecks);
        Assert.Contains(PdfMutationProof.SignatureByteRanges, plan.RequiredProofs);

        byte[] updated = result.RequireValue().ToBytes();
        Assert.True(updated.AsSpan(0, source.Length).SequenceEqual(source));
        PdfFormField field = Assert.Single(PdfInspector.Inspect(updated).FormFields, field => field.Name == "Name");
        Assert.Equal("Grace", field.Value);
    }

    [Fact]
    public void Plan_UsesRequestedFieldNamesWhenEvaluatingFieldMdpLocks() {
        byte[] source = PdfITextInspiredCoverageTests.BuildDocMdpFormPdf(
            permissionLevel: 2,
            lockDictionary: "<< /Type /SigFieldLock /Action /Include /Fields [(Name)] >>");

        PdfMutationPlan plan = PdfMutationPlanner.Plan(
            source,
            PdfMutationOperation.FillFormFields,
            fieldNames: new[] { "Name" });

        Assert.False(plan.CanExecute);
        Assert.Equal(PdfMutationExecutionMode.Blocked, plan.ExecutionMode);
        Assert.Contains("AppendOnly.SignatureFieldLock", plan.BlockerCodes);
        Assert.Contains(plan.Diagnostics, diagnostic => diagnostic.Contains("SignatureFieldLock", StringComparison.Ordinal));
    }

    [Fact]
    public void Plan_ChoosesAppendOnlyForExternalSignaturePreparation() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Signature planner source"))
            .ToBytes();

        PdfMutationPlan plan = PdfMutationPlanner.Plan(source, PdfMutationOperation.PrepareExternalSignature);

        Assert.True(plan.CanExecute);
        Assert.Equal(PdfMutationExecutionMode.AppendOnly, plan.ExecutionMode);
        Assert.False(plan.FullRewriteAvailable);
        Assert.True(plan.AppendOnlyAvailable);
        Assert.Contains(PdfMutationStructure.Signatures, plan.AffectedStructures);
        Assert.Contains(PdfMutationProof.BytePrefixPreservation, plan.RequiredProofs);
        Assert.Contains(PdfMutationProof.SignatureByteRanges, plan.RequiredProofs);
    }

    [Fact]
    public void Plan_ExposesSharedCapabilityRecordsForAffectedStructures() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Capability record source"))
            .ToBytes();

        PdfMutationPlan pagePlan = PdfMutationPlanner.Plan(source, PdfMutationOperation.ModifyPageTree);
        PdfMutationCapabilityRecord pageTree = Assert.Single(
            pagePlan.CapabilityRecords,
            record => record.Kind == PdfMutationCapabilityKind.PageTreeChanges);
        PdfMutationCapabilityRecord catalog = Assert.Single(
            pagePlan.CapabilityRecords,
            record => record.Kind == PdfMutationCapabilityKind.CatalogChanges);

        Assert.Contains(PdfMutationStructure.PageTree, pageTree.AffectedStructures);
        Assert.True(pageTree.FullRewriteImplemented);
        Assert.True(pageTree.FullRewriteAllowed);
        Assert.False(pageTree.AppendOnlyImplemented);
        Assert.Contains(PdfMutationStructure.Catalog, catalog.AffectedStructures);
        Assert.Contains(PdfMutationPermissionCheck.AssembleDocument, pageTree.PermissionChecks);
        Assert.Contains(PdfMutationProof.PageStructureReadback, pageTree.RequiredProofs);
    }

    [Fact]
    public void ExplicitAppendWorkflowRequiresAppendOnlyEvenForOrdinaryInput() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Explicit append source"))
            .ToBytes();

        PdfOperationResult<PdfDocument> result = PdfDocument.Open(source)
            .TryAppendMetadataRevision(title: "Explicit append title");

        Assert.True(result.Succeeded, string.Join(" ", result.Diagnostics));
        PdfMutationPlan plan = Assert.IsType<PdfMutationPlan>(result.MutationPlan);
        Assert.Equal(PdfMutationExecutionPreference.RequireAppendOnly, plan.ExecutionPreference);
        Assert.Equal(PdfMutationExecutionMode.AppendOnly, plan.ExecutionMode);
        Assert.True(result.RequireValue().ToBytes().AsSpan(0, source.Length).SequenceEqual(source));
    }

    [Fact]
    public void PageEditingWorkflowUsesPlannerAndBlocksSignedIncrementalInput() {
        byte[] ordinary = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("First"))
            .PageBreak()
            .Paragraph(paragraph => paragraph.Text("Second"))
            .ToBytes();

        PdfOperationResult<PdfDocument> ordinaryResult = PdfDocument.Open(ordinary)
            .Pages.TryDelete(PdfPageSelection.Parse("2"));
        PdfOperationResult<PdfDocument> signedResult = PdfDocument.Open(PdfRewritePreservationTestSupport.BuildSignedIncrementalProofPdf())
            .Pages.TryDelete(PdfPageSelection.Parse("1"));

        Assert.True(ordinaryResult.Succeeded, string.Join(" ", ordinaryResult.Diagnostics));
        Assert.Equal(PdfMutationExecutionMode.FullRewrite, ordinaryResult.MutationPlan!.ExecutionMode);
        Assert.Equal(1, ordinaryResult.RequireValue().Inspect().PageCount);
        Assert.False(signedResult.CanAttempt);
        Assert.False(signedResult.Succeeded);
        Assert.Equal(PdfMutationExecutionMode.Blocked, signedResult.MutationPlan!.ExecutionMode);
        Assert.Contains("FullRewrite.AppendOnlyRequired", signedResult.MutationPlan.BlockerCodes);
    }

    [Fact]
    public void EncryptedPageExtractionUsesExplicitPlannerExceptionAndReturnsUnencryptedOutput() {
        byte[] source = PdfDocument.Create(new PdfOptions().SetEncryption("open", "owner"))
            .Paragraph(paragraph => paragraph.Text("Encrypted extraction"))
            .ToBytes();
        var readOptions = new PdfReadOptions { Password = "open" };

        PdfOperationResult<PdfDocument> result = PdfDocument.Open(source, readOptions)
            .Pages.TryExtract(PdfPageSelection.Parse("1"));

        Assert.True(result.Succeeded, string.Join(" ", result.Diagnostics));
        PdfMutationPlan plan = Assert.IsType<PdfMutationPlan>(result.MutationPlan);
        Assert.Equal(PdfMutationOperation.ExtractPages, plan.Operation);
        Assert.Equal(PdfMutationExecutionMode.FullRewrite, plan.ExecutionMode);
        Assert.False(PdfInspector.Probe(result.RequireValue().ToBytes()).HasEncryption);
    }
}
