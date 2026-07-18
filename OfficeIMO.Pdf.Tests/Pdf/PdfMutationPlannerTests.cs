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
    public void Plan_ChoosesEncryptedAppendWithOwnerAuthorization() {
        byte[] source = PdfDocument.Create(new PdfOptions().SetEncryption("open", "owner"))
            .Paragraph(paragraph => paragraph.Text("Encrypted owner planner source"))
            .ToBytes();
        var ownerOptions = new PdfReadOptions { Password = "owner" };

        PdfMutationPlan plan = PdfMutationPlanner.Plan(source, PdfMutationOperation.UpdateMetadata, ownerOptions);
        PdfOperationResult<PdfDocument> result = PdfDocument.Open(source)
            .TryUpdateMetadata(title: "Owner-planned update", options: ownerOptions);

        Assert.True(plan.CanExecute);
        Assert.Equal(PdfMutationExecutionMode.AppendOnly, plan.ExecutionMode);
        Assert.True(plan.AppendOnlyAvailable);
        Assert.True(result.Succeeded, string.Join(" ", result.Diagnostics));
        Assert.Equal("Owner-planned update", result.Value!.Read.Metadata(ownerOptions).Title);
    }

    [Fact]
    public void TryFill_AppendsEncryptedFormRevisionWithOwnerAuthorization() {
        byte[] source = PdfDocument.Create(new PdfOptions().SetEncryption("open", "owner"))
            .TextField("Name", width: 180, height: 24, value: "Ada")
            .ToBytes();
        var ownerOptions = new PdfReadOptions { Password = "owner" };
        var values = new Dictionary<string, string> { ["Name"] = "Grace" };

        PdfOperationResult<PdfDocument> result = PdfDocument.Open(source, ownerOptions).Forms.TryFill(values);

        Assert.True(result.Succeeded, string.Join(" ", result.Diagnostics));
        Assert.Equal(PdfMutationExecutionMode.AppendOnly, result.MutationPlan!.ExecutionMode);
        byte[] updated = result.RequireValue().ToBytes();
        Assert.True(updated.AsSpan(0, source.Length).SequenceEqual(source));
        PdfDocumentInfo info = PdfInspector.Inspect(updated, ownerOptions);
        Assert.True(info.Security.HasEncryption);
        Assert.Equal("Grace", Assert.Single(info.FormFields, static field => field.Name == "Name").Value);
    }

    [Fact]
    public void TryMergeWith_UsesMergePolicyForFormBearingPrimary() {
        byte[] primary = PdfDocument.Create().TextField("Name", value: "Ada").ToBytes();
        byte[] incoming = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Incoming")).ToBytes();

        PdfOperationResult<PdfDocument> result = PdfDocument.Open(primary).TryMergeWith(incoming);

        Assert.True(result.Succeeded, string.Join(" ", result.Diagnostics));
        Assert.Equal(PdfMutationOperation.MergeDocuments, result.MutationPlan!.Operation);
        PdfDocumentInfo info = result.RequireValue().Inspect();
        Assert.Equal(2, info.PageCount);
        Assert.Single(info.FormFields, static field => field.Name == "Name");
    }

    [Fact]
    public void TryPageImports_UseMergePolicyForFormBearingTarget() {
        byte[] target = PdfDocument.Create().TextField("Name", value: "Ada").ToBytes();
        byte[] incoming = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Incoming")).ToBytes();

        PdfOperationResult<PdfDocument> appended = PdfDocument.Open(target).Pages.TryAppend(incoming);
        PdfOperationResult<PdfDocument> prepended = PdfDocument.Open(target).Pages.TryPrepend(incoming);
        PdfOperationResult<PdfDocument> inserted = PdfDocument.Open(target).Pages.TryInsert(1, incoming);

        Assert.All(new[] { appended, prepended, inserted }, result => {
            Assert.True(result.Succeeded, string.Join(" ", result.Diagnostics));
            Assert.Equal(PdfMutationOperation.MergeDocuments, result.MutationPlan!.Operation);
            Assert.Equal(2, result.RequireValue().Inspect().PageCount);
            Assert.Single(result.RequireValue().Inspect().FormFields, static field => field.Name == "Name");
        });
    }

    [Fact]
    public void Plan_StreamStopsBufferingAtConfiguredInputLimit() {
        byte[] source = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Bounded planner stream")).ToBytes();
        byte[] padded = new byte[1024 * 1024];
        source.CopyTo(padded, 0);
        using var stream = new ChunkedNonSeekableStream(padded, 256);
        var options = new PdfReadOptions { Limits = new PdfReadLimits { MaxInputBytes = 1024 } };

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            PdfMutationPlanner.Plan(stream, PdfMutationOperation.UpdateMetadata, options));

        Assert.Equal(PdfReadLimitKind.InputBytes, exception.Kind);
        Assert.InRange(stream.BytesRead, 1025, 1280);
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
    public void Plan_BlocksEncryptedExternalSignaturePreparationBeforeRawObjectAppend() {
        byte[] source = PdfDocument.Create(new PdfOptions().SetEncryption("open", "owner"))
            .Paragraph(paragraph => paragraph.Text("Encrypted signature source"))
            .ToBytes();
        var readOptions = new PdfReadOptions { Password = "owner" };

        PdfAppendOnlyMutationReport appendOnly = PdfIncrementalUpdater.AnalyzeAppendOnlyMutation(source, readOptions);
        PdfMutationPlan plan = PdfMutationPlanner.Plan(source, PdfMutationOperation.PrepareExternalSignature, readOptions);
        PdfOperationResult<PdfExternalSignaturePreparation> result = PdfDocument.Open(source, readOptions)
            .TryPrepareExternalSignature(options: readOptions);

        Assert.False(appendOnly.CanPrepareExternalSignature);
        Assert.False(plan.CanExecute);
        Assert.Equal(PdfMutationExecutionMode.Blocked, plan.ExecutionMode);
        Assert.Contains("AppendOnly.EncryptedRawSignatureObject", plan.BlockerCodes);
        Assert.False(result.CanAttempt);
        Assert.False(result.Succeeded);
    }

    [Fact]
    public void ExternalSignatureFinalizationUsesReservedPatchContract() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Signature finalization planner source"))
            .ToBytes();
        PdfExternalSignaturePreparation preparation = PdfIncrementalUpdater.PrepareExternalSignature(
            source,
            new PdfExternalSignatureOptions { ReservedSignatureContentsBytes = 256 });

        PdfMutationPlan plan = PdfMutationPlanner.Plan(
            preparation.PreparedPdf,
            PdfMutationOperation.FinalizeExternalSignature);

        Assert.True(plan.CanExecute);
        Assert.Equal(PdfMutationExecutionMode.AppendOnly, plan.ExecutionMode);
        Assert.Equal(new[] { PdfMutationStructure.Signatures }, plan.AffectedStructures);
        Assert.Contains(PdfMutationPermissionCheck.FillSignatureContentsReservation, plan.PermissionChecks);
        Assert.DoesNotContain(PdfMutationPermissionCheck.AppendRevision, plan.PermissionChecks);
        Assert.Contains(PdfMutationProof.ReservedSignatureContentsPatch, plan.RequiredProofs);
        Assert.Contains(PdfMutationProof.SignatureByteRanges, plan.RequiredProofs);
        Assert.DoesNotContain(PdfMutationProof.BytePrefixPreservation, plan.RequiredProofs);
        Assert.DoesNotContain(PdfMutationProof.RevisionChain, plan.RequiredProofs);

        byte[] signed = PdfIncrementalUpdater.ApplyExternalSignature(preparation, new byte[] { 0x30, 0x01, 0x00 });
        PdfMutationPlan completedPlan = PdfMutationPlanner.Plan(signed, PdfMutationOperation.FinalizeExternalSignature);
        Assert.False(completedPlan.CanExecute);
        Assert.Contains("AppendOnly.ActionBlocked.SignatureFinalize", completedPlan.BlockerCodes);
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

    [Fact]
    public void EncryptedPageExtractionRequiresOwnerOrCopyAndAssemblyPermissions() {
        var encryption = new PdfStandardEncryptionOptions("open") {
            OwnerPassword = "owner",
            AllowedPermissions = PdfStandardPermissions.None
        };
        byte[] source = PdfDocument.Create(new PdfOptions().SetEncryption(encryption))
            .Paragraph(paragraph => paragraph.Text("Restricted extraction"))
            .ToBytes();
        var userOptions = new PdfReadOptions { Password = "open" };
        var ownerOptions = new PdfReadOptions { Password = "owner" };

        PdfMutationPlan userPlan = PdfMutationPlanner.Plan(source, PdfMutationOperation.ExtractPages, userOptions);
        PdfMutationPlan ownerPlan = PdfMutationPlanner.Plan(source, PdfMutationOperation.ExtractPages, ownerOptions);
        PdfOperationResult<PdfDocument> userResult = PdfDocument.Open(source, userOptions)
            .Pages.TryExtract(PdfPageSelection.Parse("1"));
        PdfOperationResult<PdfDocument> ownerResult = PdfDocument.Open(source, ownerOptions)
            .Pages.TryExtract(PdfPageSelection.Parse("1"));

        Assert.False(userPlan.CanExecute);
        Assert.Contains("FullRewrite.Encryption", userPlan.BlockerCodes);
        Assert.False(userResult.CanAttempt);
        Assert.True(ownerPlan.CanExecute);
        Assert.Equal(PdfMutationExecutionMode.FullRewrite, ownerPlan.ExecutionMode);
        Assert.True(ownerResult.Succeeded, string.Join(" ", ownerResult.Diagnostics));
        Assert.False(PdfInspector.Probe(ownerResult.RequireValue().ToBytes()).HasEncryption);
    }

    private sealed class ChunkedNonSeekableStream : Stream {
        private readonly byte[] _data;
        private readonly int _maximumChunkSize;
        private int _position;
        internal ChunkedNonSeekableStream(byte[] data, int maximumChunkSize) { _data = data; _maximumChunkSize = maximumChunkSize; }
        internal int BytesRead => _position;
        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override int Read(byte[] buffer, int offset, int count) {
            if (_position >= _data.Length) return 0;
            int read = Math.Min(Math.Min(count, _maximumChunkSize), _data.Length - _position);
            Buffer.BlockCopy(_data, _position, buffer, offset, read);
            _position += read;
            return read;
        }
        public override void Flush() { }
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }
}
