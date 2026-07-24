using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfMutationPlannerTests {
    [Fact]
    public void PermissionCheckEnumPreservesPublishedNumericValues() {
        Assert.Equal(0, (int)PdfMutationPermissionCheck.ReadDocument);
        Assert.Equal(1, (int)PdfMutationPermissionCheck.ModifyDocument);
        Assert.Equal(2, (int)PdfMutationPermissionCheck.AssembleDocument);
        Assert.Equal(3, (int)PdfMutationPermissionCheck.ModifyAnnotations);
        Assert.Equal(4, (int)PdfMutationPermissionCheck.FillForms);
        Assert.Equal(5, (int)PdfMutationPermissionCheck.DocMdp);
        Assert.Equal(6, (int)PdfMutationPermissionCheck.FieldMdp);
        Assert.Equal(7, (int)PdfMutationPermissionCheck.AppendRevision);
        Assert.Equal(8, (int)PdfMutationPermissionCheck.FillSignatureContentsReservation);
        Assert.Equal(9, (int)PdfMutationPermissionCheck.OwnerAuthorization);
        Assert.Equal(10, (int)PdfMutationPermissionCheck.CopyContents);
    }

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
    public void ActiveContentBlocksAppendOnlyMetadataAndFormMutations() {
        byte[] metadataSource = WithCatalogAction(
            PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Active metadata source")).ToBytes());
        byte[] formSource = WithCatalogAction(
            PdfDocument.Create().TextField("Account", value: "before").ToBytes());

        PdfMutationPlan metadataPlan = PdfMutationPlanner.Plan(
            metadataSource,
            PdfMutationOperation.UpdateMetadata);
        PdfMutationPlan formPlan = PdfMutationPlanner.Plan(
            formSource,
            PdfMutationOperation.FillFormFields,
            fieldNames: new[] { "Account" });
        PdfOperationResult<PdfDocument> metadataResult = PdfDocument.Open(metadataSource)
            .TryUpdateMetadata(title: "should not be appended");
        PdfOperationResult<PdfDocument> formResult = PdfDocument.Open(formSource).Forms.TryFill(
            new Dictionary<string, string> { ["Account"] = "after" });

        Assert.False(metadataPlan.CanExecute);
        Assert.False(metadataPlan.AppendOnlyAvailable);
        Assert.False(formPlan.CanExecute);
        Assert.False(formPlan.AppendOnlyAvailable);
        Assert.False(metadataResult.Succeeded);
        Assert.False(formResult.Succeeded);
    }

    [Fact]
    public void GoToEActionBlocksAppendOnlyMetadataMutation() {
        byte[] source = WithCatalogAction(
            PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Embedded target action")).ToBytes(),
            "GoToE");

        PdfMutationPlan plan = PdfMutationPlanner.Plan(source, PdfMutationOperation.UpdateMetadata);
        PdfOperationResult<PdfDocument> result = PdfDocument.Open(source)
            .TryUpdateMetadata(title: "must not preserve GoToE");

        Assert.True(PdfInspector.Probe(source).HasActiveContent);
        Assert.False(plan.CanExecute);
        Assert.False(plan.AppendOnlyAvailable);
        Assert.False(result.Succeeded);
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
    public void Plan_ChoosesFullRewriteForAuthorizedEncryptedMutation() {
        byte[] source = PdfDocument.Create(new PdfOptions().SetEncryption("open", "owner"))
            .Paragraph(paragraph => paragraph.Text("Encrypted planner source"))
            .ToBytes();
        var readOptions = new PdfReadOptions { Password = "open" };

        PdfMutationPlan plan = PdfMutationPlanner.Plan(source, PdfMutationOperation.UpdateMetadata, readOptions);

        Assert.True(plan.CanExecute);
        Assert.True(plan.Preflight.CanRead);
        Assert.Equal(PdfMutationExecutionMode.FullRewrite, plan.ExecutionMode);
        Assert.True(plan.FullRewriteAvailable);
        Assert.False(plan.AppendOnlyAvailable);
        Assert.Empty(plan.BlockerCodes);
        Assert.Contains("Output.EncryptionWillBeRemoved", plan.Warnings);
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
    public void TryPageImports_PropagateExplicitTargetAuthenticationThroughExecution() {
        var encryption = new PdfStandardEncryptionOptions("open") {
            OwnerPassword = "owner",
            AllowedPermissions = PdfStandardPermissions.None
        };
        byte[] target = PdfDocument.Create(new PdfOptions().SetEncryption(encryption))
            .Paragraph(paragraph => paragraph.Text("Target one"))
            .PageBreak()
            .Paragraph(paragraph => paragraph.Text("Target two"))
            .ToBytes();
        byte[] incoming = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Incoming one"))
            .PageBreak()
            .Paragraph(paragraph => paragraph.Text("Incoming two"))
            .ToBytes();
        var targetReadOptions = new PdfReadOptions {
            Password = "open",
            PermissionPolicy = PdfPermissionPolicy.IgnoreRestrictions
        };

        PdfOperationResult<PdfDocument> appended = PdfDocument.Open(target).Pages.TryAppend(
            incoming,
            options: targetReadOptions);
        PdfOperationResult<PdfDocument> prepended = PdfDocument.Open(target).Pages.TryPrepend(
            incoming,
            PdfPageSelection.From(2),
            options: targetReadOptions);
        PdfOperationResult<PdfDocument> inserted = PdfDocument.Open(target).Pages.TryInsert(
            2,
            incoming,
            PdfPageSelection.From(1),
            options: targetReadOptions);

        Assert.All(new[] { appended, prepended, inserted }, result => {
            Assert.True(result.Succeeded, string.Join(" ", result.Diagnostics));
            Assert.Equal(PdfMutationOperation.MergeDocuments, result.MutationPlan!.Operation);
            Assert.Contains("Output.EncryptionWillBeRemoved", result.MutationPlan.Warnings);
            Assert.False(PdfInspector.Probe(result.RequireValue().ToBytes()).HasEncryption);
        });
        Assert.Equal(4, appended.RequireValue().Inspect().PageCount);
        Assert.Equal(3, prepended.RequireValue().Inspect().PageCount);
        Assert.Equal(3, inserted.RequireValue().Inspect().PageCount);
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
    public void ExternalSignatureFinalizationIgnoresContentsTextInsideMetadataStrings() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Signature contents lexical parsing"))
            .ToBytes();
        PdfExternalSignaturePreparation preparation = PdfIncrementalUpdater.PrepareExternalSignature(
            source,
            new PdfExternalSignatureOptions {
                Reason = "Review /Contents <00000000> marker as ordinary text",
                ReservedSignatureContentsBytes = 256
            });

        PdfMutationPlan plan = PdfMutationPlanner.Plan(
            preparation.PreparedPdf,
            PdfMutationOperation.FinalizeExternalSignature);
        Assert.True(plan.CanExecute);

        byte[] signed = PdfIncrementalUpdater.ApplyExternalSignature(
            preparation,
            new byte[] { 0x30, 0x01, 0x00 });

        Assert.NotEmpty(signed);
    }

    [Fact]
    public void ExternalSignatureFinalizationIgnoresObjectHeadersInsideMetadataStrings() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Signature object-boundary lexical parsing"))
            .ToBytes();
        PdfExternalSignaturePreparation preparation = PdfIncrementalUpdater.PrepareExternalSignature(
            source,
            new PdfExternalSignatureOptions {
                Reason = "Review note\n99 0 obj should remain ordinary text",
                ReservedSignatureContentsBytes = 256
            });

        PdfMutationPlan rawPlan = PdfMutationPlanner.Plan(
            preparation.PreparedPdf,
            PdfMutationOperation.FinalizeExternalSignature);
        PdfMutationPlan documentPlan = PdfDocument.Open(preparation.PreparedPdf)
            .PlanMutation(PdfMutationOperation.FinalizeExternalSignature);
        byte[] signed = PdfIncrementalUpdater.ApplyExternalSignature(
            preparation.PreparedPdf,
            new byte[] { 0x30, 0x01, 0x00 });

        Assert.True(rawPlan.CanExecute);
        Assert.True(documentPlan.CanExecute);
        Assert.NotEmpty(signed);
    }

    [Fact]
    public void DocumentMutationPlanningRetainsBytesForSignatureReservationValidation() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Document signature planning source"))
            .ToBytes();
        PdfExternalSignaturePreparation preparation = PdfIncrementalUpdater.PrepareExternalSignature(
            source,
            new PdfExternalSignatureOptions { ReservedSignatureContentsBytes = 256 });
        PdfDocument document = PdfDocument.Open(preparation.PreparedPdf);

        PdfMutationPlan plan = document.PlanMutation(PdfMutationOperation.FinalizeExternalSignature);
        PdfMutationPortfolioReport portfolio = document.AssessMutations(new[] {
            PdfMutationOperation.FinalizeExternalSignature
        });
        PdfMutationPlan portfolioPlan = portfolio.Get(PdfMutationOperation.FinalizeExternalSignature);

        Assert.True(plan.CanExecute);
        Assert.True(portfolioPlan.CanExecute);
        Assert.Equal(PdfMutationExecutionMode.AppendOnly, plan.ExecutionMode);
        Assert.Equal(PdfMutationExecutionMode.AppendOnly, portfolioPlan.ExecutionMode);
        Assert.DoesNotContain("SignatureReservation.Invalid", plan.BlockerCodes);
        Assert.DoesNotContain("SignatureReservation.Invalid", portfolioPlan.BlockerCodes);
        Assert.Same(portfolio.Preflight, portfolioPlan.Preflight);
    }

    [Fact]
    public void ExternalSignatureFinalizationRejectsByteRangeThatDoesNotMatchReservation() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Signature reservation validation"))
            .ToBytes();
        PdfExternalSignaturePreparation preparation = PdfIncrementalUpdater.PrepareExternalSignature(
            source,
            new PdfExternalSignatureOptions { ReservedSignatureContentsBytes = 256 });
        long originalTailLength = preparation.ByteRangeValues[3];
        byte[] crafted = ReplaceFirstAscii(
            preparation.PreparedPdf,
            originalTailLength.ToString("00000000000000000000", System.Globalization.CultureInfo.InvariantCulture),
            (originalTailLength + 1L).ToString("00000000000000000000", System.Globalization.CultureInfo.InvariantCulture));

        PdfMutationPlan plan = PdfMutationPlanner.Plan(crafted, PdfMutationOperation.FinalizeExternalSignature);
        PdfMutationPlan documentPlan = PdfDocument.Open(crafted).PlanMutation(PdfMutationOperation.FinalizeExternalSignature);

        Assert.False(plan.CanExecute);
        Assert.False(documentPlan.CanExecute);
        Assert.Contains("SignatureReservation.Invalid", plan.BlockerCodes);
        Assert.Contains("SignatureReservation.Invalid", documentPlan.BlockerCodes);
        Assert.Throws<PdfMutationBlockedException>(() =>
            PdfDocument.Open(crafted).CompleteExternalSignature(new byte[] { 0x30, 0x01, 0x00 }));
    }

    [Fact]
    public void ExternalSignatureFinalizationRejectsDuplicateContentsBoundToAnotherObject() {
        byte[] crafted = BuildDuplicateContentsSignatureReservation(256);

        PdfMutationPlan plan = PdfMutationPlanner.Plan(crafted, PdfMutationOperation.FinalizeExternalSignature);

        Assert.False(plan.CanExecute);
        Assert.Contains("SignatureReservation.Invalid", plan.BlockerCodes);
        Assert.Throws<ArgumentException>(() =>
            PdfDocument.Open(crafted).CompleteExternalSignature(new byte[] { 0x30, 0x01, 0x00 }));
    }

    [Fact]
    public void ExternalSignatureFinalizationRejectsEmptyAndAmbiguousRawCompletion() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Signature finalization ambiguity source"))
            .ToBytes();
        PdfExternalSignaturePreparation first = PdfIncrementalUpdater.PrepareExternalSignature(
            source,
            new PdfExternalSignatureOptions {
                FieldName = "FirstApproval",
                ReservedSignatureContentsBytes = 256
            });
        byte[] secondPlaceholder = System.Text.Encoding.ASCII.GetBytes(
            "\n999999 0 obj\n<< /Type /Sig /ByteRange [0 0 0 0] /Contents <00000000> >>\nendobj\n");
        byte[] ambiguousPdf = first.PreparedPdf.Concat(secondPlaceholder).ToArray();

        Assert.Throws<ArgumentException>(() => first.Complete(Array.Empty<byte>()));
        Assert.Throws<ArgumentException>(() => PdfDocument.Open(first.PreparedPdf).CompleteExternalSignature(Array.Empty<byte>()));
        ArgumentException ambiguous = Assert.Throws<ArgumentException>(
            () => PdfDocument.Open(ambiguousPdf).CompleteExternalSignature(new byte[] { 0x30, 0x01, 0x00 }));

        Assert.Contains("multiple", ambiguous.Message, StringComparison.OrdinalIgnoreCase);
        Assert.NotEmpty(first.Complete(new byte[] { 0x30, 0x01, 0x00 }).ToBytes());
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

    private static byte[] ReplaceFirstAscii(byte[] source, string oldValue, string newValue) {
        byte[] oldBytes = System.Text.Encoding.ASCII.GetBytes(oldValue);
        byte[] newBytes = System.Text.Encoding.ASCII.GetBytes(newValue);
        Assert.Equal(oldBytes.Length, newBytes.Length);
        byte[] result = (byte[])source.Clone();
        for (int index = 0; index <= result.Length - oldBytes.Length; index++) {
            if (!result.AsSpan(index, oldBytes.Length).SequenceEqual(oldBytes)) continue;
            Buffer.BlockCopy(newBytes, 0, result, index, newBytes.Length);
            return result;
        }
        throw new InvalidOperationException("Expected ASCII marker was not found.");
    }

    private static byte[] WithCatalogAction(byte[] source, string actionType = "JavaScript") {
        return PdfDocumentObjectGraphRewriter.Rewrite(source, null, null, (objects, security) => {
            int rootObjectNumber = Assert.IsType<int>(security.RootObjectNumber);
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(objects[rootObjectNumber].Value);
            var action = new PdfDictionary();
            action.Items["S"] = new PdfName(actionType);
            if (actionType == "JavaScript") action.Items["JS"] = new PdfStringObj("app.alert('blocked')", true);
            catalog.Items["OpenAction"] = action;
            return security.InfoObjectNumber;
        });
    }

    private static byte[] BuildDuplicateContentsSignatureReservation(int reservedBytes) {
        string zeros = new string('0', reservedBytes * 2);
        string rangePlaceholder = string.Join(" ", Enumerable.Repeat(new string('0', 20), 4));
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /AcroForm 5 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 8 0 R /Annots [4 0 R] >>",
            "endobj",
            "4 0 obj",
            "<< /FT /Sig /T (Approval) /V 6 0 R /Subtype /Widget /Rect [0 0 0 0] >>",
            "endobj",
            "5 0 obj",
            "<< /Fields [4 0 R] /SigFlags 3 >>",
            "endobj",
            "6 0 obj",
            "<< /Type /Sig /Filter /Adobe.PPKLite /SubFilter /adbe.pkcs7.detached /ByteRange [" + rangePlaceholder + "] /Contents <" + zeros + "> /Contents 7 0 R >>",
            "endobj",
            "7 0 obj",
            "<" + zeros + ">",
            "endobj",
            "8 0 obj",
            "<< /Length 0 >>\nstream\n\nendstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 9 >>",
            "%%EOF"
        });
        int literalStart = pdf.IndexOf("/Contents <" + zeros + ">", StringComparison.Ordinal) + "/Contents ".Length;
        int literalEnd = literalStart + zeros.Length + 2;
        string byteRange = string.Join(" ", new[] {
            0L,
            (long)literalStart,
            (long)literalEnd,
            (long)pdf.Length - literalEnd
        }.Select(static value => value.ToString("00000000000000000000", System.Globalization.CultureInfo.InvariantCulture)));
        return System.Text.Encoding.ASCII.GetBytes(pdf.Replace(rangePlaceholder, byteRange));
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
