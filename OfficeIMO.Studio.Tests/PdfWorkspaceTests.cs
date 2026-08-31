using OfficeIMO.Studio.Features.Workspace;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Editor;

namespace OfficeIMO.Studio.Tests;

public sealed class PdfWorkspaceTests {
    [Fact]
    public async Task EditorMutationParticipatesInRecoveryUndoRedoAndJournal() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-editor-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string source = Path.Combine(root, "source.pdf");
        CreateEditableSource(source);
        var recovery = new PdfWorkspaceRecoveryStore(Path.Combine(root, "recovery"));
        var gesture = new PdfEditorGesture(1, 40D, 50D, 180D, 100D, Array.Empty<PdfEditorVisualPoint>());
        var properties = new PdfEditorProperties(
            "Workspace note",
            "OfficeIMO Studio",
            PdfColor.FromRgb(229, 72, 77),
            "Approved",
            "https://officeimo.com",
            14D);

        try {
            using PdfWorkspace workspace = await PdfWorkspace.OpenAsync(source, CancellationToken.None, recovery);

            await workspace.ApplyEditorGestureAsync(PdfEditorTool.Note, gesture, properties, CancellationToken.None);

            Assert.True(workspace.IsDirty);
            Assert.True(workspace.CanUndo);
            Assert.Equal(PdfWorkspaceOperationKind.Annotation, Assert.Single(workspace.Journal).Kind);
            Assert.Single(PdfDocument.Open(workspace.CopyBytes()).Inspect().GetAnnotationsBySubtype("Text"));

            await workspace.UndoAsync(CancellationToken.None);
            Assert.Empty(PdfDocument.Open(workspace.CopyBytes()).Inspect().GetAnnotationsBySubtype("Text"));
            await workspace.RedoAsync(CancellationToken.None);
            Assert.Single(PdfDocument.Open(workspace.CopyBytes()).Inspect().GetAnnotationsBySubtype("Text"));
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task FormFillAndExplicitFlattenProduceReadableCurrentArtifacts() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-form-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string source = Path.Combine(root, "form.pdf");
        PdfDocument.Create(compose => compose.Page(page => page.Content(content => content.Item(item =>
            item.TextField("Customer.Name", value: "Before"))))).Save(source);

        try {
            using PdfWorkspace workspace = await PdfWorkspace.OpenAsync(source, CancellationToken.None);
            Assert.True(workspace.CanFillForms);
            Assert.True(workspace.CanFlattenForms);

            await workspace.FillFormFieldAsync("Customer.Name", "After", flatten: false, CancellationToken.None);
            PdfFormField filled = Assert.Single(PdfDocument.Open(workspace.CopyBytes()).Inspect().FormFields);
            Assert.Equal("After", filled.Value);

            await workspace.FlattenFormFieldsAsync(CancellationToken.None);
            PdfDocument flattened = PdfDocument.Open(workspace.CopyBytes());
            Assert.Empty(flattened.Inspect().FormFields);
            Assert.Contains("After", flattened.Read.Text(), StringComparison.Ordinal);
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task WatermarkAndPageNumbersAreInspectableAndRenderableAfterSave() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-overlays-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string source = Path.Combine(root, "source.pdf");
        string saved = Path.Combine(root, "saved.pdf");
        CreateEditableSource(source);

        try {
            using PdfWorkspace workspace = await PdfWorkspace.OpenAsync(source, CancellationToken.None);
            await workspace.ApplyWatermarkAsync("INTERNAL", CancellationToken.None);
            await workspace.ApplyPageNumbersAsync(CancellationToken.None);
            await workspace.SaveAsync(saved, CancellationToken.None);

            PdfDocument reopened = PdfDocument.Open(saved);
            string text = reopened.Read.Text();
            Assert.Contains("INTERNAL", text, StringComparison.Ordinal);
            Assert.Contains("1 / 1", text, StringComparison.Ordinal);
            Assert.True(Assert.Single(reopened.Read.RenderPages("1", new PdfPageRenderOptions { Format = PdfPageRenderFormat.Svg })).Succeeded);
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task CertificationPolicyAllowsAppendOnlyReviewButBlocksPageContentEdits() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-certified-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string source = Path.Combine(root, "certified.pdf");
        byte[] unsigned = PdfDocument.Create(compose => compose.Page(page => page.Content(content => content.Item(item =>
            item.Paragraph(paragraph => paragraph.Text("Certified content")))))).ToBytes();
        PdfExternalSignaturePreparation preparation = PdfDocument.Open(unsigned).Security.PrepareExternalSignature(new PdfExternalSignatureOptions {
            Profile = PdfSignatureProfile.Certification,
            CertificationPermission = PdfCertificationPermissionLevel.FormFillingAnnotationsAndSignatures,
            FieldName = "Certification",
            ReservedSignatureContentsBytes = 512
        });
        byte[] signed = preparation.Complete(new byte[] { 0x30, 0x01, 0x00 }).ToBytes();
        await File.WriteAllBytesAsync(source, signed);
        var gesture = new PdfEditorGesture(1, 40D, 50D, 60D, 70D, Array.Empty<PdfEditorVisualPoint>());
        var properties = new PdfEditorProperties("Certified review", "Studio", PdfColor.FromRgb(220, 38, 38), "Approved", "https://officeimo.com", 12D);

        try {
            using PdfWorkspace workspace = await PdfWorkspace.OpenAsync(source, CancellationToken.None);
            Assert.True(workspace.CanEditAnnotations);
            Assert.False(workspace.CanEditPageContent);

            await workspace.ApplyEditorGestureAsync(PdfEditorTool.Note, gesture, properties, CancellationToken.None);

            byte[] edited = workspace.CopyBytes();
            Assert.True(edited.AsSpan(0, signed.Length).SequenceEqual(signed));
            Assert.Single(PdfDocument.Open(edited).Inspect().GetAnnotationsBySubtype("Text"));
            await Assert.ThrowsAnyAsync<Exception>(() => workspace.ApplyWatermarkAsync("BLOCKED", CancellationToken.None));
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task FailedRedactionProofDoesNotCommitCandidateBytesOrHistory() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-redaction-proof-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string source = Path.Combine(root, "source.pdf");
        PdfDocument.Create(compose => compose.Page(page => page.Content(content => content.Item(item =>
            item.Paragraph(paragraph => paragraph.Text("Keep this marker")))))).Save(source);
        var gesture = new PdfEditorGesture(1, 300D, 300D, 340D, 340D, Array.Empty<PdfEditorVisualPoint>());
        var properties = new PdfEditorProperties(string.Empty, "Studio", PdfColor.Black, "Approved", "https://officeimo.com", 12D);

        try {
            using PdfWorkspace workspace = await PdfWorkspace.OpenAsync(source, CancellationToken.None);
            byte[] original = workspace.CopyBytes();
            PdfRedactionPlan plan = await workspace.PlanRedactionAsync(gesture, properties, CancellationToken.None);

            await Assert.ThrowsAsync<InvalidOperationException>(() => workspace.ApplyVerifiedRedactionAsync(
                plan,
                workspace.Revision,
                "Keep this marker",
                CancellationToken.None));

            Assert.Equal(original, workspace.CopyBytes());
            Assert.False(workspace.IsDirty);
            Assert.False(workspace.CanUndo);
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task ReviewedRedactionPlanCannotApplyAfterWorkspaceRevisionChanges() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-stale-redaction-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string source = Path.Combine(root, "source.pdf");
        CreateEditableSource(source);
        var gesture = new PdfEditorGesture(1, 40D, 50D, 180D, 100D, Array.Empty<PdfEditorVisualPoint>());
        var properties = new PdfEditorProperties("Review", "Studio", PdfColor.Black, "Approved", "https://officeimo.com", 12D);

        try {
            using PdfWorkspace workspace = await PdfWorkspace.OpenAsync(source, CancellationToken.None);
            long reviewedRevision = workspace.Revision;
            PdfRedactionPlan plan = await workspace.PlanRedactionAsync(gesture, properties, CancellationToken.None);
            await workspace.ApplyEditorGestureAsync(PdfEditorTool.Note, gesture, properties, CancellationToken.None);
            byte[] afterNote = workspace.CopyBytes();

            InvalidOperationException exception = await Assert.ThrowsAsync<InvalidOperationException>(() =>
                workspace.ApplyVerifiedRedactionAsync(plan, reviewedRevision, null, CancellationToken.None));

            Assert.Contains("changed after this redaction was reviewed", exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(afterNote, workspace.CopyBytes());
            Assert.Single(PdfDocument.Open(afterNote).Inspect().GetAnnotationsBySubtype("Text"));
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task MutationUndoRedoAndSaveMaintainWorkspaceContract() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-workspace-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string source = Path.Combine(root, "source.pdf");
        string saved = Path.Combine(root, "saved.pdf");
        CreateEditableSource(source);
        var recovery = new PdfWorkspaceRecoveryStore(Path.Combine(root, "recovery"));

        try {
            using PdfWorkspace workspace = await PdfWorkspace.OpenAsync(source, CancellationToken.None, recovery);

            await workspace.DuplicateAsync([1], CancellationToken.None);
            Assert.True(workspace.IsDirty);
            Assert.True(workspace.CanUndo);
            Assert.Equal(2, workspace.Pages.Count);
            Assert.False(workspace.HasRecovery);

            await workspace.UndoAsync(CancellationToken.None);
            Assert.Single(workspace.Pages);
            Assert.True(workspace.CanRedo);

            await workspace.RedoAsync(CancellationToken.None);
            Assert.Equal(2, workspace.Pages.Count);

            await workspace.SaveAsync(saved, CancellationToken.None);
            Assert.False(workspace.IsDirty);
            Assert.Equal(saved, workspace.Path);
            Assert.True(File.Exists(saved));
            Assert.False(workspace.HasRecovery);
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task ReorderRotateCropImportAndBlankProduceReadableArtifacts() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-organizer-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string source = Path.Combine(root, "source.pdf");
        CreateEditableSource(source);

        try {
            using PdfWorkspace workspace = await PdfWorkspace.OpenAsync(
                source,
                CancellationToken.None,
                new PdfWorkspaceRecoveryStore(Path.Combine(root, "recovery")));

            await workspace.DuplicateAsync([1], CancellationToken.None);
            await workspace.ReorderAsync([2, 1], CancellationToken.None);
            await workspace.RotateAsync([1], 90, CancellationToken.None);
            await workspace.CropAsync([2], 10, 10, 580, 760, CancellationToken.None);
            await workspace.ImportAsync(source, 2, CancellationToken.None);
            await workspace.InsertBlankAsync(workspace.Pages.Count + 1, 400, 500, CancellationToken.None);

            Assert.Equal(4, workspace.Pages.Count);
            Assert.Equal(90, workspace.Pages[0].RotationDegrees);
            Assert.Equal(6, workspace.Journal.Count);
            Assert.NotEmpty(workspace.CopyBytes());
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task RecoveryCanBeDiscoveredRestoredAndDiscarded() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-recovery-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string source = Path.Combine(root, "source.pdf");
        CreateEditableSource(source);
        var recovery = new PdfWorkspaceRecoveryStore(Path.Combine(root, "recovery"));

        try {
            using (PdfWorkspace edited = await PdfWorkspace.OpenAsync(source, CancellationToken.None, recovery)) {
                await edited.DuplicateAsync([1], CancellationToken.None);
                Assert.False(edited.HasRecovery);
            }

            using PdfWorkspace reopened = await PdfWorkspace.OpenAsync(source, CancellationToken.None, recovery);
            Assert.True(reopened.HasRecovery);
            Assert.Single(reopened.Pages);

            await reopened.RestoreRecoveryAsync(CancellationToken.None);

            Assert.Equal(2, reopened.Pages.Count);
            Assert.True(reopened.IsDirty);
            reopened.DiscardRecovery();
            Assert.False(reopened.HasRecovery);
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task TaggedPdfReportsPageMutationAsUnavailableBeforeAnEditStarts() {
        string fixture = Path.Combine(AppContext.BaseDirectory, "Fixtures", "openpreserve-pdfa1b-text.pdf");
        using PdfWorkspace workspace = await PdfWorkspace.OpenAsync(fixture, CancellationToken.None);

        Assert.False(workspace.CanMutatePages);
        Assert.NotNull(workspace.SecurityWarning);
        Assert.Contains("tagged", workspace.SecurityWarning, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task FailedRecoveryWriteDoesNotCommitAnInMemoryMutation() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-transaction-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string source = Path.Combine(root, "source.pdf");
        string blockedRecoveryRoot = Path.Combine(root, "not-a-directory");
        CreateEditableSource(source);
        await File.WriteAllTextAsync(blockedRecoveryRoot, "block directory creation");

        try {
            using PdfWorkspace workspace = await PdfWorkspace.OpenAsync(
                source,
                CancellationToken.None,
                new PdfWorkspaceRecoveryStore(blockedRecoveryRoot));
            byte[] original = workspace.CopyBytes();

            await Assert.ThrowsAnyAsync<IOException>(() => workspace.DuplicateAsync([1], CancellationToken.None));

            Assert.Single(workspace.Pages);
            Assert.False(workspace.IsDirty);
            Assert.False(workspace.CanUndo);
            Assert.Equal(original, workspace.CopyBytes());
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task RecoveryIsRejectedWhenTheSourceAtTheSamePathWasReplaced() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-recovery-identity-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string source = Path.Combine(root, "source.pdf");
        var recovery = new PdfWorkspaceRecoveryStore(Path.Combine(root, "recovery"));
        CreateEditableSource(source);

        try {
            using (PdfWorkspace edited = await PdfWorkspace.OpenAsync(source, CancellationToken.None, recovery)) {
                await edited.DuplicateAsync([1], CancellationToken.None);
            }
            PdfDocument.Create(compose => compose.Page(page => page.Size(300, 400))).Save(source);

            using PdfWorkspace replaced = await PdfWorkspace.OpenAsync(source, CancellationToken.None, recovery);

            Assert.False(replaced.HasRecovery);
            Assert.Single(replaced.Pages);
            Assert.Equal(300D, replaced.Pages[0].Width);
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task ExtractRejectsTheOpenDocumentPathWithoutChangingTheSource() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-extract-identity-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string source = Path.Combine(root, "source.pdf");
        CreateEditableSource(source);
        byte[] original = await File.ReadAllBytesAsync(source);

        try {
            using PdfWorkspace workspace = await PdfWorkspace.OpenAsync(source, CancellationToken.None);

            InvalidOperationException exception = await Assert.ThrowsAsync<InvalidOperationException>(
                () => workspace.ExtractAsync([1], source, CancellationToken.None));

            Assert.Contains("different file", exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(original, await File.ReadAllBytesAsync(source));
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task NewEditAfterUndoCannotReuseTheSavedRevisionIdentity() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-revision-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string source = Path.Combine(root, "source.pdf");
        CreateEditableSource(source);

        try {
            using PdfWorkspace workspace = await PdfWorkspace.OpenAsync(source, CancellationToken.None);
            await workspace.DuplicateAsync([1], CancellationToken.None);
            await workspace.SaveAsync(path: null, CancellationToken.None);
            await workspace.DuplicateAsync([1], CancellationToken.None);
            await workspace.UndoAsync(CancellationToken.None);
            Assert.False(workspace.IsDirty);

            await workspace.InsertBlankAsync(1, 200, 300, CancellationToken.None);

            Assert.True(workspace.IsDirty);
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task FailedRecoveryWriteDuringRedoLeavesHistoryAndDocumentConsistent() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-redo-transaction-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string source = Path.Combine(root, "source.pdf");
        string recoveryRoot = Path.Combine(root, "recovery");
        CreateEditableSource(source);

        try {
            using PdfWorkspace workspace = await PdfWorkspace.OpenAsync(
                source,
                CancellationToken.None,
                new PdfWorkspaceRecoveryStore(recoveryRoot));
            await workspace.DuplicateAsync([1], CancellationToken.None);
            await workspace.UndoAsync(CancellationToken.None);
            Directory.Delete(recoveryRoot, recursive: true);
            await File.WriteAllTextAsync(recoveryRoot, "block directory creation");

            await Assert.ThrowsAnyAsync<IOException>(() => workspace.RedoAsync(CancellationToken.None));

            Assert.Single(workspace.Pages);
            Assert.False(workspace.IsDirty);
            Assert.True(workspace.CanRedo);
            Assert.False(workspace.CanUndo);
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    private static void CreateEditableSource(string path) =>
        PdfDocument.Create(compose => compose.Page(page => page.Size(600, 800))).Save(path);
}
