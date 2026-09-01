using OfficeIMO.Studio.Features.Workspace;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Editor;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;

namespace OfficeIMO.Studio.Tests;

public sealed class PdfWorkspaceTests {
    [Fact]
    public async Task ExistingTextSelectionSupportsReplaceMoveDeleteAndDocumentWideReplace() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-existing-text-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string source = Path.Combine(root, "text.pdf");
        CreateTextSource(source, "Alpha target Omega");

        try {
            using PdfWorkspace workspace = await PdfWorkspace.OpenAsync(source, CancellationToken.None);
            PdfEditorSelection selection = CreateTextSelection(workspace.CopyBytes(), "target");

            await workspace.ReplaceSelectedTextAsync(selection, "replacement", options: null, CancellationToken.None);
            string replacedText = PdfDocument.Load(workspace.CopyBytes()).Read().Text;
            Assert.Contains("Alpha", replacedText, StringComparison.Ordinal);
            Assert.Contains("replacement", replacedText, StringComparison.Ordinal);
            Assert.Contains("Omega", replacedText, StringComparison.Ordinal);
            Assert.Equal(PdfWorkspaceOperationKind.TextEdit, workspace.Journal[^1].Kind);

            PdfEditorSelection movedSelection = CreateTextSelection(workspace.CopyBytes(), "replacement");
            await workspace.MoveSelectedTextAsync(movedSelection, 12D, -30D, CancellationToken.None);
            string movedText = PdfDocument.Load(workspace.CopyBytes()).Read().Text;
            Assert.Contains("Alpha", movedText, StringComparison.Ordinal);
            Assert.Contains("replacement", movedText, StringComparison.Ordinal);
            Assert.Contains("Omega", movedText, StringComparison.Ordinal);

            await workspace.ReplaceAllTextAsync("replacement", "final", matchCase: true, wholeWords: true, CancellationToken.None);
            Assert.Contains("final", PdfDocument.Load(workspace.CopyBytes()).Read().Text, StringComparison.Ordinal);

            PdfEditorSelection deleteSelection = CreateTextSelection(workspace.CopyBytes(), "final");
            await workspace.ReplaceSelectedTextAsync(deleteSelection, string.Empty, options: null, CancellationToken.None);
            string deletedText = PdfDocument.Load(workspace.CopyBytes()).Read().Text;
            Assert.Contains("Alpha", deletedText, StringComparison.Ordinal);
            Assert.DoesNotContain("final", deletedText, StringComparison.Ordinal);
            Assert.Contains("Omega", deletedText, StringComparison.Ordinal);
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task ExactImageSelectionSupportsMoveReplaceAndDelete() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-existing-image-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string source = Path.Combine(root, "image.pdf");
        PdfDocument baseDocument = PdfDocument.Create(compose => compose.Page(page => page.Size(600D, 800D)));
        baseDocument.Images.Add(new PdfPageRegion(1, 50D, 60D, 40D, 20D), TinyPng).Document.Save(source);

        try {
            using PdfWorkspace workspace = await PdfWorkspace.OpenAsync(source, CancellationToken.None);
            PdfEditorSelection selection = CreateImageSelection(workspace.CopyBytes());

            await workspace.MoveSelectedImageAsync(selection, 15D, 5D, CancellationToken.None);
            PdfImagePlacement moved = Assert.Single(PdfDocument.Load(workspace.CopyBytes()).Images.Placements());
            Assert.Equal(65D, moved.X, 3);
            Assert.Equal(65D, moved.Y, 3);

            PdfEditorSelection replacementSelection = CreateImageSelection(workspace.CopyBytes());
            await workspace.ReplaceSelectedImageAsync(replacementSelection, TinyPng, CancellationToken.None);
            Assert.Single(PdfDocument.Load(workspace.CopyBytes()).Images.Placements());

            PdfEditorSelection removalSelection = CreateImageSelection(workspace.CopyBytes());
            await workspace.RemoveSelectedImageAsync(removalSelection, CancellationToken.None);
            Assert.Empty(PdfDocument.Load(workspace.CopyBytes()).Images.Placements());
            Assert.Equal(PdfWorkspaceOperationKind.ImageEdit, workspace.Journal[^1].Kind);
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

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
            Assert.Single(PdfDocument.Load(workspace.CopyBytes()).Inspect().GetAnnotationsBySubtype("Text"));

            await workspace.UndoAsync(CancellationToken.None);
            Assert.Empty(PdfDocument.Load(workspace.CopyBytes()).Inspect().GetAnnotationsBySubtype("Text"));
            await workspace.RedoAsync(CancellationToken.None);
            Assert.Single(PdfDocument.Load(workspace.CopyBytes()).Inspect().GetAnnotationsBySubtype("Text"));
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
            PdfFormField filled = Assert.Single(PdfDocument.Load(workspace.CopyBytes()).Inspect().FormFields);
            Assert.Equal("After", filled.Value);

            await workspace.FlattenFormFieldsAsync(CancellationToken.None);
            PdfDocument flattened = PdfDocument.Load(workspace.CopyBytes());
            Assert.Empty(flattened.Inspect().FormFields);
            Assert.Contains("After", flattened.Read().Text, StringComparison.Ordinal);
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task TypedFormEditingAuthoringAndSelectiveFlatteningUseCanonicalFieldContracts() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-typed-forms-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string source = Path.Combine(root, "forms.pdf");
        CreateEditableSource(source);

        try {
            using PdfWorkspace workspace = await PdfWorkspace.OpenAsync(source, CancellationToken.None);
            await workspace.CreateFormFieldAsync(new PdfFormFieldCreateOptions {
                Name = "Account.Secret",
                Kind = PdfFormFieldCreationKind.Text,
                PageNumber = 1,
                X = 40,
                Y = 650,
                Width = 180,
                Height = 24,
                Style = new PdfFormFieldStyle { IsPassword = true }
            }, CancellationToken.None);
            await workspace.CreateFormFieldAsync(new PdfFormFieldCreateOptions {
                Name = "Account.Enabled",
                Kind = PdfFormFieldCreationKind.CheckBox,
                PageNumber = 1,
                X = 40,
                Y = 610,
                Width = 18,
                Height = 18,
                Value = "Off"
            }, CancellationToken.None);
            await workspace.CreateFormFieldAsync(new PdfFormFieldCreateOptions {
                Name = "Account.Regions",
                Kind = PdfFormFieldCreationKind.Choice,
                PageNumber = 1,
                X = 40,
                Y = 520,
                Width = 180,
                Height = 60,
                ChoiceOptions = ["EU", "US", "APAC"],
                Value = "EU",
                FieldFlags = 2097152
            }, CancellationToken.None);

            PdfDocumentInfo authored = PdfDocument.Load(workspace.CopyBytes()).Inspect();
            Assert.Single(workspace.Pages);
            Assert.True(authored.FormFieldsByName["Account.Secret"].IsPassword);
            Assert.True(authored.FormFieldsByName["Account.Enabled"].IsCheckBox);
            Assert.True(authored.FormFieldsByName["Account.Regions"].AllowsMultipleSelection);
            Assert.Equal(PdfWorkspaceOperationKind.FormAuthor, workspace.Journal[^1].Kind);

            var checkBox = new PdfFormFieldViewModel(authored.FormFieldsByName["Account.Enabled"]) { IsChecked = true };
            await workspace.FillFormFieldAsync(checkBox.Name, checkBox.CreateValue(), flatten: false, CancellationToken.None);
            Assert.Equal("Yes", PdfDocument.Load(workspace.CopyBytes()).Inspect().FormFieldsByName[checkBox.Name].Value);

            var regions = new PdfFormFieldViewModel(PdfDocument.Load(workspace.CopyBytes()).Inspect().FormFieldsByName["Account.Regions"]);
            foreach (PdfFormChoiceViewModel choice in regions.Choices) {
                choice.IsSelected = choice.ExportValue is "US" or "APAC";
            }
            await workspace.FillFormFieldAsync(regions.Name, regions.CreateValue(), flatten: false, CancellationToken.None);
            PdfFormField persistedRegions = PdfDocument.Load(workspace.CopyBytes()).Inspect().FormFieldsByName[regions.Name];
            Assert.Equal(["US", "APAC"], persistedRegions.Values);
            var reopenedRegions = new PdfFormFieldViewModel(persistedRegions);
            Assert.Equal(2, reopenedRegions.Choices.Count(static choice => choice.IsSelected));
            Assert.Equal(["US", "APAC"], reopenedRegions.CreateValue().Values);

            await workspace.FlattenFormFieldAsync("Account.Enabled", CancellationToken.None);
            PdfDocumentInfo selectivelyFlattened = PdfDocument.Load(workspace.CopyBytes()).Inspect();
            Assert.DoesNotContain("Account.Enabled", selectivelyFlattened.FormFieldsByName.Keys);
            Assert.Contains("Account.Secret", selectivelyFlattened.FormFieldsByName.Keys);
            Assert.Contains("Account.Regions", selectivelyFlattened.FormFieldsByName.Keys);
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task ProtectAndDecryptCopiesPreserveTheOpenWorkspaceAndApplyTypedPermissions() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-protection-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string source = Path.Combine(root, "source.pdf");
        string protectedCopy = Path.Combine(root, "protected.pdf");
        string decryptedCopy = Path.Combine(root, "decrypted.pdf");
        CreateTextSource(source, "Protected copy source");

        try {
            using (PdfWorkspace workspace = await PdfWorkspace.OpenAsync(source, CancellationToken.None)) {
                var encryption = new PdfStandardEncryptionOptions("open") {
                    OwnerPassword = "owner",
                    AllowedPermissions = PdfStandardPermissions.Print | PdfStandardPermissions.CopyContents |
                                         PdfStandardPermissions.Accessibility | PdfStandardPermissions.FillForms
                };
                await workspace.SaveProtectedCopyAsync(protectedCopy, encryption, null, CancellationToken.None);

                Assert.False(workspace.IsEncrypted);
                Assert.False(workspace.IsDirty);
                Assert.Equal(await File.ReadAllBytesAsync(source), workspace.CopyBytes());
            }

            Assert.Throws<PdfPasswordRequiredException>(() => PdfDocument.Load(protectedCopy).Inspect());
            PdfDocument protectedDocument = PdfDocument.Load(protectedCopy, new PdfLoadOptions { Password = "owner" });
            Assert.True(protectedDocument.Inspect().Security.HasEncryption);
            Assert.Equal(
                PdfStandardPermissions.Print | PdfStandardPermissions.CopyContents |
                PdfStandardPermissions.Accessibility | PdfStandardPermissions.FillForms,
                protectedDocument.Inspect().Security.AllowedStandardPermissions);

            using PdfWorkspace encryptedWorkspace = await PdfWorkspace.OpenAsync(
                protectedCopy,
                CancellationToken.None,
                password: "open");
            Assert.False(encryptedWorkspace.CanChangeEncryption(ownerPassword: null));
            Assert.True(encryptedWorkspace.CanChangeEncryption("owner"));
            await encryptedWorkspace.SaveDecryptedCopyAsync(decryptedCopy, "owner", CancellationToken.None);
            Assert.False(PdfDocument.Load(decryptedCopy).Inspect().Security.HasEncryption);
            Assert.Contains("Protected copy source", PdfDocument.Load(decryptedCopy).Read().Text, StringComparison.Ordinal);
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task BatesNumberingAndCertificateSigningProduceInspectableCurrentArtifacts() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-security-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string source = Path.Combine(root, "source.pdf");
        CreateTextSource(source, "Page one", "Page two");

        try {
            using PdfWorkspace workspace = await PdfWorkspace.OpenAsync(source, CancellationToken.None);
            await workspace.ApplyBatesNumberingAsync(new PdfBatesNumberingOptions {
                StartNumber = 42,
                MinimumDigits = 4,
                Prefix = "CASE-",
                Position = PdfBatesPosition.BottomCenter
            }, CancellationToken.None);
            string numberedText = PdfDocument.Load(workspace.CopyBytes()).Read().Text;
            Assert.Contains("CASE-0042", numberedText, StringComparison.Ordinal);
            Assert.Contains("CASE-0043", numberedText, StringComparison.Ordinal);
            Assert.Equal(PdfWorkspaceOperationKind.BatesNumbering, workspace.Journal[^1].Kind);

            using X509Certificate2 certificate = CreateSigningCertificate();
            await workspace.SignAsync(certificate, new PdfExternalSignatureOptions {
                FieldName = "Approval",
                Name = "Studio test signer",
                Reason = "Verified workflow"
            }, CancellationToken.None);
            PdfSignatureValidationReport report = await workspace.ValidateSignaturesAsync(CancellationToken.None);
            Assert.Single(report.Signatures);
            Assert.True(report.IsStructurallyValid);
            Assert.True(report.MathematicalSignaturesVerified);
            Assert.True(report.DigestVerified);
            Assert.Equal(PdfWorkspaceOperationKind.Signature, workspace.Journal[^1].Kind);
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

            PdfDocument reopened = PdfDocument.Load(saved);
            string text = reopened.Read().Text;
            Assert.Contains("INTERNAL", text, StringComparison.Ordinal);
            Assert.Contains("1 / 1", text, StringComparison.Ordinal);
            Assert.True(Assert.Single(reopened.Render.Pages("1", new PdfPageRenderOptions { Format = PdfPageRenderFormat.Svg })).Succeeded);
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
        PdfExternalSignaturePreparation preparation = PdfDocument.Load(unsigned).Security.PrepareExternalSignature(new PdfExternalSignatureOptions {
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
            Assert.False(workspace.CanChangeEncryption(ownerPassword: null));

            await workspace.ApplyEditorGestureAsync(PdfEditorTool.Note, gesture, properties, CancellationToken.None);

            byte[] edited = workspace.CopyBytes();
            Assert.True(edited.AsSpan(0, signed.Length).SequenceEqual(signed));
            Assert.Single(PdfDocument.Load(edited).Inspect().GetAnnotationsBySubtype("Text"));
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
            Assert.Single(PdfDocument.Load(afterNote).Inspect().GetAnnotationsBySubtype("Text"));
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
    public async Task MultiPdfImportIsOneUndoableMutationAndPreservesPickerOrder() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-multi-import-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string target = Path.Combine(root, "target.pdf");
        string first = Path.Combine(root, "first.pdf");
        string second = Path.Combine(root, "second.pdf");
        CreateTextSource(target, "Target");
        CreateTextSource(first, "First A", "First B");
        CreateTextSource(second, "Second");

        try {
            using PdfWorkspace workspace = await PdfWorkspace.OpenAsync(target, CancellationToken.None);

            int importedPageCount = await workspace.ImportAsync([first, second], 1, CancellationToken.None);

            Assert.Equal(3, importedPageCount);
            Assert.Equal(4, workspace.Pages.Count);
            Assert.Equal(
                ["First A", "First B", "Second", "Target"],
                PdfReadDocument.Open(workspace.CopyBytes()).Pages.Select(page => page.ExtractText().Trim()).ToArray());
            Assert.Single(workspace.Journal);
            Assert.Equal(PdfWorkspaceOperationKind.Import, workspace.Journal[0].Kind);

            await workspace.UndoAsync(CancellationToken.None);

            Assert.Single(workspace.Pages);
            Assert.Contains("Target", PdfReadDocument.Open(workspace.CopyBytes()).ExtractText(), StringComparison.Ordinal);
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task MalformedLaterImportDoesNotCommitEarlierSources() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-multi-import-failure-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string target = Path.Combine(root, "target.pdf");
        string valid = Path.Combine(root, "valid.pdf");
        string malformed = Path.Combine(root, "malformed.pdf");
        CreateTextSource(target, "Target");
        CreateTextSource(valid, "Would have imported");
        await File.WriteAllBytesAsync(malformed, "%PDF-1.7\nbroken"u8.ToArray());

        try {
            using PdfWorkspace workspace = await PdfWorkspace.OpenAsync(target, CancellationToken.None);
            byte[] original = workspace.CopyBytes();

            await Assert.ThrowsAnyAsync<Exception>(
                () => workspace.ImportAsync([valid, malformed], 1, CancellationToken.None));

            Assert.Equal(original, workspace.CopyBytes());
            Assert.False(workspace.IsDirty);
            Assert.False(workspace.CanUndo);
            Assert.Empty(workspace.Journal);
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task SplitPublishesCompleteBatchUsingConfiguredPageCount() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-split-batch-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string source = Path.Combine(root, "source.pdf");
        string output = Path.Combine(root, "split");
        CreateTextSource(source, "One", "Two", "Three");

        try {
            using PdfWorkspace workspace = await PdfWorkspace.OpenAsync(source, CancellationToken.None);

            IReadOnlyList<string> outputs = await workspace.SplitAsync(output, 2, CancellationToken.None);

            Assert.Equal(2, outputs.Count);
            Assert.Equal(2, PdfDocument.Load(outputs[0]).Inspect().PageCount);
            Assert.Single(PdfDocument.Load(outputs[1]).Inspect().Pages);
            Assert.Empty(Directory.EnumerateDirectories(output, ".officeimo-studio-split-*"));
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task SplitCollisionLeavesExistingOutputsAndBatchUntouched() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-split-collision-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string source = Path.Combine(root, "source.pdf");
        string output = Path.Combine(root, "split");
        Directory.CreateDirectory(output);
        CreateTextSource(source, "One", "Two", "Three");
        string collision = Path.Combine(output, "source-part-001.pdf");
        byte[] existing = [1, 2, 3, 4];
        await File.WriteAllBytesAsync(collision, existing);

        try {
            using PdfWorkspace workspace = await PdfWorkspace.OpenAsync(source, CancellationToken.None);

            IOException exception = await Assert.ThrowsAsync<IOException>(
                () => workspace.SplitAsync(output, 2, CancellationToken.None));

            Assert.Contains("already", exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(existing, await File.ReadAllBytesAsync(collision));
            Assert.False(File.Exists(Path.Combine(output, "source-part-002.pdf")));
            Assert.Empty(Directory.EnumerateDirectories(output, ".officeimo-studio-split-*"));
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task CancelledSplitRemovesItsStagingAndPublishesNoOutputs() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-split-cancel-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string source = Path.Combine(root, "source.pdf");
        string output = Path.Combine(root, "split");
        CreateTextSource(source, "One", "Two", "Three");
        using var cancellation = new CancellationTokenSource();
        var progress = new InlineProgress<PdfWorkspaceProgress>(value => {
            if (value.Stage == "Preparing part 2 of 3") cancellation.Cancel();
        });

        try {
            using PdfWorkspace workspace = await PdfWorkspace.OpenAsync(source, CancellationToken.None);

            await Assert.ThrowsAnyAsync<OperationCanceledException>(
                () => workspace.SplitAsync(output, 1, cancellation.Token, progress));

            Assert.False(Directory.Exists(output));
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
        Assert.False(workspace.CanExtractPages);
        Assert.False(workspace.CanImportPages);
        Assert.NotNull(workspace.SecurityWarning);
        Assert.Contains("tagged", workspace.SecurityWarning, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task CancellationReturnsPromptlyWithoutStartingConcurrentCpuRewrites() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-cpu-cancel-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string source = Path.Combine(root, "source.pdf");
        string otherSource = Path.Combine(root, "other.pdf");
        CreateEditableSource(source);
        CreateEditableSource(otherSource);
        using var started = new ManualResetEventSlim();
        using var release = new ManualResetEventSlim();
        using var cancellation = new CancellationTokenSource();
        PdfWorkspace? firstWorkspace = null;

        try {
            firstWorkspace = await PdfWorkspace.OpenAsync(source, CancellationToken.None);
            Task<int> operation = firstWorkspace.RunCancellableCpuWorkAsync(() => {
                started.Set();
                release.Wait();
                return 42;
            }, cancellation.Token);

            Assert.True(started.Wait(TimeSpan.FromSeconds(5)));
            cancellation.Cancel();
            await Assert.ThrowsAnyAsync<OperationCanceledException>(
                () => operation.WaitAsync(TimeSpan.FromSeconds(2)));
            firstWorkspace.Dispose();
            firstWorkspace = null;

            using PdfWorkspace nextWorkspace = await PdfWorkspace.OpenAsync(otherSource, CancellationToken.None);
            int concurrentStarts = 0;
            using var waitingCancellation = new CancellationTokenSource(TimeSpan.FromMilliseconds(250));
            Task<int> waitingOperation = nextWorkspace.RunCancellableCpuWorkAsync(
                () => Interlocked.Increment(ref concurrentStarts),
                waitingCancellation.Token);
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() => waitingOperation);
            Assert.Equal(0, Volatile.Read(ref concurrentStarts));

            release.Set();
            int followUp = await nextWorkspace.RunCancellableCpuWorkAsync(() => 7, CancellationToken.None)
                .WaitAsync(TimeSpan.FromSeconds(2));
            Assert.Equal(7, followUp);
        } finally {
            release.Set();
            firstWorkspace?.Dispose();
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task NonDetachableCpuWorkKeepsItsCallerAttachedUntilTheWorkerFinishes() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-nondetachable-worker-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string source = Path.Combine(root, "source.pdf");
        CreateEditableSource(source);
        using var started = new ManualResetEventSlim();
        using var release = new ManualResetEventSlim();
        using var cancellation = new CancellationTokenSource();

        try {
            using PdfWorkspace workspace = await PdfWorkspace.OpenAsync(source, CancellationToken.None);
            Task<int> operation = workspace.RunNonDetachableCpuWorkAsync(() => {
                started.Set();
                release.Wait();
                return 42;
            }, cancellation.Token);

            Assert.True(started.Wait(TimeSpan.FromSeconds(5)));
            cancellation.Cancel();
            await Task.Delay(100);
            Assert.False(operation.IsCompleted);

            release.Set();
            Assert.Equal(42, await operation.WaitAsync(TimeSpan.FromSeconds(2)));
        } finally {
            release.Set();
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task PasswordProtectedPdfRequiresCredentialsBeforeWorkspaceCreation() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-encrypted-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string source = Path.Combine(root, "encrypted.pdf");
        byte[] plain = PdfDocument.Create(compose => compose.Page(page => page.Content(content => content.Item(item =>
            item.Paragraph(paragraph => paragraph.Text("Protected content")))))).ToBytes();
        var encryption = new PdfStandardEncryptionOptions("open") { OwnerPassword = "owner" };
        await File.WriteAllBytesAsync(source, PdfDocument.Load(plain).Security.Encrypt(encryption).Pdf);

        try {
            PdfPasswordRequiredException exception = await Assert.ThrowsAsync<PdfPasswordRequiredException>(
                () => PdfWorkspace.OpenAsync(source, CancellationToken.None));

            Assert.Contains("password", exception.Message, StringComparison.OrdinalIgnoreCase);
        } finally {
            Directory.Delete(root, recursive: true);
        }
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

    private static X509Certificate2 CreateSigningCertificate() {
        using RSA rsa = RSA.Create(2048);
        var request = new CertificateRequest(
            "CN=OfficeIMO Studio Test",
            rsa,
            HashAlgorithmName.SHA256,
            RSASignaturePadding.Pkcs1);
        request.CertificateExtensions.Add(new X509KeyUsageExtension(X509KeyUsageFlags.DigitalSignature, critical: true));
        return request.CreateSelfSigned(DateTimeOffset.UtcNow.AddMinutes(-5), DateTimeOffset.UtcNow.AddDays(1));
    }

    private static void CreateTextSource(string path, params string[] pageTexts) =>
        PdfDocument.Create(compose => {
            foreach (string text in pageTexts) {
                compose.Page(page => page.Content(content => content.Item(item =>
                    item.Paragraph(paragraph => paragraph.Text(text)))));
            }
        }).Save(path);

    private static PdfEditorSelection CreateTextSelection(byte[] pdf, string text) {
        PdfPageInteractionMap map = PdfPageInteractionMap.Create(pdf, 1);
        int firstIndex = FindSequence(map.TextRegions.Select(static region => region.Text).ToArray(), text);
        PdfPageInteractionRegion[] regions = map.TextRegions
            .Where((_, index) => index >= firstIndex)
            .Take(text.Length)
            .ToArray();
        Assert.Equal(text, string.Concat(regions.Select(static region => region.Text)));
        return new PdfEditorSelection(
            PdfEditorSelectionKind.Text,
            1,
            new PdfEditorVisualBounds(
                regions.Min(static region => region.Quad.Left),
                regions.Min(static region => region.Quad.Top),
                regions.Max(static region => region.Quad.Right),
                regions.Max(static region => region.Quad.Bottom)),
            Text: text);
    }

    private static PdfEditorSelection CreateImageSelection(byte[] pdf) {
        PdfPageInteractionRegion region = Assert.Single(
            PdfPageInteractionMap.Create(pdf, 1).Regions,
            candidate => candidate.Kind == PdfInteractionKind.Image);
        return new PdfEditorSelection(
            PdfEditorSelectionKind.Image,
            1,
            new PdfEditorVisualBounds(region.Quad.Left, region.Quad.Top, region.Quad.Right, region.Quad.Bottom),
            ImagePlacement: region.ImagePlacement);
    }

    private static int FindSequence(IReadOnlyList<string?> elements, string text) {
        string joined = string.Concat(elements);
        int characterIndex = joined.IndexOf(text, StringComparison.Ordinal);
        Assert.True(characterIndex >= 0, "Expected text was not present in the interaction map.");
        return characterIndex;
    }

    private static readonly byte[] TinyPng = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");

    private sealed class InlineProgress<T>(Action<T> report) : IProgress<T> {
        public void Report(T value) => report(value);
    }
}
