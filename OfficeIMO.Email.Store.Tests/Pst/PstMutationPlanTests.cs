using OfficeIMO.Email;

namespace OfficeIMO.Email.Store.Tests;

public sealed class PstMutationPlanTests {
    [Fact]
    public void DryRunCopyAndPatchesRemainReadOnlyThenCommitWithVerification() {
        string path = TemporaryPstPath();
        try {
            using (EmailStorePstWriter writer = EmailStorePstWriter.Create(path)) {
                string inbox = writer.AddFolder("Inbox", containerClass: "IPF.Note");
                var source = new EmailDocument { MessageClass = "IPM.Note", Subject = "Original" };
                source.Attachments.Add(new EmailAttachment {
                    FileName = "old.txt",
                    ContentType = "text/plain",
                    Content = Encoding.UTF8.GetBytes("old")
                });
                writer.AddItem(inbox, source);
                writer.Complete();
            }
            byte[] originalBytes = File.ReadAllBytes(path);

            using (var transaction = EmailStorePstMutationTransaction.Open(path)) {
                string originalId = Assert.Single(transaction.EnumerateItems()).Id;
                string archiveId = transaction.CreateFolder("Archive", containerClass: "IPF.Note");
                string copyId = transaction.CopyItem(originalId, archiveId);
                transaction.PatchItemProperties(originalId, new MapiPropertyPatch()
                    .Set(MapiKnownProperties.PidTag.Subject, "Patched"));
                transaction.PatchItemAttachments(originalId, new EmailAttachmentPatch()
                    .RemoveAt(0)
                    .Add(new EmailAttachment {
                        FileName = "new.bin",
                        ContentType = "application/octet-stream",
                        Content = new byte[] { 4, 5, 6 }
                    }));

                EmailStorePstMutationPlan plan = transaction.DryRun();

                Assert.True(plan.HasChanges);
                Assert.Equal(2, plan.ResultingItemCount);
                Assert.Contains(plan.Operations, operation =>
                    operation.Kind == EmailStorePstMutationOperationKind.CopyItem &&
                    operation.EntityId == copyId && operation.DestinationId == originalId);
                Assert.Contains(plan.Operations, operation =>
                    operation.Kind == EmailStorePstMutationOperationKind.PatchItem &&
                    operation.EntityId == originalId && operation.ChangeCount == 3);
                Assert.Equal(originalBytes, File.ReadAllBytes(path));

                EmailStorePstMutationReport report = transaction.Commit();

                Assert.True(report.Verification?.IsSuccessful);
                Assert.Equal(1, report.AddedItems);
                Assert.Equal(1, report.CopiedItems);
                Assert.Equal(1, report.PatchedItems);
                Assert.False(report.HasDataLoss);
            }

            using EmailStoreSession result = EmailStoreSession.Open(path);
            EmailStoreItem[] items = result.EnumerateItems()
                .Select(reference => result.ReadItem(reference))
                .OrderBy(item => item.Document.Subject, StringComparer.Ordinal)
                .ToArray();
            Assert.Equal(new[] { "Original", "Patched" },
                items.Select(item => item.Document.Subject).ToArray());
            Assert.Equal("old.txt", Assert.Single(items[0].Document.Attachments).FileName);
            EmailAttachment patched = Assert.Single(items[1].Document.Attachments);
            Assert.Equal("new.bin", patched.FileName);
            Assert.Equal(new byte[] { 4, 5, 6 }, patched.Content);
        } finally {
            TryDelete(path);
        }
    }

    [Fact]
    public void TypedPropertyAndAttachmentPatchesValidateBeforeMutation() {
        Assert.Throws<ArgumentException>(() => new MapiPropertyPatch().Set(
            MapiKnownProperties.PidTag.Subject, "value", MapiPropertyType.Integer32));
        var attachments = new List<EmailAttachment>();
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            new EmailAttachmentPatch().RemoveAt(0).Apply(attachments));
        Assert.Empty(attachments);
    }

    [Fact]
    public void AttachmentAndCombinedItemPatchesRejectTheWholeInvalidSequenceBeforeMutation() {
        var first = new EmailAttachment { FileName = "first.txt" };
        var second = new EmailAttachment { FileName = "second.txt" };
        var attachments = new List<EmailAttachment> { first, second };
        var attachmentPatch = new EmailAttachmentPatch().RemoveAt(0).RemoveAt(99);

        Assert.Throws<ArgumentOutOfRangeException>(() => attachmentPatch.Apply(attachments));
        Assert.Equal(new[] { first, second }, attachments);

        var document = new EmailDocument();
        document.Attachments.Add(first);
        var itemPatch = new EmailStoreItemPatch().SetReadState(true)
            .PatchAttachments(new EmailAttachmentPatch().RemoveAt(0).RemoveAt(99));

        Assert.Throws<ArgumentOutOfRangeException>(() => itemPatch.Apply(document));
        Assert.Null(document.MessageMetadata.IsRead);
        Assert.Same(first, Assert.Single(document.Attachments));
    }

    [Fact]
    public async System.Threading.Tasks.Task AsyncCommitFlushesVerifiesAndCopiesBackupBeforeAtomicReplacement() {
        string path = TemporaryPstPath();
        string backupPath = string.Concat(path, ".backup");
        try {
            using (EmailStorePstWriter writer = EmailStorePstWriter.Create(path)) {
                string inbox = writer.AddFolder("Inbox", containerClass: "IPF.Note");
                var source = new EmailDocument { MessageClass = "IPM.Note", Subject = "Before" };
                source.Attachments.Add(new EmailAttachment {
                    FileName = "payload.bin",
                    ContentType = "application/octet-stream",
                    Content = Enumerable.Range(0, 1024 * 1024)
                        .Select(index => (byte)(index % 251)).ToArray()
                });
                writer.AddItem(inbox, source);
                writer.Complete();
            }
            byte[] original = File.ReadAllBytes(path);

            using (var transaction = EmailStorePstMutationTransaction.Open(path,
                new EmailStorePstMutationOptions(backupPath: backupPath))) {
                string itemId = Assert.Single(transaction.EnumerateItems()).Id;
                transaction.PatchItemProperties(itemId, new MapiPropertyPatch()
                    .Set(MapiKnownProperties.PidTag.Subject, "After"));

                EmailStorePstMutationReport report = await transaction.CommitAsync();

                Assert.True(report.Verification?.IsSuccessful);
                Assert.Equal(backupPath, report.BackupPath);
                Assert.Equal(original, File.ReadAllBytes(backupPath));
            }

            using EmailStoreSession result = EmailStoreSession.Open(path);
            EmailStoreItemReference reference = Assert.Single(result.EnumerateItems());
            Assert.Equal("After", result.ReadItem(reference).Document.Subject);
        } finally {
            TryDelete(path);
            TryDelete(backupPath);
        }
    }

    [Fact]
    public async System.Threading.Tasks.Task AsyncCommitDoesNotPublishBackupWhenSourceChanged() {
        string path = TemporaryPstPath();
        string backupPath = string.Concat(path, ".backup");
        byte[] sentinel = Encoding.ASCII.GetBytes("existing-backup");
        try {
            using (EmailStorePstWriter writer = EmailStorePstWriter.Create(path)) {
                string inbox = writer.AddFolder("Inbox", containerClass: "IPF.Note");
                writer.AddItem(inbox, new EmailDocument { MessageClass = "IPM.Note", Subject = "Before" });
                writer.Complete();
            }
            File.WriteAllBytes(backupPath, sentinel);

            using (var transaction = EmailStorePstMutationTransaction.Open(path,
                new EmailStorePstMutationOptions(backupPath: backupPath, overwriteBackup: true))) {
                string itemId = Assert.Single(transaction.EnumerateItems()).Id;
                transaction.PatchItemProperties(itemId, new MapiPropertyPatch()
                    .Set(MapiKnownProperties.PidTag.Subject, "After"));
                // The source session intentionally prevents a cooperating writer from changing this file.
                // Alter the captured identity to deterministically exercise the late source-change guard.
                System.Reflection.FieldInfo sourceLength = typeof(EmailStorePstMutationTransaction)
                    .GetField("_sourceLength", System.Reflection.BindingFlags.Instance |
                        System.Reflection.BindingFlags.NonPublic)!;
                sourceLength.SetValue(transaction, new FileInfo(path).Length + 1L);

                await Assert.ThrowsAsync<IOException>(() => transaction.CommitAsync());
            }

            Assert.Equal(sentinel, File.ReadAllBytes(backupPath));
        } finally {
            TryDelete(path);
            TryDelete(backupPath);
        }
    }

    [Fact]
    public void FolderSubtreeCopyAndTypedQueryPatchCommitAsOneVerifiedRewrite() {
        string path = TemporaryPstPath();
        try {
            using (EmailStorePstWriter writer = EmailStorePstWriter.Create(path)) {
                string project = writer.AddFolder("Project", containerClass: "IPF.Note");
                string child = writer.AddFolder("Child", project, "IPF.Note");
                writer.AddItem(project, new EmailDocument { MessageClass = "IPM.Note", Subject = "Batch root" });
                writer.AddItem(child, new EmailDocument { MessageClass = "IPM.Note", Subject = "Batch child" });
                writer.Complete();
            }

            using (var transaction = EmailStorePstMutationTransaction.Open(path)) {
                EmailStorePstMutationFolder project = Assert.Single(transaction.Folders,
                    folder => folder.Name == "Project");
                EmailStorePstFolderCopyResult copy = transaction.CopyFolder(
                    project.Id, transaction.RootFolderId, includeDescendants: true,
                    conflictPolicy: EmailStorePstCopyConflictPolicy.AllowDuplicate);
                var patch = new EmailStoreItemPatch()
                    .SetReadState(true)
                    .SetImportance(EmailMessageImportance.High)
                    .SetCategories(new[] { "Automation", "Premium" });
                EmailStorePstMutationSelectionReport selection = transaction.PatchItems(
                    new EmailStoreTableQuery(filter: EmailStoreFields.Subject.StartsWith("Batch"), pageSize: 1),
                    patch);

                Assert.Equal(2, copy.FolderIdMap.Count);
                Assert.Equal(2, copy.ItemIdMap.Count);
                Assert.Equal(2, selection.PatchedItems);
                EmailStorePstMutationPlan plan = transaction.DryRun();
                Assert.Equal(4, plan.ResultingItemCount);
                Assert.Equal(2, plan.Operations.Count(operation =>
                    operation.Kind == EmailStorePstMutationOperationKind.CopyItem));
                Assert.Equal(2, plan.Operations.Count(operation =>
                    operation.Kind == EmailStorePstMutationOperationKind.PatchItem));

                EmailStorePstMutationReport report = transaction.Commit();
                Assert.True(report.Verification?.IsSuccessful);
                Assert.Equal(2, report.CopiedItems);
                Assert.Equal(2, report.PatchedItems);
            }

            using EmailStoreSession result = EmailStoreSession.Open(path);
            EmailStoreItem[] items = result.EnumerateItems()
                .Select(reference => result.ReadItem(reference)).ToArray();
            Assert.Equal(4, items.Length);
            EmailStoreItem[] originals = items.Where(item => item.Document.MessageMetadata.IsRead == true).ToArray();
            Assert.Equal(2, originals.Length);
            Assert.All(originals, item => {
                Assert.Equal(EmailMessageImportance.High, item.Document.MessageMetadata.Importance);
                Assert.Equal(new[] { "Automation", "Premium" }, item.Document.MessageMetadata.Categories);
            });
            Assert.Equal(2, result.Folders.Count(folder => folder.Name == "Project"));
            Assert.Equal(2, result.Folders.Count(folder => folder.Name == "Child"));
        } finally {
            TryDelete(path);
        }
    }

    private static string TemporaryPstPath() => Path.Combine(Path.GetTempPath(),
        string.Concat("officeimo-pst-plan-", Guid.NewGuid().ToString("N"), ".pst"));

    private static void TryDelete(string path) {
        try { if (File.Exists(path)) File.Delete(path); }
        catch (IOException) { }
        catch (UnauthorizedAccessException) { }
    }
}
