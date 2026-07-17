using OfficeIMO.Email;

namespace OfficeIMO.Email.Store.Tests;

public sealed class PstMutationTransactionTests {
    [Fact]
    public void Existing_pst_mutation_rewrites_verifies_backs_up_and_replaces_atomically() {
        string path = TemporaryPstPath();
        string backupPath = string.Concat(path, ".backup.pst");
        try {
            CreateSource(path);
            string inboxId;
            string moveId;
            string replaceId;
            string deleteId;
            using (var transaction = EmailStorePstMutationTransaction.Open(path,
                new EmailStorePstMutationOptions(backupPath: backupPath))) {
                inboxId = Assert.Single(transaction.Folders,
                    folder => folder.Name == "Inbox").Id;
                Dictionary<string, string> ids = transaction.EnumerateItems().ToDictionary(
                    reference => transaction.ReadItem(reference.Id).Document.Subject!,
                    reference => reference.Id, StringComparer.Ordinal);
                moveId = ids["Move me"];
                replaceId = ids["Replace me"];
                deleteId = ids["Delete me"];

                string archiveId = transaction.CreateFolder("Archive", containerClass: "IPF.Note");
                transaction.RenameFolder(inboxId, "Inbox Renamed");
                transaction.MoveItem(moveId, archiveId);
                transaction.ReplaceItem(replaceId, new EmailDocument {
                    Subject = "Replacement",
                    MessageClass = "IPM.Note"
                });
                transaction.DeleteItem(deleteId);
                string addedId = transaction.AddItem(archiveId, new EmailDocument {
                    Subject = "Added",
                    MessageClass = "IPM.Note"
                });

                EmailStorePstMutationReport report = transaction.Commit();
                Assert.Equal(Path.GetFullPath(path), report.SourcePath);
                Assert.Equal(Path.GetFullPath(backupPath), report.BackupPath);
                Assert.True(report.Verification?.IsSuccessful);
                Assert.Equal(1, report.CreatedFolders);
                Assert.Equal(1, report.RenamedFolders);
                Assert.Equal(1, report.AddedItems);
                Assert.Equal(1, report.ReplacedItems);
                Assert.Equal(1, report.MovedItems);
                Assert.Equal(1, report.DeletedItems);
                Assert.True(report.FolderIdMap.ContainsKey(archiveId));
                Assert.True(report.ItemIdMap.ContainsKey(addedId));
                Assert.False(report.ItemIdMap.ContainsKey(deleteId));
                Assert.False(report.HasDataLoss);
            }

            using (EmailStoreSession mutated = EmailStoreSession.Open(path)) {
                EmailStoreFolderInfo archive = Assert.Single(mutated.Folders,
                    folder => folder.Name == "Archive");
                Assert.Single(mutated.Folders, folder => folder.Name == "Inbox Renamed");
                EmailStoreItemReference[] references = mutated.EnumerateItems().ToArray();
                Dictionary<string, EmailStoreItemReference> bySubject = references.ToDictionary(
                    reference => mutated.ReadItem(reference).Document.Subject!,
                    reference => reference, StringComparer.Ordinal);
                Assert.Equal(new[] { "Added", "Move me", "Replacement" },
                    bySubject.Keys.OrderBy(value => value, StringComparer.Ordinal).ToArray());
                Assert.Equal(archive.Id, bySubject["Move me"].FolderId);
                Assert.Equal(archive.Id, bySubject["Added"].FolderId);
            }

            using (EmailStoreSession backup = EmailStoreSession.Open(backupPath)) {
                string[] subjects = backup.EnumerateItems()
                    .Select(reference => backup.ReadItem(reference).Document.Subject!)
                    .OrderBy(value => value, StringComparer.Ordinal).ToArray();
                Assert.Equal(new[] { "Delete me", "Move me", "Replace me" }, subjects);
                Assert.Single(backup.Folders, folder => folder.Name == "Inbox");
            }
        } finally {
            TryDelete(path);
            TryDelete(backupPath);
        }
    }

    [Fact]
    public void Disposing_uncommitted_transaction_leaves_source_bytes_unchanged() {
        string path = TemporaryPstPath();
        try {
            CreateSource(path);
            byte[] before = File.ReadAllBytes(path);
            using (EmailStorePstMutationTransaction transaction =
                EmailStorePstMutationTransaction.Open(path)) {
                string inbox = Assert.Single(transaction.Folders,
                    folder => folder.Name == "Inbox").Id;
                transaction.RenameFolder(inbox, "Not committed");
            }
            Assert.Equal(before, File.ReadAllBytes(path));
        } finally {
            TryDelete(path);
        }
    }

    [Fact]
    public void Recursive_folder_delete_is_explicit_and_folder_cycles_are_rejected() {
        string path = TemporaryPstPath();
        try {
            using (EmailStorePstWriter writer = EmailStorePstWriter.Create(path)) {
                string parent = writer.AddFolder("Parent");
                string child = writer.AddFolder("Child", parent);
                writer.AddItem(child, new EmailDocument {
                    Subject = "Nested",
                    MessageClass = "IPM.Note"
                });
                writer.Complete();
            }

            using (var transaction = EmailStorePstMutationTransaction.Open(path)) {
                EmailStorePstMutationFolder parent = Assert.Single(transaction.Folders,
                    folder => folder.Name == "Parent");
                EmailStorePstMutationFolder child = Assert.Single(transaction.Folders,
                    folder => folder.Name == "Child");
                Assert.Throws<InvalidOperationException>(() =>
                    transaction.MoveFolder(parent.Id, child.Id));
                Assert.Throws<InvalidOperationException>(() =>
                    transaction.DeleteFolder(parent.Id));
                transaction.DeleteFolder(parent.Id, recursive: true);
                EmailStorePstMutationReport report = transaction.Commit();
                Assert.Equal(2, report.DeletedFolders);
                Assert.Equal(1, report.DeletedItems);
                Assert.True(report.Verification?.IsSuccessful);
            }

            using EmailStoreSession result = EmailStoreSession.Open(path);
            Assert.DoesNotContain(result.Folders, folder =>
                folder.Name == "Parent" || folder.Name == "Child");
            Assert.Empty(result.EnumerateItems());
        } finally {
            TryDelete(path);
        }
    }

    [Fact]
    public void Commit_without_an_effective_mutation_is_rejected_without_rewriting() {
        string path = TemporaryPstPath();
        try {
            CreateSource(path);
            byte[] before = File.ReadAllBytes(path);
            using (EmailStorePstMutationTransaction transaction =
                EmailStorePstMutationTransaction.Open(path)) {
                Assert.Throws<InvalidOperationException>(() => transaction.Commit());
            }
            Assert.Equal(before, File.ReadAllBytes(path));
        } finally {
            TryDelete(path);
        }
    }

    [Fact]
    public void Ansi_pst_is_rejected_instead_of_silently_converted_during_mutation() {
        string path = TemporaryPstPath();
        try {
            byte[] bytes = PstTestFileBuilder.Create(ansi: true);
            File.WriteAllBytes(path, bytes);
            Assert.Throws<NotSupportedException>(() =>
                EmailStorePstMutationTransaction.Open(path));
            Assert.Equal(bytes, File.ReadAllBytes(path));
        } finally {
            TryDelete(path);
        }
    }

    private static void CreateSource(string path) {
        using var writer = EmailStorePstWriter.Create(path,
            new EmailStorePstWriterOptions("Mutation source"));
        string inbox = writer.AddFolder("Inbox", containerClass: "IPF.Note");
        foreach (string subject in new[] { "Move me", "Replace me", "Delete me" }) {
            writer.AddItem(inbox, new EmailDocument {
                Subject = subject,
                MessageClass = "IPM.Note"
            });
        }
        writer.Complete();
    }

    private static string TemporaryPstPath() => Path.Combine(Path.GetTempPath(),
        string.Concat("officeimo-pst-mutation-", Guid.NewGuid().ToString("N"), ".pst"));

    private static void TryDelete(string path) {
        try { if (File.Exists(path)) File.Delete(path); }
        catch (IOException) { }
        catch (UnauthorizedAccessException) { }
    }
}
