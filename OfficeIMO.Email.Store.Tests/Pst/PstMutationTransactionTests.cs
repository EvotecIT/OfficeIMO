using OfficeIMO.Email;
using System.Collections;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;

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
    public void MutationPreservesStandardSpecialFolderIdentityAndDistinctRoots() {
        string path = TemporaryPstPath();
        try {
            using (EmailStorePstWriter writer = EmailStorePstWriter.Create(path)) {
                writer.AddFolder("Inbox source", EmailStoreSpecialFolderKind.Inbox,
                    containerClass: "IPF.Note");
                writer.AddFolder("Sent source", EmailStoreSpecialFolderKind.SentItems,
                    containerClass: "IPF.Note");
                writer.AddFolder("Outbox source", EmailStoreSpecialFolderKind.Outbox,
                    containerClass: "IPF.Note");
                writer.AddFolder("Calendar source", EmailStoreSpecialFolderKind.Calendar,
                    containerClass: "IPF.Appointment");
                writer.AddFolder("Contacts source", EmailStoreSpecialFolderKind.Contacts,
                    containerClass: "IPF.Contact");
                writer.Complete();
            }

            using (var transaction = EmailStorePstMutationTransaction.Open(path)) {
                EmailStorePstMutationFolder inbox = Assert.Single(transaction.Folders,
                    folder => folder.SpecialFolderKind == EmailStoreSpecialFolderKind.Inbox);
                string rootId = Assert.Single(transaction.Folders,
                    folder => folder.SpecialFolderKind == EmailStoreSpecialFolderKind.Root).Id;
                string ipmSubtreeId = Assert.Single(transaction.Folders,
                    folder => folder.SpecialFolderKind == EmailStoreSpecialFolderKind.IpmSubtree).Id;
                transaction.RenameFolder(inbox.Id, "Renamed but still Inbox");
                EmailStorePstMutationReport report = transaction.Commit();

                Assert.True(report.Verification?.IsSuccessful);
                Assert.NotEqual(report.FolderIdMap[rootId], report.FolderIdMap[ipmSubtreeId]);
            }

            using EmailStoreSession session = EmailStoreSession.Open(path);
            foreach (EmailStoreSpecialFolderKind role in new[] {
                EmailStoreSpecialFolderKind.Inbox,
                EmailStoreSpecialFolderKind.SentItems,
                EmailStoreSpecialFolderKind.Outbox,
                EmailStoreSpecialFolderKind.Calendar,
                EmailStoreSpecialFolderKind.Contacts
            }) {
                EmailStoreFolderInfo folder = Assert.Single(session.Folders,
                    candidate => candidate.SpecialFolderKind == role);
                Assert.Equal(EmailStoreFolderClassificationSource.SourceIdentifier,
                    folder.ClassificationSource);
            }
            Assert.Equal("Renamed but still Inbox", Assert.Single(session.Folders,
                folder => folder.SpecialFolderKind == EmailStoreSpecialFolderKind.Inbox).Name);
            Assert.NotEqual(
                Assert.Single(session.Folders, folder =>
                    folder.SpecialFolderKind == EmailStoreSpecialFolderKind.Root).Id,
                Assert.Single(session.Folders, folder =>
                    folder.SpecialFolderKind == EmailStoreSpecialFolderKind.IpmSubtree).Id);
        } finally {
            TryDelete(path);
        }
    }

    [Fact]
    public void MutationPreservesMandatoryFolderDisplayMetadata() {
        string path = TemporaryPstPath();
        try {
            using (EmailStorePstWriter writer = EmailStorePstWriter.Create(path)) {
                writer.ConfigureFolderMetadata(writer.MessageStoreRootFolderId, "Korzen skrzynki", null);
                writer.ConfigureFolderMetadata(writer.RootFolderId, "Gora folderow osobistych", "IPF.Note");
                writer.ConfigureFolderMetadata(writer.DeletedItemsFolderId, "Elementy usuniete", "IPF.Note");
                writer.ConfigureFolderMetadata(writer.SearchRootFolderId, "Wyszukiwanie", null);
                writer.ConfigureFolderMetadata(writer.SpamSearchFolderId, "Wyszukiwanie spamu", "IPF.Note");
                writer.Complete();
            }

            using (var transaction = EmailStorePstMutationTransaction.Open(path)) {
                transaction.CreateFolder("Commit trigger");
                Assert.True(transaction.Commit().Verification?.IsSuccessful);
            }

            using EmailStoreSession session = EmailStoreSession.Open(path);
            Assert.Equal("Korzen skrzynki", Assert.Single(session.Folders,
                folder => folder.SpecialFolderKind == EmailStoreSpecialFolderKind.Root).Name);
            Assert.Equal("Gora folderow osobistych", Assert.Single(session.Folders,
                folder => folder.SpecialFolderKind == EmailStoreSpecialFolderKind.IpmSubtree).Name);
            Assert.Equal("Elementy usuniete", Assert.Single(session.Folders,
                folder => folder.SpecialFolderKind == EmailStoreSpecialFolderKind.DeletedItems).Name);
            Assert.Equal("Wyszukiwanie", Assert.Single(session.Folders,
                folder => folder.SpecialFolderKind == EmailStoreSpecialFolderKind.SearchRoot).Name);
            Assert.Contains(session.Folders, folder => folder.Name == "Wyszukiwanie spamu");
        } finally {
            TryDelete(path);
        }
    }

    [Fact]
    public void MutationKeepsDisplayNameDuplicateSeparateFromSourceIdentifiedSystemFolder() {
        string path = TemporaryPstPath();
        try {
            using (EmailStorePstWriter writer = EmailStorePstWriter.Create(path)) {
                string duplicate = writer.AddFolder("Deleted Items", containerClass: "IPF.Note");
                writer.AddItem(duplicate, new EmailDocument { Subject = "Keep duplicate separate" });
                writer.Complete();
            }
            string sourceSystemId;
            string sourceDuplicateId;
            using (EmailStoreSession source = EmailStoreSession.Open(path)) {
                EmailStoreFolderInfo[] matching = source.Folders.Where(folder =>
                    folder.Name == "Deleted Items").ToArray();
                sourceSystemId = Assert.Single(matching, folder =>
                    folder.ClassificationSource == EmailStoreFolderClassificationSource.SourceIdentifier).Id;
                sourceDuplicateId = Assert.Single(matching, folder =>
                    folder.ClassificationSource == EmailStoreFolderClassificationSource.DisplayName).Id;
            }

            EmailStorePstMutationReport report;
            using (var transaction = EmailStorePstMutationTransaction.Open(path)) {
                transaction.CreateFolder("Commit trigger");
                report = transaction.Commit();
            }

            Assert.NotEqual(report.FolderIdMap[sourceSystemId], report.FolderIdMap[sourceDuplicateId]);
            using EmailStoreSession rewritten = EmailStoreSession.Open(path);
            EmailStoreFolderInfo[] rewrittenMatching = rewritten.Folders.Where(folder =>
                folder.Name == "Deleted Items").ToArray();
            Assert.Equal(2, rewrittenMatching.Length);
            EmailStoreFolderInfo duplicateFolder = Assert.Single(rewrittenMatching, folder =>
                folder.ClassificationSource == EmailStoreFolderClassificationSource.DisplayName);
            Assert.Equal("Keep duplicate separate", Assert.Single(rewritten.EnumerateItems(
                new EmailStoreEnumerationOptions(folderId: duplicateFolder.Id))
                .Select(reference => rewritten.ReadSummary(reference))).Subject);
        } finally {
            TryDelete(path);
        }
    }

    [Fact]
    public void MutationDoesNotAddWriterOwnedSpamSearchFolderWhenSourceDoesNotHaveOne() {
        string path = TemporaryPstPath();
        try {
            using (EmailStorePstWriter writer = EmailStorePstWriter.Create(path)) {
                writer.SuppressWriterOwnedSpamSearchFolder();
                writer.AddFolder("Inbox", containerClass: "IPF.Note");
                writer.Complete();
            }
            using (EmailStoreSession source = EmailStoreSession.Open(path)) {
                Assert.DoesNotContain(source.Folders, folder => folder.Name == "SPAM Search Folder 2");
            }

            using (var transaction = EmailStorePstMutationTransaction.Open(path)) {
                transaction.CreateFolder("Commit trigger");
                Assert.True(transaction.Commit().Verification?.IsSuccessful);
            }

            using EmailStoreSession rewritten = EmailStoreSession.Open(path);
            Assert.DoesNotContain(rewritten.Folders, folder => folder.Name == "SPAM Search Folder 2");
        } finally {
            TryDelete(path);
        }
    }

    [Fact]
    public void FixedSearchFolderNidWithoutWriterProvenanceIsHandledConservatively() {
        string path = TemporaryPstPath();
        try {
            byte[] original = PstTestFileBuilder.Create(includeFixedNidSearchFolder: true);
            File.WriteAllBytes(path, original);
            using var transaction = EmailStorePstMutationTransaction.Open(path);
            transaction.CreateFolder("Commit trigger");

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
                transaction.Commit());

            Assert.Contains("EMAIL_STORE_PST_MUTATE_SEARCH_FOLDER_STATIC", exception.Message,
                StringComparison.Ordinal);
            Assert.Equal(original, File.ReadAllBytes(path));
        } finally {
            TryDelete(path);
        }
    }

    [Theory]
    [InlineData(0x8022U)]
    [InlineData(0xA002U)]
    public void MissingOrSelfParentFolderTriggersTheDefaultFidelityGuard(uint parentNid) {
        string path = TemporaryPstPath();
        try {
            byte[] original = PstTestFileBuilder.Create(inboxParentNid: parentNid);
            File.WriteAllBytes(path, original);
            using var transaction = EmailStorePstMutationTransaction.Open(path);
            transaction.CreateFolder("Commit trigger");

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
                transaction.Commit());

            Assert.Contains("EMAIL_STORE_PST_MUTATE_FOLDER_PARENT_RECOVERED", exception.Message,
                StringComparison.Ordinal);
            Assert.Equal(original, File.ReadAllBytes(path));
        } finally {
            TryDelete(path);
        }
    }

    [Fact]
    public void MutationPreservesAndAcceptsSurroundingWhitespaceInOrdinaryFolderNames() {
        string path = TemporaryPstPath();
        try {
            using (EmailStorePstWriter writer = EmailStorePstWriter.Create(path)) {
                writer.AddFolder("  Existing folder  ", containerClass: "IPF.Note");
                writer.Complete();
            }

            using (var transaction = EmailStorePstMutationTransaction.Open(path)) {
                string existingId = Assert.Single(transaction.Folders,
                    folder => folder.Name == "  Existing folder  ").Id;
                transaction.RenameFolder(existingId, "  Renamed folder  ");
                transaction.CreateFolder("  Created folder  ", containerClass: "IPF.Note");
                Assert.True(transaction.Commit().Verification?.IsSuccessful);
            }

            using EmailStoreSession rewritten = EmailStoreSession.Open(path);
            Assert.Single(rewritten.Folders, folder => folder.Name == "  Renamed folder  ");
            Assert.Single(rewritten.Folders, folder => folder.Name == "  Created folder  ");
        } finally {
            TryDelete(path);
        }
    }

    [Fact]
    public void DeepFolderHierarchyMappingIsLinearAndCancellationAware() {
        string sourcePath = TemporaryPstPath();
        string destinationPath = TemporaryPstPath();
        const int depth = 20000;
        try {
            using (EmailStorePstWriter sourceWriter = EmailStorePstWriter.Create(sourcePath)) {
                sourceWriter.Complete();
            }
            using var transaction = EmailStorePstMutationTransaction.Open(sourcePath,
                new EmailStorePstMutationOptions(maxFolderCount: depth + 100));
            string parent = transaction.RootFolderId;
            for (int index = 0; index < depth; index++) {
                parent = transaction.CreateFolder("Level " + index.ToString(), parent);
            }

            using var destinationWriter = EmailStorePstWriter.Create(destinationPath,
                new EmailStorePstWriterOptions(maxFolderCount: depth + 100));
            var folderMap = new Dictionary<string, string>(StringComparer.Ordinal);
            var parentMap = new Dictionary<string, string?>(StringComparer.Ordinal);
            var stopwatch = System.Diagnostics.Stopwatch.StartNew();

            transaction.BuildFolderMap(destinationWriter, folderMap, parentMap,
                CancellationToken.None);

            stopwatch.Stop();
            Assert.Equal(transaction.Folders.Count, folderMap.Count);
            Assert.True(stopwatch.Elapsed < TimeSpan.FromSeconds(10),
                "Deep hierarchy mapping took " + stopwatch.Elapsed + ".");
            using var cancelled = new CancellationTokenSource();
            cancelled.Cancel();
            Assert.Throws<OperationCanceledException>(() => transaction.BuildFolderMap(
                destinationWriter, new Dictionary<string, string>(),
                new Dictionary<string, string?>(), cancelled.Token));
        } finally {
            TryDelete(sourcePath);
            TryDelete(destinationPath);
        }
    }

    [Fact]
    public void PathIdentityUsesTheActualDirectoryCaseBehavior() {
        string directory = Path.Combine(Path.GetTempPath(),
            "officeimo-path-identity-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        string path = Path.Combine(directory, "Mailbox.pst");
        string alias = Path.Combine(directory, "MAILBOX.PST");
        try {
            File.WriteAllText(path, "identity");
            bool actualCaseInsensitive = File.Exists(alias);

            Assert.Equal(EmailStorePathIdentity.Normalize(path, caseInsensitive: true),
                EmailStorePathIdentity.Normalize(alias, caseInsensitive: true));
            Assert.NotEqual(EmailStorePathIdentity.Normalize(path, caseInsensitive: false),
                EmailStorePathIdentity.Normalize(alias, caseInsensitive: false));
            Assert.Equal(actualCaseInsensitive, EmailStorePathIdentity.IsCaseInsensitiveFileSystem(path));
            Assert.Equal(actualCaseInsensitive, EmailStorePathIdentity.AreEquivalent(path, alias));
        } finally {
            try { Directory.Delete(directory, recursive: true); }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }
    }

    [Fact]
    public void MutationLockRejectsASecondOfficeIMOTransactionForTheSamePath() {
        string path = TemporaryPstPath();
        try {
            CreateSource(path);
            using (EmailStorePstMutationTransaction first = EmailStorePstMutationTransaction.Open(path)) {
                IOException exception = Assert.Throws<IOException>(() =>
                    EmailStorePstMutationTransaction.Open(path));
                Assert.Contains("already owns", exception.Message, StringComparison.OrdinalIgnoreCase);
            }
            using EmailStorePstMutationTransaction reopened = EmailStorePstMutationTransaction.Open(path);
            Assert.NotEmpty(reopened.Folders);
        } finally {
            TryDelete(path);
        }
    }

    [Fact]
    public void WriterOwnedSearchFolderRejectsAddedAndMovedItemsBeforeCommit() {
        string path = TemporaryPstPath();
        try {
            CreateSource(path);
            using EmailStorePstMutationTransaction transaction =
                EmailStorePstMutationTransaction.Open(path);
            string searchFolder = Assert.Single(transaction.Folders,
                folder => folder.Name == "SPAM Search Folder 2").Id;
            string item = transaction.EnumerateItems().First().Id;

            Assert.Throws<InvalidOperationException>(() => transaction.AddItem(
                searchFolder, new EmailDocument { Subject = "Not searchable" }));
            Assert.Throws<InvalidOperationException>(() => transaction.MoveItem(item, searchFolder));
        } finally {
            TryDelete(path);
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
    public void MoveFolderRejectsAnExistingAncestorCycleInsteadOfLooping() {
        string path = TemporaryPstPath();
        try {
            using (EmailStorePstWriter writer = EmailStorePstWriter.Create(path)) {
                writer.AddFolder("First");
                writer.AddFolder("Second");
                writer.AddFolder("Moved");
                writer.Complete();
            }
            using EmailStorePstMutationTransaction transaction =
                EmailStorePstMutationTransaction.Open(path);
            EmailStorePstMutationFolder first = Assert.Single(transaction.Folders,
                folder => folder.Name == "First");
            EmailStorePstMutationFolder second = Assert.Single(transaction.Folders,
                folder => folder.Name == "Second");
            EmailStorePstMutationFolder moved = Assert.Single(transaction.Folders,
                folder => folder.Name == "Moved");
            FieldInfo foldersField = typeof(EmailStorePstMutationTransaction).GetField(
                "_folders", BindingFlags.Instance | BindingFlags.NonPublic)!;
            var folders = (IDictionary)foldersField.GetValue(transaction)!;
            object firstState = folders[first.Id]!;
            object secondState = folders[second.Id]!;
            PropertyInfo parentId = firstState.GetType().GetProperty(
                "ParentId", BindingFlags.Instance | BindingFlags.NonPublic)!;
            parentId.SetValue(firstState, second.Id);
            parentId.SetValue(secondState, first.Id);

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
                transaction.MoveFolder(moved.Id, first.Id));

            Assert.Contains("already contains a cycle", exception.Message,
                StringComparison.OrdinalIgnoreCase);
        } finally {
            TryDelete(path);
        }
    }

    [Fact]
    public void InvalidMandatoryFolderParentTriggersTheDefaultFidelityGuard() {
        string path = TemporaryPstPath();
        try {
            using (EmailStorePstWriter writer = EmailStorePstWriter.Create(path)) writer.Complete();
            byte[] original = File.ReadAllBytes(path);
            using var transaction = EmailStorePstMutationTransaction.Open(path);
            string ipmSubtreeId = Assert.Single(transaction.Folders,
                folder => folder.SpecialFolderKind == EmailStoreSpecialFolderKind.IpmSubtree).Id;
            FieldInfo foldersField = typeof(EmailStorePstMutationTransaction).GetField(
                "_folders", BindingFlags.Instance | BindingFlags.NonPublic)!;
            var folders = (IDictionary)foldersField.GetValue(transaction)!;
            object ipmSubtree = folders[ipmSubtreeId]!;
            PropertyInfo parentId = ipmSubtree.GetType().GetProperty(
                "ParentId", BindingFlags.Instance | BindingFlags.NonPublic)!;
            parentId.SetValue(ipmSubtree, ipmSubtreeId);

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
                transaction.Commit());

            Assert.Contains("EMAIL_STORE_PST_MUTATE_FOLDER_PARENT_RECOVERED", exception.Message,
                StringComparison.Ordinal);
            Assert.Equal(original, File.ReadAllBytes(path));
        } finally {
            TryDelete(path);
        }
    }

    [Fact]
    public void LinuxAllowsCaseDistinctBackupAndSourcePaths() {
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Linux)) return;
        string directory = Path.Combine(Path.GetTempPath(),
            "officeimo-pst-backup-case-" + Guid.NewGuid().ToString("N"));
        string source = Path.Combine(directory, "Archive.pst");
        string backup = Path.Combine(directory, "archive.pst");
        try {
            Directory.CreateDirectory(directory);
            CreateSource(source);

            using EmailStorePstMutationTransaction transaction =
                EmailStorePstMutationTransaction.Open(source,
                    new EmailStorePstMutationOptions(backupPath: backup));

            Assert.NotEmpty(transaction.Folders);
        } finally {
            if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true);
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

    [Fact]
    public void Password_protected_pst_is_rejected_when_protection_cannot_be_preserved() {
        string path = TemporaryPstPath();
        try {
            const string password = "OfficeIMO";
            uint checksum = PstPassword.ComputeChecksum(Encoding.ASCII.GetBytes(password));
            byte[] bytes = PstTestFileBuilder.Create(storePasswordChecksum: checksum);
            File.WriteAllBytes(path, bytes);

            NotSupportedException exception = Assert.Throws<NotSupportedException>(() =>
                EmailStorePstMutationTransaction.Open(path,
                    new EmailStorePstMutationOptions(pstPassword: password)));

            Assert.Contains("cannot preserve password protection", exception.Message,
                StringComparison.OrdinalIgnoreCase);
            Assert.Equal(bytes, File.ReadAllBytes(path));
        } finally {
            TryDelete(path);
        }
    }

    [Fact]
    public void Late_source_read_diagnostics_abort_default_mutation() {
        string path = TemporaryPstPath();
        try {
            byte[] bytes = PstTestFileBuilder.Create(includeEmbeddedMessage: true);
            File.WriteAllBytes(path, bytes);
            using var transaction = EmailStorePstMutationTransaction.Open(path,
                new EmailStorePstMutationOptions(maxNestedMessageDepth: 0));
            string inbox = Assert.Single(transaction.Folders, folder => folder.Name == "Inbox").Id;
            transaction.RenameFolder(inbox, "Inbox renamed");

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
                transaction.Commit());

            Assert.Contains("EMAIL_STORE_PST_EMBEDDED_DEPTH_LIMIT", exception.Message,
                StringComparison.Ordinal);
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
