using OfficeIMO.Email;

namespace OfficeIMO.Email.Store.Tests;

public sealed class PstWriterTests {
    [Fact]
    public void Empty_unicode_pst_reopens_with_mandatory_store_and_folder_structure() {
        string path = TemporaryPstPath();
        try {
            EmailStorePstWriteReport report;
            using (EmailStorePstWriter writer = EmailStorePstWriter.Create(path,
                new EmailStorePstWriterOptions("Synthetic Store"))) {
                report = writer.Complete();
            }

            Assert.True(File.Exists(path));
            Assert.True(report.BytesWritten >= 0x4800);
            Assert.Equal(0, report.ItemCount);
            AssertMaterializedNameToIdStreams(path);
            using EmailStoreSession session = EmailStoreSession.Open(path);
            Assert.Equal(EmailStoreFormat.Pst, session.Format);
            Assert.Equal("Synthetic Store", session.DisplayName);
            Assert.Contains(session.Folders, item => item.Name == "Top of Personal Folders");
            Assert.Contains(session.Folders, item => item.Name == "Deleted Items");
            Assert.Empty(session.EnumerateItems());
            Assert.DoesNotContain(session.Diagnostics,
                item => item.Severity == EmailStoreDiagnosticSeverity.Error);
            EmailStoreValidationReport validation = session.Validate(
                new EmailStoreValidationOptions(
                    mode: EmailStoreValidationMode.Shallow,
                    verifyStructuralIntegrity: true,
                    maxStructuralPages: 10_000,
                    maxStructuralBlocks: 10_000,
                    maxStructuralBytes: 128 * 1024 * 1024));
            Assert.True(validation.StructuralFailures == 0,
                string.Join(" | ", validation.Diagnostics.Select(item => item.Code + ":" + item.Message)));
        } finally {
            TryDelete(path);
        }
    }

    [Fact]
    public void Unmapped_named_property_is_preserved_under_a_diagnostic_placeholder_mapping() {
        string path = TemporaryPstPath();
        try {
            var document = new EmailDocument { Subject = "Unknown named property" };
            document.MapiProperties.Add(new MapiProperty(0x8000,
                MapiPropertyType.Unicode, "preserved-value"));
            EmailStorePstWriteReport report;
            using (EmailStorePstWriter writer = EmailStorePstWriter.Create(path)) {
                string folder = writer.AddFolder("Mailbox");
                writer.AddItem(folder, document);
                report = writer.Complete();
            }

            Assert.Contains(report.Diagnostics, diagnostic =>
                diagnostic.Code == "EMAIL_STORE_PST_WRITE_NAMED_PROPERTY_PLACEHOLDER" &&
                diagnostic.Severity == EmailStoreDiagnosticSeverity.Warning);
            using EmailStoreSession session = EmailStoreSession.Open(path);
            EmailStoreItem item = session.ReadItem(Assert.Single(session.EnumerateItems()));
            MapiProperty property = Assert.Single(item.Document.MapiProperties,
                candidate => candidate.Name?.PropertySet ==
                    new Guid("E962B602-9F1E-4F76-BC29-4795CD1752F7") &&
                    candidate.Name.LocalId == 0x8000);
            Assert.Equal("preserved-value", property.Value);
        } finally {
            TryDelete(path);
        }
    }

    [Fact]
    public void Message_recipient_attachment_named_and_multi_value_properties_round_trip() {
        string path = TemporaryPstPath();
        try {
            var document = new EmailDocument {
                Subject = "Round-trip subject",
                MessageClass = "IPM.Note",
                From = new EmailAddress("sender@example.test", "Sender"),
                Date = new DateTimeOffset(2025, 1, 2, 3, 4, 5, TimeSpan.Zero)
            };
            document.Body.Text = "Plain body";
            document.Body.Html = "<p>HTML body</p>";
            document.Recipients.Add(new EmailRecipient(EmailRecipientKind.To,
                new EmailAddress("recipient@example.test", "Recipient")));
            document.Attachments.Add(new EmailAttachment {
                FileName = "payload.bin",
                ContentType = "application/octet-stream",
                Content = new byte[] { 1, 2, 3, 4, 5 },
                Length = 5
            });
            var named = new MapiNamedProperty(
                new Guid("00020329-0000-0000-C000-000000000046"), "OfficeIMO-Test");
            document.MapiProperties.Add(new MapiProperty(0x8000,
                MapiPropertyType.Unicode, "named-value", name: named));
            document.MapiProperties.Add(new MapiProperty(0x7001,
                MapiPropertyType.MultipleUnicode, new[] { "one", "two" }));

            string folderId;
            using (EmailStorePstWriter writer = EmailStorePstWriter.Create(path)) {
                folderId = writer.AddFolder("Inbox", containerClass: "IPF.Note");
                writer.AddItem(folderId, document);
                writer.Complete();
            }

            using EmailStoreSession session = EmailStoreSession.Open(path);
            EmailStoreFolderInfo folder = Assert.Single(session.Folders, item => item.Name == "Inbox");
            EmailStoreItemReference reference = Assert.Single(session.EnumerateItems(
                new EmailStoreEnumerationOptions(folderId: folder.Id)));
            EmailStoreItem item = session.ReadItem(reference);
            Assert.Equal(document.Subject, item.Document.Subject);
            Assert.Equal(document.Body.Text, item.Document.Body.Text);
            Assert.Equal(document.Body.Html, item.Document.Body.Html);
            Assert.Equal("recipient@example.test", Assert.Single(item.Document.Recipients).Address.Address);
            EmailAttachment attachment = Assert.Single(item.Document.Attachments);
            Assert.Equal("payload.bin", attachment.FileName);
            Assert.Equal(new byte[] { 1, 2, 3, 4, 5 }, attachment.Content);
            MapiProperty mapped = Assert.Single(item.Document.MapiProperties,
                property => property.Name?.Name == "OfficeIMO-Test");
            Assert.Equal("named-value", mapped.Value);
            Assert.Equal(new[] { "one", "two" }, Assert.IsType<string[]>(
                Assert.Single(item.Document.MapiProperties,
                    property => property.PropertyId == 0x7001).Value));
            Assert.DoesNotContain(session.Diagnostics,
                diagnostic => diagnostic.Severity == EmailStoreDiagnosticSeverity.Error);
        } finally {
            TryDelete(path);
        }
    }

    [Fact]
    public void Ost_session_converts_to_a_different_new_pst_without_mutating_source() {
        byte[] sourceBytes = PstTestFileBuilder.Create(ost: true,
            attachmentContent: new byte[] { 9, 8, 7 });
        using var source = new MemoryStream(sourceBytes, writable: false);
        string path = TemporaryPstPath();
        try {
            using (EmailStoreSession sourceSession = EmailStoreSession.Open(
                source, "mailbox.ost", leaveOpen: true)) {
                EmailStorePstConversionReport report = sourceSession.ExportToPst(path);
                Assert.Equal(EmailStoreFormat.Ost, report.SourceFormat);
                Assert.Equal(1, report.ConvertedItems);
                Assert.Equal(0, report.SkippedItems);
                Assert.NotNull(report.Verification);
                Assert.True(report.Verification.IsSuccessful,
                    string.Join(" | ", report.Verification.Issues.SelectMany(issue =>
                        issue.Differences.Select(difference => string.Concat(
                            difference.Kind.ToString(), ":", difference.Path,
                            "[", difference.SourceLength?.ToString() ?? "null", ",",
                            difference.DestinationLength?.ToString() ?? "null", "]")))));
                Assert.Equal(1, report.Verification.AttemptedItems);
                Assert.Equal(1, report.Verification.MatchedItems);
                Assert.Empty(report.Verification.Issues);
                Assert.Equal(Path.GetFullPath(path), report.WriteReport.DestinationPath);
            }
            Assert.Equal(sourceBytes, source.ToArray());
            Assert.DoesNotContain(Directory.EnumerateFiles(Path.GetDirectoryName(path)!), file =>
                Path.GetFileName(file).StartsWith(
                    string.Concat(".", Path.GetFileNameWithoutExtension(path), "."),
                    StringComparison.Ordinal) &&
                string.Equals(Path.GetExtension(file), ".pst", StringComparison.OrdinalIgnoreCase));

            using EmailStoreSession converted = EmailStoreSession.Open(path);
            EmailStoreItemReference reference = Assert.Single(converted.EnumerateItems());
            EmailStoreItem item = converted.ReadItem(reference);
            Assert.Equal("Synthetic PST message", item.Document.Subject);
            Assert.Equal(new byte[] { 9, 8, 7 }, Assert.Single(item.Document.Attachments).Content);
        } finally {
            TryDelete(path);
        }
    }

    [Fact]
    public void Conversion_manifest_contains_keyed_proof_without_private_message_values() {
        const string privateSubject = "private-subject-do-not-persist";
        const string privateAddress = "private-address@example.test";
        const string privateFileName = "private-filename.bin";
        string sourcePath = TemporaryPstPath();
        string destinationPath = TemporaryPstPath();
        string manifestPath = string.Concat(destinationPath, ".verification.tsv");
        try {
            using (EmailStorePstWriter writer = EmailStorePstWriter.Create(sourcePath)) {
                string folder = writer.AddFolder("Mailbox");
                var document = new EmailDocument {
                    Subject = privateSubject,
                    From = new EmailAddress(privateAddress),
                    MessageClass = "IPM.Note"
                };
                document.Attachments.Add(new EmailAttachment {
                    FileName = privateFileName,
                    Content = new byte[] { 1, 2, 3, 4 },
                    Length = 4
                });
                writer.AddItem(folder, document);
                writer.Complete();
            }

            using (EmailStoreSession source = EmailStoreSession.Open(sourcePath)) {
                EmailStorePstConversionReport report = source.ExportToPst(destinationPath,
                    new EmailStorePstConversionOptions(
                        verificationManifestPath: manifestPath));
                Assert.NotNull(report.Verification);
                Assert.True(report.Verification.IsSuccessful);
                Assert.Equal(Path.GetFullPath(manifestPath), report.Verification.ManifestPath);
            }

            string manifest = File.ReadAllText(manifestPath);
            Assert.Contains("digest_algorithm\tHMAC-SHA-256", manifest, StringComparison.Ordinal);
            Assert.Contains("\tMATCH\t", manifest, StringComparison.Ordinal);
            Assert.Contains("summary\t1\t1\t0\t0\t", manifest, StringComparison.Ordinal);
            Assert.DoesNotContain(privateSubject, manifest, StringComparison.Ordinal);
            Assert.DoesNotContain(privateAddress, manifest, StringComparison.Ordinal);
            Assert.DoesNotContain(privateFileName, manifest, StringComparison.Ordinal);
            Assert.DoesNotContain(Directory.EnumerateFiles(Path.GetDirectoryName(manifestPath)!), file =>
                Path.GetFileName(file).StartsWith(
                    string.Concat(".", Path.GetFileName(manifestPath), "."),
                    StringComparison.Ordinal));
        } finally {
            TryDelete(sourcePath);
            TryDelete(destinationPath);
            TryDelete(manifestPath);
        }
    }

    [Fact]
    public void Failed_semantic_verification_preserves_existing_destination_and_manifest() {
        string directory = Path.Combine(Path.GetTempPath(),
            string.Concat("officeimo-pst-verified-commit-", Guid.NewGuid().ToString("N")));
        Directory.CreateDirectory(directory);
        string sourcePath = Path.Combine(directory, "source.pst");
        string destinationPath = Path.Combine(directory, "destination.pst");
        string manifestPath = Path.Combine(directory, "verification.tsv");
        byte[] originalDestination = Encoding.ASCII.GetBytes("existing-destination");
        const string originalManifest = "existing-manifest";
        try {
            using (EmailStorePstWriter writer = EmailStorePstWriter.Create(sourcePath)) {
                string folder = writer.AddFolder("Mailbox");
                var document = new EmailDocument { Subject = "verification failure" };
                document.Attachments.Add(new EmailAttachment {
                    FileName = "payload.bin",
                    Content = new byte[] { 1, 2, 3, 4 },
                    Length = 4
                });
                writer.AddItem(folder, document);
                writer.Complete();
            }
            File.WriteAllBytes(destinationPath, originalDestination);
            File.WriteAllText(manifestPath, originalManifest);
            byte[] key = Enumerable.Range(1, 32).Select(value => checked((byte)value)).ToArray();

            using EmailStoreSession source = EmailStoreSession.Open(sourcePath);
            Assert.Throws<InvalidOperationException>(() => source.ExportToPst(destinationPath,
                new EmailStorePstConversionOptions(
                    overwriteExisting: true,
                    failOnDataLoss: true,
                    verificationOptions: new EmailSemanticComparisonOptions(
                        digestKey: key, maxAttachmentBytes: 1),
                    verificationManifestPath: manifestPath)));

            Assert.Equal(originalDestination, File.ReadAllBytes(destinationPath));
            Assert.Equal(originalManifest, File.ReadAllText(manifestPath));
            Assert.Equal(new[] { sourcePath, destinationPath, manifestPath }
                    .OrderBy(value => value, StringComparer.Ordinal),
                Directory.EnumerateFiles(directory).OrderBy(value => value, StringComparer.Ordinal));
        } finally {
            try { if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true); }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }
    }

    [Fact]
    public void Checkpoint_resume_continues_without_duplicates_and_cleans_working_files() {
        string directory = Path.Combine(Path.GetTempPath(),
            string.Concat("officeimo-pst-resume-", Guid.NewGuid().ToString("N")));
        Directory.CreateDirectory(directory);
        string path = Path.Combine(directory, "resumed.pst");
        string checkpoint = Path.Combine(directory, "resumed.checkpoint");
        string folderId;
        var stages = new List<EmailStorePstWriteStage>();
        var progress = new InlineProgress<EmailStorePstWriteProgress>(value => stages.Add(value.Stage));
        try {
            using (EmailStorePstWriter writer = EmailStorePstWriter.Create(path,
                new EmailStorePstWriterOptions(checkpointPath: checkpoint,
                    checkpointIntervalItems: 1, progress: progress))) {
                folderId = writer.AddFolder("Inbox");
                writer.AddItem(folderId, new EmailDocument { Subject = "first" });
                Assert.True(File.Exists(checkpoint));
            }
            Assert.False(File.Exists(path));

            using (EmailStorePstWriter resumed = EmailStorePstWriter.Resume(checkpoint, progress)) {
                resumed.AddItem(folderId, new EmailDocument { Subject = "second" });
                EmailStorePstWriteReport report = resumed.Complete();
                Assert.Equal(2, report.ItemCount);
            }

            Assert.True(File.Exists(path));
            Assert.False(File.Exists(checkpoint));
            Assert.DoesNotContain(Directory.EnumerateFiles(directory), file =>
                !string.Equals(file, path, StringComparison.OrdinalIgnoreCase));
            using EmailStoreSession session = EmailStoreSession.Open(path);
            string[] subjects = session.EnumerateItems()
                .Select(reference => session.ReadItem(reference).Document.Subject ?? string.Empty)
                .OrderBy(value => value, StringComparer.Ordinal).ToArray();
            Assert.Equal(new[] { "first", "second" }, subjects);
            Assert.Contains(EmailStorePstWriteStage.Checkpointing, stages);
            Assert.Contains(EmailStorePstWriteStage.Completed, stages);
        } finally {
            try { if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true); }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }
    }

    [Fact]
    public void Delete_checkpoint_removes_only_its_writer_owned_crash_artifacts() {
        string directory = Path.Combine(Path.GetTempPath(),
            string.Concat("officeimo-pst-abandon-", Guid.NewGuid().ToString("N")));
        Directory.CreateDirectory(directory);
        string path = Path.Combine(directory, "abandoned.pst");
        string checkpoint = Path.Combine(directory, "abandoned.checkpoint");
        string unrelated = Path.Combine(directory, "keep.txt");
        string? similarName = null;
        try {
            File.WriteAllText(unrelated, "keep");
            using (EmailStorePstWriter writer = EmailStorePstWriter.Create(path,
                new EmailStorePstWriterOptions(checkpointPath: checkpoint))) {
                string folder = writer.AddFolder("Mailbox");
                writer.AddItem(folder, new EmailDocument { Subject = "checkpointed" });
                writer.Checkpoint();
            }
            Assert.True(File.Exists(checkpoint));
            string workingFile = Assert.Single(Directory.EnumerateFiles(directory), file =>
                Path.GetFileName(file).EndsWith(".tmp", StringComparison.Ordinal));
            similarName = string.Concat(workingFile, ".notes");
            File.WriteAllText(similarName, "keep-similar");
            string[] tableArtifacts = {
                string.Concat(workingFile, ".table-matrix.", Guid.NewGuid().ToString("N")),
                string.Concat(workingFile, ".table-row-index.", Guid.NewGuid().ToString("N")),
                string.Concat(workingFile, ".table-subnodes.", Guid.NewGuid().ToString("N"))
            };
            foreach (string artifact in tableArtifacts) File.WriteAllText(artifact, "remove");
            string malformedTableArtifact = string.Concat(workingFile, ".table-matrix.not-a-guid");
            File.WriteAllText(malformedTableArtifact, "keep-malformed");
            string checkpointCommitArtifact = Path.Combine(directory, string.Concat(
                ".", Path.GetFileName(checkpoint), ".", Guid.NewGuid().ToString("N"), ".tmp"));
            File.WriteAllText(checkpointCommitArtifact, "remove");
            string malformedCheckpointArtifact = Path.Combine(directory, string.Concat(
                ".", Path.GetFileName(checkpoint), ".not-a-guid.tmp"));
            File.WriteAllText(malformedCheckpointArtifact, "keep-malformed");

            EmailStorePstWriter.DeleteCheckpoint(checkpoint);

            Assert.False(File.Exists(checkpoint));
            Assert.All(tableArtifacts, artifact => Assert.False(File.Exists(artifact)));
            Assert.False(File.Exists(checkpointCommitArtifact));
            Assert.True(File.Exists(unrelated));
            Assert.True(File.Exists(similarName));
            Assert.True(File.Exists(malformedTableArtifact));
            Assert.True(File.Exists(malformedCheckpointArtifact));
            Assert.Equal(new[] { unrelated, similarName, malformedTableArtifact,
                    malformedCheckpointArtifact }.OrderBy(value => value, StringComparer.Ordinal),
                Directory.EnumerateFiles(directory).OrderBy(value => value, StringComparer.Ordinal));
        } finally {
            try { if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true); }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }
    }

    [Fact]
    public void Checkpoint_path_cannot_replace_the_destination() {
        string path = TemporaryPstPath();
        try {
            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
                EmailStorePstWriter.Create(path,
                    new EmailStorePstWriterOptions(checkpointPath: path)));

            Assert.Contains("different paths", exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.False(File.Exists(path));
        } finally {
            TryDelete(path);
        }
    }

    [Fact]
    public void Verification_manifest_paths_are_preflighted_before_destination_creation() {
        string directory = Path.Combine(Path.GetTempPath(),
            string.Concat("officeimo-pst-manifest-preflight-", Guid.NewGuid().ToString("N")));
        Directory.CreateDirectory(directory);
        string sourcePath = Path.Combine(directory, "source.pst");
        string destinationPath = Path.Combine(directory, "destination.pst");
        string existingManifest = Path.Combine(directory, "existing.tsv");
        try {
            using (EmailStorePstWriter writer = EmailStorePstWriter.Create(sourcePath)) {
                string folder = writer.AddFolder("Inbox");
                writer.AddItem(folder, new EmailDocument { Subject = "preflight" });
                writer.Complete();
            }
            File.WriteAllText(existingManifest, "keep");

            Assert.Throws<InvalidOperationException>(() => EmailStoreConverter.ConvertToPst(
                sourcePath, destinationPath, conversionOptions: new EmailStorePstConversionOptions(
                    verificationManifestPath: destinationPath)));
            Assert.False(File.Exists(destinationPath));

            Assert.Throws<InvalidOperationException>(() => EmailStoreConverter.ConvertToPst(
                sourcePath, destinationPath, conversionOptions: new EmailStorePstConversionOptions(
                    verificationManifestPath: sourcePath)));
            Assert.False(File.Exists(destinationPath));

            Assert.Throws<IOException>(() => EmailStoreConverter.ConvertToPst(
                sourcePath, destinationPath, conversionOptions: new EmailStorePstConversionOptions(
                    verificationManifestPath: existingManifest)));
            Assert.False(File.Exists(destinationPath));
            Assert.Equal("keep", File.ReadAllText(existingManifest));
        } finally {
            try { if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true); }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }
    }

    [Fact]
    public void Verification_manifest_requires_post_write_verification() {
        ArgumentException exception = Assert.Throws<ArgumentException>(() =>
            new EmailStorePstConversionOptions(
                verifyAfterWrite: false,
                verificationManifestPath: "verification.tsv"));

        Assert.Equal("verificationManifestPath", exception.ParamName);
    }

    [Fact]
    public void Large_data_tree_rtf_embedded_message_and_associated_item_round_trip() {
        string path = TemporaryPstPath();
        try {
            byte[] payload = Enumerable.Range(0, 100_000)
                .Select(index => unchecked((byte)index)).ToArray();
            var embedded = new EmailDocument {
                Subject = "Embedded subject",
                MessageClass = "IPM.Note"
            };
            embedded.Body.Text = "Embedded body";
            var document = new EmailDocument {
                Subject = "Parent subject",
                MessageClass = "IPM.Note"
            };
            document.Body.Rtf = "{\\rtf1\\ansi Parent RTF body}";
            document.Attachments.Add(new EmailAttachment {
                FileName = "large.bin",
                Content = payload,
                Length = payload.Length
            });
            document.Attachments.Add(new EmailAttachment {
                FileName = "embedded.msg",
                EmbeddedDocument = embedded,
                MapiAttachMethod = 5
            });

            string folderId;
            using (EmailStorePstWriter writer = EmailStorePstWriter.Create(path)) {
                folderId = writer.AddFolder("Mailbox");
                writer.AddItem(folderId, document);
                writer.AddItem(folderId, new EmailDocument {
                    Subject = "Associated configuration",
                    MessageClass = "IPM.Configuration.Test"
                }, isAssociated: true);
                writer.Complete();
            }

            using EmailStoreSession session = EmailStoreSession.Open(path);
            EmailStoreFolderInfo folder = Assert.Single(session.Folders, item => item.Name == "Mailbox");
            EmailStoreItemReference visible = Assert.Single(session.EnumerateItems(
                new EmailStoreEnumerationOptions(folderId: folder.Id)));
            EmailStoreItem item = session.ReadItem(visible);
            Assert.Equal(document.Body.Rtf, item.Document.Body.Rtf);
            Assert.Equal(payload, item.Document.Attachments.Single(
                attachment => attachment.FileName == "large.bin").Content);
            Assert.Equal("Embedded subject", item.Document.Attachments.Single(
                attachment => attachment.FileName == "embedded.msg").EmbeddedDocument?.Subject);
            EmailStoreItemReference[] all = session.EnumerateItems(
                new EmailStoreEnumerationOptions(folderId: folder.Id,
                    includeAssociatedItems: true)).ToArray();
            Assert.Equal(2, all.Length);
            Assert.Single(all, reference => reference.IsAssociated);
        } finally {
            TryDelete(path);
        }
    }

    private static string TemporaryPstPath() => Path.Combine(Path.GetTempPath(),
        string.Concat("officeimo-email-store-", Guid.NewGuid().ToString("N"), ".pst"));

    private static void AssertMaterializedNameToIdStreams(string path) {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        PstHeader header = PstHeader.Read(stream, EmailStoreFormat.Pst);
        var ndb = new PstNdbReader(stream, header, EmailStoreReaderOptions.Default, default);
        Assert.True(ndb.TryGetNode(0x61, out PstNodeReference? node));
        Assert.NotNull(node);
        var heap = new PstHeap(ndb.OpenDataTree(node!.DataBid, 16 * 1024 * 1024),
            ndb.ReadSubnodes(node.SubnodeBid), ndb, EmailStoreReaderOptions.Default, default);
        MapiProperty[] properties = new PstPropertyContextReader(heap,
            EmailStoreReaderOptions.Default, default).ReadProperties().ToArray();
        foreach (ushort propertyId in new ushort[] { 0x0002, 0x0003, 0x0004 }) {
            Assert.NotEmpty(Assert.IsType<byte[]>(Assert.Single(properties,
                property => property.PropertyId == propertyId).Value));
        }
    }

    private static void TryDelete(string path) {
        try { if (File.Exists(path)) File.Delete(path); }
        catch (IOException) { }
        catch (UnauthorizedAccessException) { }
    }

    private sealed class InlineProgress<T> : IProgress<T> {
        private readonly Action<T> _report;
        internal InlineProgress(Action<T> report) { _report = report; }
        public void Report(T value) => _report(value);
    }
}
