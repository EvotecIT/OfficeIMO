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
            }
            Assert.Equal(sourceBytes, source.ToArray());

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
}
