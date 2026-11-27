using System.Buffers.Binary;
using System.IO.Compression;
using OfficeIMO.Drawing;
using OfficeIMO.Email.Tests;

namespace OfficeIMO.Email.Store.Tests.Olm;

public sealed class OlmSalvageTests {
    private const string MessageEntry =
        "Local/com.microsoft.__Messages/OfficeIMO native Outlook validation PST/OfficeIMO Mail/message_00000.xml";
    private const long MessageEntryOffset = 20354;
    private const long ContactsEntryOffset = 1197;
    private const long FirstAttachmentOffset = 19808;

    [Theory]
    [InlineData(null, ContactsEntryOffset)]
    [InlineData("native-outlook-no-central-directory.olm", ContactsEntryOffset)]
    [InlineData("native-outlook-deflate-no-central-directory.olm", 391L)]
    public void RecoversEveryRecordWithAttachmentsFromIntactOrDirectorylessArchives(string? damaged, long contactsOffset) {
        string source = damaged == null
            ? Path.Combine(DbxStoreReaderTests.CorpusRoot, "native-apple-20261003", "outlook", "native-outlook.olm")
            : DbxStoreReaderTests.DamagedFixture(damaged);
        DbxStoreReaderTests.WithDestination(destination => {
            EmailStoreSalvageReport report = EmailStoreSalvage.RecoverOlmToDirectory(source, destination);

            Assert.Equal(new[] { OutlookItemKind.Contact, OutlookItemKind.Appointment, OutlookItemKind.Message },
                report.Entries.Select(entry => entry.Kind!.Value));
            Assert.All(report.Entries, entry => Assert.Equal(EmailStoreSalvageStatus.Recovered, entry.Status));
            Assert.True(report.IsComplete);
            Assert.Equal(contactsOffset, report.Entries[0].SourceOffset);
            Assert.Equal(damaged != null, report.NativeDiagnostics.Any(diagnostic =>
                diagnostic.Code == "EMAIL_STORE_OLM_SALVAGE_DIRECTORY_UNREADABLE"));

            using EmailReadResult mail = new EmailDocumentReader().ReadDocument(report.Entries[2].DestinationPath!);
            Assert.Equal("OfficeIMO validation — Zażółć 日本語", mail.Document.Subject);
            Assert.Equal(new[] { 42L, 68L }, mail.Document.Attachments
                .Select(attachment => (long)attachment.Content!.Length).OrderBy(length => length));
            DbxStoreReaderTests.AssertOnlyArtifactsInside(destination, report);
        });
    }

    [Fact]
    public void SkipsATruncatedXmlEntryAndKeepsTheRecordsBeforeIt() {
        DbxStoreReaderTests.WithDestination(destination => {
            EmailStoreSalvageReport report = EmailStoreSalvage.RecoverOlmToDirectory(
                DbxStoreReaderTests.DamagedFixture("native-outlook-truncated.olm"), destination);

            Assert.Collection(report.Entries,
                contact => {
                    Assert.Equal(OutlookItemKind.Contact, contact.Kind);
                    Assert.Equal(EmailStoreSalvageStatus.Recovered, contact.Status);
                },
                appointment => {
                    Assert.Equal(OutlookItemKind.Appointment, appointment.Kind);
                    Assert.Equal(EmailStoreSalvageStatus.Recovered, appointment.Status);
                },
                mail => {
                    Assert.Equal("olm:entry:" + MessageEntry, mail.ItemId);
                    Assert.Equal(EmailStoreSalvageStatus.NotRecovered, mail.Status);
                    Assert.Equal(MessageEntryOffset, mail.SourceOffset);
                    Assert.Equal("EMAIL_STORE_OLM_SALVAGE_XML_ENTRY_UNREADABLE", mail.DamageCode);
                    Assert.Null(mail.DestinationPath);
                });
            Assert.Contains(report.NativeDiagnostics, diagnostic =>
                diagnostic.Code == "EMAIL_STORE_OLM_SALVAGE_XML_ENTRY_UNREADABLE" &&
                diagnostic.ByteOffset == MessageEntryOffset);
            Assert.True(report.HasErrors);
            string[] manifest = File.ReadAllLines(report.ManifestPath!);
            Assert.Equal(4, manifest.Length);
            Assert.StartsWith("olm:entry:" + MessageEntry + "\tNotRecovered\t\t", manifest[3], StringComparison.Ordinal);
        });
    }

    [Fact]
    public void ReportsADamagedAttachmentAsAPartialMessage() {
        DbxStoreReaderTests.WithDestination(destination => {
            EmailStoreSalvageReport report = EmailStoreSalvage.RecoverOlmToDirectory(
                DbxStoreReaderTests.DamagedFixture("native-outlook-attachment-crc.olm"), destination);

            Assert.Equal(new[] { EmailStoreSalvageStatus.Recovered, EmailStoreSalvageStatus.Recovered, EmailStoreSalvageStatus.Partial },
                report.Entries.Select(entry => entry.Status));
            EmailStoreSalvageEntry mail = report.Entries[2];
            Assert.Equal(OutlookItemKind.Message, mail.Kind);
            Assert.Equal("EMAIL_STORE_OLM_SALVAGE_ATTACHMENT_MISSING", mail.DamageCode);
            Assert.Equal(FirstAttachmentOffset, mail.DamageOffset);
            Assert.Contains(report.NativeDiagnostics, diagnostic =>
                diagnostic.Code == "EMAIL_STORE_OLM_SALVAGE_ENTRY_DAMAGED" &&
                diagnostic.ByteOffset == FirstAttachmentOffset &&
                diagnostic.Message.Contains("CRC-32", StringComparison.Ordinal));

            using EmailReadResult exported = new EmailDocumentReader().ReadDocument(mail.DestinationPath!);
            Assert.Equal(42, Assert.Single(exported.Document.Attachments).Content!.Length);
        });
    }

    [Fact]
    public void KeepsWellFormedRecordsFromAnXmlEntryWithABadChecksumAsUnverified() {
        DbxStoreReaderTests.WithDestination(destination => {
            EmailStoreSalvageReport report = EmailStoreSalvage.RecoverOlmToDirectory(
                DbxStoreReaderTests.DamagedFixture("native-outlook-contacts-crc.olm"), destination);

            Assert.Equal(new[] { EmailStoreSalvageStatus.Partial, EmailStoreSalvageStatus.Recovered, EmailStoreSalvageStatus.Recovered },
                report.Entries.Select(entry => entry.Status));
            EmailStoreSalvageEntry contact = report.Entries[0];
            Assert.Equal(OutlookItemKind.Contact, contact.Kind);
            Assert.Equal("EMAIL_STORE_OLM_SALVAGE_ENTRY_REPAIRED", contact.DamageCode);
            Assert.Equal(ContactsEntryOffset, contact.DamageOffset);
            using EmailReadResult exported = new EmailDocumentReader().ReadDocument(contact.DestinationPath!);
            Assert.Equal("officeIMO native Zażółć 日本語 Validation", exported.Document.Contact!.DisplayName);
        });
    }

    [Fact]
    public void KeepsTheRecordsThatFitTheReaderItemBound() {
        DbxStoreReaderTests.WithDestination(destination => {
            EmailStoreSalvageReport report = EmailStoreSalvage.RecoverOlmToDirectory(
                Path.Combine(DbxStoreReaderTests.CorpusRoot, "native-apple-20261003", "outlook", "native-outlook.olm"),
                destination, new EmailStoreSalvageOptions(readerOptions: new EmailStoreReaderOptions(maxItemCount: 1)));

            EmailStoreSalvageEntry contact = Assert.Single(report.Entries);
            Assert.Equal(OutlookItemKind.Contact, contact.Kind);
            Assert.Equal(EmailStoreSalvageStatus.Recovered, contact.Status);
            Assert.True(report.DiscoveryCompleted);
            Assert.True(report.StoppedAtLimit);
            Assert.Contains(report.NativeDiagnostics, diagnostic => diagnostic.Code == "EMAIL_STORE_OLM_SALVAGE_SOURCE_ITEM_LIMIT");
        });
    }

    [Fact]
    public void ChargesTheItemBoundOnlyForRecordsTheReaderProjects() {
        // A stray <email> under <contacts> is not projected, and the reader ignores a second copy of the same path,
        // so neither may use the item or decoded-byte allowance that the two real records need.
        const string contacts = "<contacts><email /><contact><OPFContactCopyDisplayName>Ada Lovelace</OPFContactCopyDisplayName></contact></contacts>";
        const string mail = "<emails><email><OPFMessageCopySubject>Kept</OPFMessageCopySubject></email></emails>";
        byte[] archive;
        using (var builder = new OlmTestArchiveBuilder()) {
            archive = builder.AddText("Local/Contacts/Contacts.xml", contacts)
                .AddText("Local/Contacts/Contacts.xml", contacts)
                .AddText("Local/com.microsoft.__Messages/Inbox/message_00000.xml", mail)
                .Build();
        }
        DbxStoreReaderTests.WithDestination(destination => {
            using var input = new MemoryStream(archive, writable: false);
            EmailStoreSalvageReport report = EmailStoreSalvage.RecoverOlmToDirectory(input, "archive.olm", destination,
                new EmailStoreSalvageOptions(readerOptions: new EmailStoreReaderOptions(maxItemCount: 2,
                    maxArchiveDecodedBytes: Encoding.UTF8.GetByteCount(contacts) + Encoding.UTF8.GetByteCount(mail))));

            Assert.Equal(new[] { OutlookItemKind.Contact, OutlookItemKind.Message },
                report.Entries.Select(entry => entry.Kind!.Value));
            Assert.All(report.Entries, entry => Assert.Equal(EmailStoreSalvageStatus.Recovered, entry.Status));
            Assert.False(report.StoppedAtLimit);
            Assert.Contains(report.NativeDiagnostics, diagnostic => diagnostic.Code == "EMAIL_STORE_OLM_DUPLICATE_ENTRY");
        });
    }

    [Fact]
    public void FindsSignaturelessZip64DescriptorsWhenTheDirectoryIsMissing() {
        byte[] source = File.ReadAllBytes(Path.Combine(DbxStoreReaderTests.CorpusRoot, "native-apple-20261003", "outlook",
            "native-outlook.olm"));
        DbxStoreReaderTests.WithDestination(destination => {
            using var input = new MemoryStream(RewriteWithSignaturelessZip64Descriptors(source), writable: false);
            EmailStoreSalvageReport report = EmailStoreSalvage.RecoverOlmToDirectory(input, "native-outlook.olm", destination);

            Assert.Equal(new[] { OutlookItemKind.Contact, OutlookItemKind.Appointment, OutlookItemKind.Message },
                report.Entries.Select(entry => entry.Kind!.Value));
            Assert.All(report.Entries, entry => Assert.Equal(EmailStoreSalvageStatus.Recovered, entry.Status));
            Assert.DoesNotContain(report.NativeDiagnostics, diagnostic =>
                diagnostic.Code == "EMAIL_STORE_OLM_SALVAGE_ENTRY_REPAIRED" ||
                diagnostic.Code == "EMAIL_STORE_OLM_SALVAGE_ENTRY_DAMAGED");
            using EmailReadResult mail = new EmailDocumentReader().ReadDocument(report.Entries[2].DestinationPath!);
            Assert.Equal(new[] { 42L, 68L }, mail.Document.Attachments
                .Select(attachment => (long)attachment.Content!.Length).OrderBy(length => length));
        });
    }

    /// <summary>
    /// Writes every entry as a stored ZIP64 entry sized only by a data descriptor without its optional signature,
    /// and leaves out the central directory, as a streaming producer that never finished the archive would.
    /// </summary>
    private static byte[] RewriteWithSignaturelessZip64Descriptors(byte[] archiveBytes) {
        using var archive = new ZipArchive(new MemoryStream(archiveBytes, writable: false), ZipArchiveMode.Read);
        using var output = new MemoryStream();
        foreach (ZipArchiveEntry entry in archive.Entries) {
            byte[] data;
            using (Stream stream = entry.Open())
            using (var buffer = new MemoryStream()) {
                stream.CopyTo(buffer);
                data = buffer.ToArray();
            }
            byte[] name = Encoding.UTF8.GetBytes(entry.FullName);
            var header = new byte[30 + name.Length + 20];
            BinaryPrimitives.WriteUInt32LittleEndian(header.AsSpan(0), 0x04034b50);
            BinaryPrimitives.WriteUInt16LittleEndian(header.AsSpan(4), 45);
            BinaryPrimitives.WriteUInt16LittleEndian(header.AsSpan(6), 0x0808);
            BinaryPrimitives.WriteUInt16LittleEndian(header.AsSpan(26), (ushort)name.Length);
            BinaryPrimitives.WriteUInt16LittleEndian(header.AsSpan(28), 20);
            name.CopyTo(header, 30);
            BinaryPrimitives.WriteUInt16LittleEndian(header.AsSpan(30 + name.Length), 1);
            BinaryPrimitives.WriteUInt16LittleEndian(header.AsSpan(32 + name.Length), 16);
            output.Write(header, 0, header.Length);
            output.Write(data, 0, data.Length);
            var descriptor = new byte[20];
            BinaryPrimitives.WriteUInt32LittleEndian(descriptor.AsSpan(0), OfficePngCrc32.Compute(data, 0, data.Length));
            BinaryPrimitives.WriteUInt64LittleEndian(descriptor.AsSpan(4), (ulong)data.Length);
            BinaryPrimitives.WriteUInt64LittleEndian(descriptor.AsSpan(12), (ulong)data.Length);
            output.Write(descriptor, 0, descriptor.Length);
        }
        return output.ToArray();
    }
}
