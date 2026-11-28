using System.Globalization;
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
            Assert.StartsWith("olm:entry:" + MessageEntry + "\tNotRecovered\t\t", manifest[3], StringComparison.Ordinal);
            Assert.All(manifest.Skip(4), line => Assert.StartsWith("source:", line, StringComparison.Ordinal));
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
    public void RecoversALaterReadableCopyWhenTheFirstIsNotWellFormed() {
        // The first copy has a valid checksum but is cut off mid-record; the second copy compresses well, so the
        // uncompressed rebuilt archive is larger than the source-length bound that admits the source.
        string name = new string('a', 4000);
        string contacts = "<contacts><contact><OPFContactCopyDisplayName>" + name +
            "</OPFContactCopyDisplayName></contact></contacts>";
        byte[] archive;
        using (var builder = new OlmTestArchiveBuilder()) {
            archive = builder.AddText("Local/Contacts/Contacts.xml", "<contacts><contact>")
                .AddText("Local/Contacts/Contacts.xml", contacts)
                .Build();
        }
        DbxStoreReaderTests.WithDestination(destination => {
            using var input = new MemoryStream(archive, writable: false);
            EmailStoreSalvageReport report = EmailStoreSalvage.RecoverOlmToDirectory(input, "archive.olm", destination,
                new EmailStoreSalvageOptions(readerOptions: new EmailStoreReaderOptions(maxInputBytes: archive.Length)));

            Assert.Collection(report.Entries,
                unreadable => {
                    Assert.Equal("olm:entry:Local/Contacts/Contacts.xml", unreadable.ItemId);
                    Assert.Equal(EmailStoreSalvageStatus.NotRecovered, unreadable.Status);
                    Assert.Equal(0, unreadable.SourceOffset);
                },
                contact => {
                    Assert.Equal(OutlookItemKind.Contact, contact.Kind);
                    Assert.Equal(EmailStoreSalvageStatus.Recovered, contact.Status);
                    Assert.True(contact.SourceOffset > 0);
                });
            Assert.True(report.DiscoveryCompleted);
        });
    }

    private const string TwoContacts = "<contacts><contact><OPFContactCopyDisplayName>First</OPFContactCopyDisplayName></contact>" +
        "<contact><OPFContactCopyDisplayName>Second</OPFContactCopyDisplayName></contact></contacts>";

    [Fact]
    public void KeepsTheFirstRecordWhenTheEntryUsesTheWholeDecodedBudget() {
        byte[] archive;
        using (var builder = new OlmTestArchiveBuilder()) {
            archive = builder.AddText("Local/Contacts/Contacts.xml", TwoContacts).Build();
        }
        DbxStoreReaderTests.WithDestination(destination => {
            using var input = new MemoryStream(archive, writable: false);
            EmailStoreSalvageReport report = EmailStoreSalvage.RecoverOlmToDirectory(input, "archive.olm", destination,
                new EmailStoreSalvageOptions(readerOptions: new EmailStoreReaderOptions(maxItemCount: 1,
                    maxArchiveDecodedBytes: Encoding.UTF8.GetByteCount(TwoContacts))));

            EmailStoreSalvageEntry contact = Assert.Single(report.Entries);
            Assert.Equal(EmailStoreSalvageStatus.Recovered, contact.Status);
            Assert.True(report.StoppedAtLimit);
            using EmailReadResult exported = new EmailDocumentReader().ReadDocument(contact.DestinationPath!);
            Assert.Equal("First", exported.Document.Contact!.DisplayName);
        });
    }

    [Fact]
    public void OpensARepairedEntryThatReserializesLongerThanItsSource() {
        // Each <contact/> is written back as <contact />, so the repaired entry is longer than the source entry
        // and than MaxArchiveEntryBytes; the trailing text after the root forces the repair.
        string source = "<contacts>" + string.Concat(Enumerable.Repeat("<contact/>", 10)) + "</contacts>x";
        byte[] archive;
        using (var builder = new OlmTestArchiveBuilder()) {
            archive = builder.AddText("Local/Contacts/Contacts.xml", source).Build();
        }
        DbxStoreReaderTests.WithDestination(destination => {
            using var input = new MemoryStream(archive, writable: false);
            EmailStoreSalvageReport report = EmailStoreSalvage.RecoverOlmToDirectory(input, "archive.olm", destination,
                new EmailStoreSalvageOptions(readerOptions: new EmailStoreReaderOptions(
                    maxArchiveEntryBytes: Encoding.UTF8.GetByteCount(source))));

            Assert.Equal(10, report.Entries.Count);
            Assert.All(report.Entries, entry => Assert.Equal(EmailStoreSalvageStatus.Partial, entry.Status));
            Assert.True(report.DiscoveryCompleted);
        });
    }

    [Fact]
    public void ChargesFailedXmlRepairsToTheDecodedBudget() {
        // Each broken entry costs its verification and its failed repair read; the good entry fits only if the
        // repair reads were free.
        const string broken = "<contacts><contact>";
        byte[] archive;
        using (var builder = new OlmTestArchiveBuilder()) {
            archive = builder.AddText("Local/A/Contacts.xml", broken)
                .AddText("Local/B/Contacts.xml", broken)
                .AddText("Local/C/Contacts.xml", TwoContacts)
                .Build();
        }
        long budget = 2L * Encoding.UTF8.GetByteCount(broken) + Encoding.UTF8.GetByteCount(TwoContacts) + 1;
        DbxStoreReaderTests.WithDestination(destination => {
            using var input = new MemoryStream(archive, writable: false);
            EmailStoreSalvageReport report = EmailStoreSalvage.RecoverOlmToDirectory(input, "archive.olm", destination,
                new EmailStoreSalvageOptions(readerOptions: new EmailStoreReaderOptions(maxArchiveDecodedBytes: budget)));

            Assert.DoesNotContain(report.Entries, entry => entry.Kind == OutlookItemKind.Contact);
            Assert.Contains(report.NativeDiagnostics, diagnostic =>
                diagnostic.Code == "EMAIL_STORE_OLM_SALVAGE_DECODED_LIMIT" && diagnostic.Location == "Local/C/Contacts.xml");
        });
    }

    [Fact]
    public void DoesNotAdmitAnEmptyRootFollowedByMalformedXml() {
        byte[] archive;
        using (var builder = new OlmTestArchiveBuilder()) {
            archive = builder.AddText("Local/Contacts/Contacts.xml", "<contacts/><broken>")
                .AddText("Local/Contacts/Contacts.xml", TwoContacts)
                .Build();
        }
        DbxStoreReaderTests.WithDestination(destination => {
            using var input = new MemoryStream(archive, writable: false);
            EmailStoreSalvageReport report = EmailStoreSalvage.RecoverOlmToDirectory(input, "archive.olm", destination);

            Assert.Equal(EmailStoreSalvageStatus.NotRecovered, report.Entries[0].Status);
            Assert.Equal(new[] { OutlookItemKind.Contact, OutlookItemKind.Contact },
                report.Entries.Skip(1).Select(entry => entry.Kind!.Value));
            Assert.All(report.Entries.Skip(1), entry => Assert.Equal(EmailStoreSalvageStatus.Recovered, entry.Status));
        });
    }

    [Fact]
    public void KeepsACompleteRecordWhenTheNextTagIsCutOff() {
        byte[] archive;
        using (var builder = new OlmTestArchiveBuilder()) {
            archive = builder.AddText("Local/Contacts/Contacts.xml",
                "<contacts><contact><OPFContactCopyDisplayName>First</OPFContactCopyDisplayName></contact><conta").Build();
        }
        DbxStoreReaderTests.WithDestination(destination => {
            using var input = new MemoryStream(archive, writable: false);
            EmailStoreSalvageReport report = EmailStoreSalvage.RecoverOlmToDirectory(input, "archive.olm", destination);

            EmailStoreSalvageEntry contact = Assert.Single(report.Entries);
            Assert.Equal(EmailStoreSalvageStatus.Partial, contact.Status);
            Assert.Equal("EMAIL_STORE_OLM_SALVAGE_ENTRY_REPAIRED", contact.DamageCode);
            using EmailReadResult exported = new EmailDocumentReader().ReadDocument(contact.DestinationPath!);
            Assert.Equal("First", exported.Document.Contact!.DisplayName);
        });
    }

    [Fact]
    public void ReportsOneRecordOverThePropertyLimitAndExportsTheOthers() {
        const string contacts = "<contacts>" +
            "<contact><OPFContactCopyDisplayName>First</OPFContactCopyDisplayName></contact>" +
            "<contact><OPFContactCopyDisplayName>Large</OPFContactCopyDisplayName><a/><b/><c/></contact>" +
            "<contact><OPFContactCopyDisplayName>Third</OPFContactCopyDisplayName></contact></contacts>";
        byte[] archive;
        using (var builder = new OlmTestArchiveBuilder()) {
            archive = builder.AddText("Local/Contacts/Contacts.xml", contacts).Build();
        }
        DbxStoreReaderTests.WithDestination(destination => {
            using var input = new MemoryStream(archive, writable: false);
            EmailStoreSalvageReport report = EmailStoreSalvage.RecoverOlmToDirectory(input, "archive.olm", destination,
                new EmailStoreSalvageOptions(readerOptions: new EmailStoreReaderOptions(maxPropertiesPerItem: 2)));

            Assert.Collection(report.Entries,
                first => Assert.Equal(EmailStoreSalvageStatus.Recovered, first.Status),
                large => {
                    Assert.Equal("olm:item:Local/Contacts/Contacts.xml#1", large.ItemId);
                    Assert.Equal(EmailStoreSalvageStatus.NotRecovered, large.Status);
                    Assert.Equal("EMAIL_STORE_OLM_ITEM_UNREADABLE", large.DamageCode);
                    Assert.Equal(0, large.SourceOffset);
                },
                third => {
                    Assert.Equal(EmailStoreSalvageStatus.Recovered, third.Status);
                    using EmailReadResult exported = new EmailDocumentReader().ReadDocument(third.DestinationPath!);
                    Assert.Equal("Third", exported.Document.Contact!.DisplayName);
                });
        });
    }

    [Fact]
    public void AppliesTheOutcomeBoundInSourceOrder() {
        byte[] archive;
        using (var builder = new OlmTestArchiveBuilder()) {
            archive = builder.AddText("Local/A/Contacts.xml", TwoContacts)
                .AddText("Local/B/Contacts.xml", "<contacts><contact>")
                .Build();
        }
        DbxStoreReaderTests.WithDestination(destination => {
            using var input = new MemoryStream(archive, writable: false);
            EmailStoreSalvageReport report = EmailStoreSalvage.RecoverOlmToDirectory(input, "archive.olm", destination,
                new EmailStoreSalvageOptions(maxRecoveredItems: 1));

            EmailStoreSalvageEntry contact = Assert.Single(report.Entries);
            Assert.Equal(OutlookItemKind.Contact, contact.Kind);
            Assert.Equal(EmailStoreSalvageStatus.Recovered, contact.Status);
            Assert.True(report.StoppedAtLimit);
        });
    }

    [Fact]
    public void ReportsAnOversizedCollectionAsNotRecoveredAndIncomplete() {
        string large = "<contacts><contact><OPFContactCopyDisplayName>" + new string('a', 512) +
            "</OPFContactCopyDisplayName></contact></contacts>";
        byte[] archive;
        using (var builder = new OlmTestArchiveBuilder()) {
            archive = builder.AddText("Local/A/Contacts.xml", TwoContacts)
                .AddText("Local/B/Contacts.xml", large)
                .Build();
        }
        DbxStoreReaderTests.WithDestination(destination => {
            using var input = new MemoryStream(archive, writable: false);
            EmailStoreSalvageReport report = EmailStoreSalvage.RecoverOlmToDirectory(input, "archive.olm", destination,
                new EmailStoreSalvageOptions(readerOptions: new EmailStoreReaderOptions(maxArchiveEntryBytes: 256)));

            Assert.Equal(3, report.Entries.Count);
            EmailStoreSalvageEntry skipped = report.Entries[2];
            Assert.Equal("olm:entry:Local/B/Contacts.xml", skipped.ItemId);
            Assert.Equal(EmailStoreSalvageStatus.NotRecovered, skipped.Status);
            Assert.Equal("EMAIL_STORE_OLM_SALVAGE_ENTRY_LIMIT", skipped.DamageCode);
            Assert.True(report.StoppedAtLimit);
            Assert.False(report.IsComplete);
            string[] manifest = File.ReadAllLines(report.ManifestPath!);
            Assert.StartsWith("olm:entry:Local/B/Contacts.xml\tNotRecovered\t", manifest[3], StringComparison.Ordinal);
            Assert.Contains(manifest.Skip(4), line =>
                line.StartsWith("source:EMAIL_STORE_OLM_SALVAGE_ENTRY_LIMIT\tError\t", StringComparison.Ordinal));
        });
    }

    [Fact]
    public void KeepsTheRecordsReadBeforeTheXmlCharacterLimit() {
        string contacts = "<contacts><contact><OPFContactCopyDisplayName>First</OPFContactCopyDisplayName></contact>" +
            "<contact><OPFContactCopyDisplayName>" + new string('a', 40_000) + "</OPFContactCopyDisplayName></contact></contacts>";
        byte[] archive;
        using (var builder = new OlmTestArchiveBuilder()) {
            archive = builder.AddText("Local/Contacts/Contacts.xml", contacts).Build();
        }
        DbxStoreReaderTests.WithDestination(destination => {
            using var input = new MemoryStream(archive, writable: false);
            EmailStoreSalvageReport report = EmailStoreSalvage.RecoverOlmToDirectory(input, "archive.olm", destination,
                new EmailStoreSalvageOptions(readerOptions: new EmailStoreReaderOptions(maxXmlCharactersPerItem: 20_000)));

            EmailStoreSalvageEntry contact = Assert.Single(report.Entries);
            Assert.Equal(EmailStoreSalvageStatus.Partial, contact.Status);
            Assert.True(report.StoppedAtLimit);
            using EmailReadResult exported = new EmailDocumentReader().ReadDocument(contact.DestinationPath!);
            Assert.Equal("First", exported.Document.Contact!.DisplayName);
        });
    }

    [Fact]
    public void RecoversFromAStreamPositionedAwayFromTheStart() {
        DbxStoreReaderTests.WithDestination(destination => {
            using var input = new FileStream(DbxStoreReaderTests.DamagedFixture("native-outlook-no-central-directory.olm"),
                FileMode.Open, FileAccess.Read);
            input.Position = 11;
            EmailStoreSalvageReport report = EmailStoreSalvage.RecoverOlmToDirectory(input, "archive.olm", destination);

            Assert.Equal(11, input.Position);
            Assert.Equal(3, report.RecoveredCount);
            Assert.All(report.Entries, entry => Assert.Equal(EmailStoreSalvageStatus.Recovered, entry.Status));
        });
    }

    [Fact]
    public void ReportsEveryRejectedRecordAfterTheDiagnosticCatalogFills() {
        const int rejected = 10_001;
        string contacts = "<contacts>" + string.Concat(Enumerable.Repeat("<contact><a/><b/><c/></contact>", rejected)) +
            "<contact><OPFContactCopyDisplayName>Last</OPFContactCopyDisplayName></contact></contacts>";
        byte[] archive;
        using (var builder = new OlmTestArchiveBuilder()) {
            archive = builder.AddText("Local/Contacts/Contacts.xml", contacts).Build();
        }
        DbxStoreReaderTests.WithDestination(destination => {
            using var input = new MemoryStream(archive, writable: false);
            EmailStoreSalvageReport report = EmailStoreSalvage.RecoverOlmToDirectory(input, "archive.olm", destination,
                new EmailStoreSalvageOptions(readerOptions: new EmailStoreReaderOptions(maxPropertiesPerItem: 2),
                    maxRecoveredItems: 20_000));

            Assert.Equal(rejected + 1, report.Entries.Count);
            Assert.Equal(rejected, report.Entries.Count(entry => entry.DamageCode == "EMAIL_STORE_OLM_ITEM_UNREADABLE"));
            Assert.Equal("olm:item:Local/Contacts/Contacts.xml#" + (rejected - 1), report.Entries[rejected - 1].ItemId);
            Assert.Equal(EmailStoreSalvageStatus.Recovered, report.Entries[rejected].Status);
        });
    }

    [Fact]
    public void OpensARepairedRecordThatEscapesPastTheXmlCharacterLimit() {
        // The writer escapes each '>' in text, so the repaired record is longer than the character limit that the
        // source met; the trailing text after the root forces the repair.
        string source = "<contacts><contact><OPFContactCopyDisplayName>" + new string('>', 40) +
            "</OPFContactCopyDisplayName></contact></contacts>x";
        byte[] archive;
        using (var builder = new OlmTestArchiveBuilder()) {
            archive = builder.AddText("Local/Contacts/Contacts.xml", source).Build();
        }
        DbxStoreReaderTests.WithDestination(destination => {
            using var input = new MemoryStream(archive, writable: false);
            EmailStoreSalvageReport report = EmailStoreSalvage.RecoverOlmToDirectory(input, "archive.olm", destination,
                new EmailStoreSalvageOptions(readerOptions: new EmailStoreReaderOptions(
                    maxXmlCharactersPerItem: source.Length)));

            EmailStoreSalvageEntry contact = Assert.Single(report.Entries);
            Assert.Equal(EmailStoreSalvageStatus.Partial, contact.Status);
            using EmailReadResult exported = new EmailDocumentReader().ReadDocument(contact.DestinationPath!);
            Assert.Equal(new string('>', 40), exported.Document.Contact!.DisplayName);
        });
    }

    [Fact]
    public void DecodesARepeatedAttachmentEntryOnce() {
        byte[] archive;
        using (var builder = new OlmTestArchiveBuilder()) {
            archive = builder.AddText("Local/Contacts/Contacts.xml", OneContact)
                .AddText("Local/Attachments/file.bin", "first").AddText("Local/Attachments/file.bin", "second").Build();
        }
        DbxStoreReaderTests.WithDestination(destination => {
            using var input = new MemoryStream(archive, writable: false);
            EmailStoreSalvageReport report = EmailStoreSalvage.RecoverOlmToDirectory(input, "archive.olm", destination);

            // Recovery skips the later copy before decoding it, so only one copy reaches the rebuilt archive.
            EmailStoreDiagnostic duplicate = Assert.Single(report.NativeDiagnostics,
                diagnostic => diagnostic.Code == "EMAIL_STORE_OLM_DUPLICATE_ENTRY");
            Assert.Equal("salvage", duplicate.Operation);
            Assert.Equal("Local/Attachments/file.bin", duplicate.Location);
            Assert.Equal(EmailStoreSalvageStatus.Recovered, Assert.Single(report.Entries).Status);
        });
    }

    [Fact]
    public void BoundsTheDescriptorSearchForEntriesWithoutSizes() {
        // Unsized entries whose descriptors are missing would each search the rest of the file.
        const int entries = 24;
        var data = new byte[96 * 1024];
        using var output = new MemoryStream();
        for (int index = 0; index < entries; index++) {
            byte[] name = Encoding.ASCII.GetBytes("Local/Attachments/file" + index.ToString("D2", CultureInfo.InvariantCulture) + ".bin");
            var header = new byte[30 + name.Length];
            BinaryPrimitives.WriteUInt32LittleEndian(header.AsSpan(0), 0x04034b50);
            BinaryPrimitives.WriteUInt16LittleEndian(header.AsSpan(4), 20);
            BinaryPrimitives.WriteUInt16LittleEndian(header.AsSpan(6), 0x0808);
            BinaryPrimitives.WriteUInt16LittleEndian(header.AsSpan(26), (ushort)name.Length);
            name.CopyTo(header, 30);
            output.Write(header, 0, header.Length);
            output.Write(data, 0, data.Length);
        }
        byte[] archive = output.ToArray();
        DbxStoreReaderTests.WithDestination(destination => {
            using var input = new MemoryStream(archive, writable: false);
            EmailStoreSalvageReport report = EmailStoreSalvage.RecoverOlmToDirectory(input, "archive.olm", destination);

            Assert.Single(report.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_STORE_OLM_SALVAGE_SEARCH_LIMIT");
            Assert.True(report.StoppedAtLimit);
        });
    }

#if NET8_0_OR_GREATER
    [Fact]
    public void CreatesSalvageSpillsOnlyTheCurrentUserCanOpen() {
        if (OperatingSystem.IsWindows()) return;
        DbxStoreReaderTests.WithDestination(destination => {
            Directory.CreateDirectory(destination);
            var writer = new EmailStoreSalvageWriter(destination, new EmailStoreSalvageOptions(), sourcePath: null);
            using FileStream spill = writer.CreateSpill(".olm", deleteOnClose: true, out string path);

            Assert.Equal(UnixFileMode.UserRead | UnixFileMode.UserWrite, File.GetUnixFileMode(path));
        });
    }
#endif

    private const string OneContact = "<contacts><contact><OPFContactCopyDisplayName>Only</OPFContactCopyDisplayName></contact></contacts>";

    private static byte[] TwoCollections() {
        using var builder = new OlmTestArchiveBuilder();
        return builder.AddText("Local/A/Contacts.xml", TwoContacts).AddText("Local/B/Contacts.xml", OneContact).Build();
    }

    private static int CentralDirectoryOffset(byte[] archive) =>
        BinaryPrimitives.ReadInt32LittleEndian(archive.AsSpan(archive.Length - 22 + 16));

    [Fact]
    public void FindsTheNextEntryWhenADamagedSizeRunsPastTheEndOfTheFile() {
        byte[] intact = TwoCollections();
        byte[] archive = intact.AsSpan(0, CentralDirectoryOffset(intact)).ToArray();
        BinaryPrimitives.WriteInt32LittleEndian(archive.AsSpan(18), int.MaxValue);
        DbxStoreReaderTests.WithDestination(destination => {
            using var input = new MemoryStream(archive, writable: false);
            EmailStoreSalvageReport report = EmailStoreSalvage.RecoverOlmToDirectory(input, "archive.olm", destination);

            // The damaged entry's records are read up to the end of the file and kept as unverified.
            Assert.Equal(new[] { EmailStoreSalvageStatus.Partial, EmailStoreSalvageStatus.Partial, EmailStoreSalvageStatus.Recovered },
                report.Entries.Select(entry => entry.Status));
            EmailStoreSalvageEntry contact = report.Entries[2];
            using EmailReadResult exported = new EmailDocumentReader().ReadDocument(contact.DestinationPath!);
            Assert.Equal("Only", exported.Document.Contact!.DisplayName);
            Assert.Contains(report.NativeDiagnostics, diagnostic => diagnostic.Code == "EMAIL_STORE_OLM_SALVAGE_ENTRY_SPAN_INVALID");
        });
    }

    [Fact]
    public void ScansLocalHeadersWhenTheDirectoryPointsAtTheWrongOffset() {
        byte[] archive = TwoCollections();
        BinaryPrimitives.WriteInt32LittleEndian(archive.AsSpan(CentralDirectoryOffset(archive) + 42), 5);
        DbxStoreReaderTests.WithDestination(destination => {
            using var input = new MemoryStream(archive, writable: false);
            EmailStoreSalvageReport report = EmailStoreSalvage.RecoverOlmToDirectory(input, "archive.olm", destination);

            Assert.Equal(3, report.Entries.Count);
            Assert.All(report.Entries, entry => Assert.Equal(EmailStoreSalvageStatus.Recovered, entry.Status));
            Assert.True(report.DiscoveryCompleted);
            Assert.Contains(report.NativeDiagnostics, diagnostic => diagnostic.Code == "EMAIL_STORE_OLM_SALVAGE_DIRECTORY_INCONSISTENT");
        });
    }

    [Fact]
    public void ReportsAnArchiveTruncatedBeforeItsFirstEntryAsIncomplete() {
        byte[] archive = TwoCollections().AsSpan(0, 20).ToArray();
        DbxStoreReaderTests.WithDestination(destination => {
            using var input = new MemoryStream(archive, writable: false);
            EmailStoreSalvageReport report = EmailStoreSalvage.RecoverOlmToDirectory(input, "archive.olm", destination);

            Assert.Empty(report.Entries);
            Assert.False(report.DiscoveryCompleted);
            Assert.False(report.IsComplete);
        });
    }

    [Fact]
    public void SkipsAnEntryNestedDeeperThanTheDirectoryBound() {
        byte[] archive;
        using (var builder = new OlmTestArchiveBuilder()) {
            archive = builder.AddText("Local/" + string.Concat(Enumerable.Repeat("d/", 70)) + "Contacts.xml", OneContact)
                .AddText("Local/B/Contacts.xml", OneContact)
                .Build();
        }
        DbxStoreReaderTests.WithDestination(destination => {
            using var input = new MemoryStream(archive, writable: false);
            EmailStoreSalvageReport report = EmailStoreSalvage.RecoverOlmToDirectory(input, "archive.olm", destination);

            Assert.Equal("EMAIL_STORE_OLM_SALVAGE_PATH_LIMIT", report.Entries[0].DamageCode);
            Assert.Equal(EmailStoreSalvageStatus.NotRecovered, report.Entries[0].Status);
            Assert.Equal(EmailStoreSalvageStatus.Recovered, report.Entries[1].Status);
            Assert.True(report.StoppedAtLimit);
        });
    }

    [Fact]
    public void ARejectedRecordDoesNotSpendTheNextRecordsAttachmentBudget() {
        const string messages = "<emails>" +
            "<email><OPFMessageCopySubject>Two</OPFMessageCopySubject><OPFMessageCopyAttachmentList>" +
            "<messageAttachment OPFAttachmentName='a.bin' OPFAttachmentURL='Local/a.bin' />" +
            "<messageAttachment OPFAttachmentName='b.bin' OPFAttachmentURL='Local/b.bin' />" +
            "</OPFMessageCopyAttachmentList></email>" +
            "<email><OPFMessageCopySubject>One</OPFMessageCopySubject><OPFMessageCopyAttachmentList>" +
            "<messageAttachment OPFAttachmentName='c.bin' OPFAttachmentURL='Local/c.bin' />" +
            "</OPFMessageCopyAttachmentList></email></emails>";
        byte[] archive;
        using (var builder = new OlmTestArchiveBuilder()) {
            archive = builder.Add("Local/a.bin", new byte[8]).Add("Local/b.bin", new byte[8]).Add("Local/c.bin", new byte[8])
                .AddText("Local/com.microsoft.__Messages/Inbox/message_00000.xml", messages)
                .Build();
        }
        DbxStoreReaderTests.WithDestination(destination => {
            using var input = new MemoryStream(archive, writable: false);
            EmailStoreSalvageReport report = EmailStoreSalvage.RecoverOlmToDirectory(input, "archive.olm", destination,
                new EmailStoreSalvageOptions(readerOptions: new EmailStoreReaderOptions(maxAttachmentsPerItem: 1,
                    maxTotalAttachmentBytes: 8)));

            Assert.Equal(EmailStoreSalvageStatus.NotRecovered, report.Entries[0].Status);
            EmailStoreSalvageEntry one = report.Entries[1];
            Assert.Equal(EmailStoreSalvageStatus.Recovered, one.Status);
            using EmailReadResult exported = new EmailDocumentReader().ReadDocument(one.DestinationPath!);
            Assert.Equal(8, Assert.Single(exported.Document.Attachments).Content!.Length);
        });
    }

    [Theory]
    [InlineData("<messageAttachment OPFAttachmentName='lost.bin' OPFAttachmentURL='' />")]
    [InlineData("<messageAttachment OPFAttachmentName='lost.bin' />")]
    public void ReportsAnAttachmentRecordWithoutAPathAsPartial(string attachment) {
        string messages = "<emails><email><OPFMessageCopySubject>Blank</OPFMessageCopySubject><OPFMessageCopyAttachmentList>" +
            attachment + "</OPFMessageCopyAttachmentList></email></emails>";
        byte[] archive;
        using (var builder = new OlmTestArchiveBuilder()) {
            archive = builder.AddText("Local/com.microsoft.__Messages/Inbox/message_00000.xml", messages).Build();
        }
        DbxStoreReaderTests.WithDestination(destination => {
            using var input = new MemoryStream(archive, writable: false);
            EmailStoreSalvageReport report = EmailStoreSalvage.RecoverOlmToDirectory(input, "archive.olm", destination);

            EmailStoreSalvageEntry mail = Assert.Single(report.Entries);
            Assert.Equal(EmailStoreSalvageStatus.Partial, mail.Status);
            Assert.Equal("EMAIL_STORE_OLM_SALVAGE_ATTACHMENT_MISSING", mail.DamageCode);
            using EmailReadResult exported = new EmailDocumentReader().ReadDocument(mail.DestinationPath!);
            Assert.Empty(exported.Document.Attachments);
        });
    }

    [Fact]
    public void ScansLocalHeadersWhenTheDirectoryPointsAtAnotherEntryOfTheSameNameLength() {
        byte[] archive = TwoCollections();
        int directory = CentralDirectoryOffset(archive);
        int firstLength = 46 + BinaryPrimitives.ReadUInt16LittleEndian(archive.AsSpan(directory + 28)) +
            BinaryPrimitives.ReadUInt16LittleEndian(archive.AsSpan(directory + 30)) +
            BinaryPrimitives.ReadUInt16LittleEndian(archive.AsSpan(directory + 32));
        int second = BinaryPrimitives.ReadInt32LittleEndian(archive.AsSpan(directory + firstLength + 42));
        BinaryPrimitives.WriteInt32LittleEndian(archive.AsSpan(directory + 42), second);
        DbxStoreReaderTests.WithDestination(destination => {
            using var input = new MemoryStream(archive, writable: false);
            EmailStoreSalvageReport report = EmailStoreSalvage.RecoverOlmToDirectory(input, "archive.olm", destination);

            Assert.Contains(report.NativeDiagnostics, diagnostic => diagnostic.Code == "EMAIL_STORE_OLM_SALVAGE_DIRECTORY_INCONSISTENT");
            Assert.Equal(new long?[] { 0, 0, second }, report.Entries.Select(entry => entry.SourceOffset));
            Assert.All(report.Entries, entry => Assert.Equal(EmailStoreSalvageStatus.Recovered, entry.Status));
        });
    }

    [Fact]
    public void SkipsAnEntryWhoseNameIsLongerThanRecoveryKeeps() {
        byte[] archive;
        using (var builder = new OlmTestArchiveBuilder()) {
            archive = builder.AddText("Local/" + new string('n', 17_000) + ".xml", OneContact)
                .AddText("Local/B/Contacts.xml", OneContact)
                .Build();
        }
        DbxStoreReaderTests.WithDestination(destination => {
            using var input = new MemoryStream(archive, writable: false);
            EmailStoreSalvageReport report = EmailStoreSalvage.RecoverOlmToDirectory(input, "archive.olm", destination);

            Assert.Equal(EmailStoreSalvageStatus.Recovered, Assert.Single(report.Entries).Status);
            Assert.True(report.StoppedAtLimit);
            Assert.Contains(report.NativeDiagnostics, diagnostic => diagnostic.Code == "EMAIL_STORE_OLM_SALVAGE_PATH_LIMIT");
        });
    }

    [Fact]
    public void ReportsTheDecodedBudgetWhenItStopsARepair() {
        const string contacts = "<contacts><contact/><contact><OPFContactCopyDisplayName>Second</OPFContactCopyDisplayName></contact></contacts>x";
        byte[] archive;
        using (var builder = new OlmTestArchiveBuilder()) {
            archive = builder.AddText("Local/Contacts/Contacts.xml", contacts).Build();
        }
        DbxStoreReaderTests.WithDestination(destination => {
            using var input = new MemoryStream(archive, writable: false);
            EmailStoreSalvageReport report = EmailStoreSalvage.RecoverOlmToDirectory(input, "archive.olm", destination,
                new EmailStoreSalvageOptions(readerOptions: new EmailStoreReaderOptions(
                    maxArchiveDecodedBytes: Encoding.UTF8.GetByteCount(contacts) + 32)));

            Assert.Equal(EmailStoreSalvageStatus.Partial, Assert.Single(report.Entries).Status);
            Assert.True(report.StoppedAtLimit);
            Assert.Contains(report.NativeDiagnostics, diagnostic => diagnostic.Code == "EMAIL_STORE_OLM_SALVAGE_ENTRY_REPAIRED" &&
                diagnostic.Message.Contains(nameof(EmailStoreReaderOptions.MaxArchiveDecodedBytes), StringComparison.Ordinal));
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
