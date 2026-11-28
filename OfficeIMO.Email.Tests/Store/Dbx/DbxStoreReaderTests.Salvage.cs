using System.Buffers.Binary;
using System.Security.Cryptography;
using OfficeIMO.Email.Store;

namespace OfficeIMO.Email.Tests;

public sealed partial class DbxStoreReaderTests {
    private const string WelcomeMimeSha256 = "5690ac3f898d12554c351767385901b1281720a1b485b08057b47ced59891ec9";
    private const string WelcomeItem = "dbx:item:00002e10";
    private const int WelcomeMimeBytes = 10139;
    private const long FragmentPayload = 0x200;

    [Fact]
    public void SalvagesAnIntactDbxByteForByteWithProvenance() {
        WithDestination(destination => {
            EmailStoreSalvageReport report = EmailStoreSalvage.RecoverDbxToDirectory(
                Path.Combine(CorpusRoot, "dbx-metadata-unknown-producer", "Inbox.dbx"), destination);

            EmailStoreSalvageEntry entry = Assert.Single(report.Entries);
            Assert.Equal(WelcomeItem, entry.ItemId);
            Assert.Equal(EmailStoreSalvageStatus.Recovered, entry.Status);
            Assert.Equal(true, entry.IsInActiveIndex);
            Assert.Equal(0x2e10, entry.SourceOffset);
            Assert.Equal(WelcomeMimeBytes, entry.RecoveredBytes);
            Assert.Equal("Welcome to Outlook Express 6", entry.Subject);
            Assert.Equal(WelcomeMimeSha256, Sha256(File.ReadAllBytes(entry.DestinationPath!)));
            Assert.True(report.IsComplete);
            Assert.False(report.HasLoss);
            Assert.Empty(report.NativeDiagnostics);
            AssertOnlyArtifactsInside(destination, report);
        });
    }

    [Fact]
    public void SalvagesTheReadablePrefixOfATruncatedDbxAndReportsWhereItStopped() {
        WithDestination(destination => {
            EmailStoreSalvageReport report = EmailStoreSalvage.RecoverDbxToDirectory(
                DamagedFixture("Inbox-truncated.dbx"), destination);

            EmailStoreSalvageEntry entry = Assert.Single(report.Entries);
            Assert.Equal(WelcomeItem, entry.ItemId);
            Assert.Equal(EmailStoreSalvageStatus.Partial, entry.Status);
            Assert.Null(entry.IsInActiveIndex);
            Assert.Equal(7 * FragmentPayload + 100, entry.RecoveredBytes);
            Assert.Equal(WelcomeMimeBytes, entry.DeclaredBytes);
            Assert.Equal(0xf944 + 16 + 100, entry.DamageOffset);
            Assert.Equal("EMAIL_STORE_DBX_SALVAGE_FRAGMENT_TRUNCATED", entry.DamageCode);
            Assert.Equal("Welcome to Outlook Express 6", entry.Subject);
            Assert.Equal(entry.RecoveredBytes, new FileInfo(entry.DestinationPath!).Length);
            Assert.Contains(report.NativeDiagnostics, diagnostic => diagnostic.Code == "EMAIL_STORE_DBX_SALVAGE_INDEX_DAMAGED");
            Assert.False(report.IsComplete);
            Assert.True(report.HasLoss);

            string[] manifest = File.ReadAllLines(report.ManifestPath!);
            Assert.Equal(3, manifest.Length);
            string[] row = manifest[1].Split('\t');
            Assert.Equal(new[] { WelcomeItem, "Partial", "Message", "record 0x2e10", "0x2e10", "3684", "10139", "", "0xf9b8",
                "EMAIL_STORE_DBX_SALVAGE_FRAGMENT_TRUNCATED" }, row.Take(10));
            Assert.StartsWith("source:EMAIL_STORE_DBX_SALVAGE_INDEX_DAMAGED\tWarning\t", manifest[2], StringComparison.Ordinal);
            AssertOnlyArtifactsInside(destination, report);
        });
    }

    [Fact]
    public void StopsAtABrokenDbxFragmentLinkAndReportsTheDetachedTail() {
        WithDestination(destination => {
            EmailStoreSalvageReport report = EmailStoreSalvage.RecoverDbxToDirectory(
                DamagedFixture("Inbox-broken-link.dbx"), destination);

            EmailStoreSalvageEntry entry = Assert.Single(report.Entries);
            Assert.Equal(EmailStoreSalvageStatus.Partial, entry.Status);
            Assert.Equal(true, entry.IsInActiveIndex);
            Assert.Equal(5 * FragmentPayload, entry.RecoveredBytes);
            Assert.Equal(0x12340, entry.DamageOffset);
            Assert.Equal("EMAIL_STORE_DBX_SALVAGE_FRAGMENT_INVALID", entry.DamageCode);
            EmailStoreDiagnostic detached = Assert.Single(report.NativeDiagnostics,
                diagnostic => diagnostic.Code == "EMAIL_STORE_DBX_SALVAGE_UNATTACHED_FRAGMENTS");
            Assert.Equal(0xf524, detached.ByteOffset);
            Assert.StartsWith("15 ", detached.Message, StringComparison.Ordinal);
        });
    }

    [Fact]
    public void RecoversEveryMessageWhenOnlyTheDbxIndexIsDamaged() {
        using var input = new FileStream(DamagedFixture("Inbox-no-index.dbx"), FileMode.Open, FileAccess.Read);
        Assert.Throws<InvalidDataException>(() => EmailStoreSession.Open(input, "Inbox.dbx"));

        WithDestination(destination => {
            input.Position = 11;
            EmailStoreSalvageReport report = EmailStoreSalvage.RecoverDbxToDirectory(input, "Inbox.dbx", destination);
            Assert.Equal(11, input.Position);
            EmailStoreSalvageEntry entry = Assert.Single(report.Entries);
            Assert.Equal(EmailStoreSalvageStatus.Recovered, entry.Status);
            Assert.Null(entry.IsInActiveIndex);
            Assert.Equal(WelcomeMimeSha256, Sha256(File.ReadAllBytes(entry.DestinationPath!)));
            EmailStoreDiagnostic index = Assert.Single(report.NativeDiagnostics);
            Assert.Equal("EMAIL_STORE_DBX_SALVAGE_INDEX_DAMAGED", index.Code);
            Assert.Equal(0xabcde0, index.ByteOffset);
        });
    }

    [Fact]
    public void RebuildsAMessageFromItsFragmentsWhenTheIndexedRecordIsLost() {
        WithDestination(destination => {
            EmailStoreSalvageReport report = EmailStoreSalvage.RecoverDbxToDirectory(
                DamagedFixture("Inbox-lost-record.dbx"), destination);

            Assert.Collection(report.Entries,
                record => {
                    Assert.Equal(WelcomeItem, record.ItemId);
                    Assert.Equal(EmailStoreSalvageStatus.NotRecovered, record.Status);
                    Assert.Equal(true, record.IsInActiveIndex);
                    Assert.Equal("EMAIL_STORE_DBX_SALVAGE_RECORD_INVALID", record.DamageCode);
                    Assert.Null(record.DestinationPath);
                },
                chain => {
                    Assert.Equal("dbx:chain:0000ead4", chain.ItemId);
                    Assert.Equal(EmailStoreSalvageStatus.Recovered, chain.Status);
                    Assert.Equal(false, chain.IsInActiveIndex);
                    Assert.Equal("Welcome to Outlook Express 6", chain.Subject);
                    Assert.Equal("_recovered", Path.GetFileName(Path.GetDirectoryName(chain.DestinationPath)));
                    Assert.Equal(WelcomeMimeSha256, Sha256(File.ReadAllBytes(chain.DestinationPath!)));
                });
        });

        WithDestination(destination => {
            EmailStoreSalvageReport report = EmailStoreSalvage.RecoverDbxToDirectory(
                DamagedFixture("Inbox-lost-record.dbx"), destination,
                new EmailStoreSalvageOptions(includeUnindexedFragmentChains: false));
            Assert.Equal(EmailStoreSalvageStatus.NotRecovered, Assert.Single(report.Entries).Status);
        });
    }

    [Fact]
    public void SeparatesABrokenMessageFromAnIntactOneInTheSameDbx() {
        byte[] database = CreateDatabase();
        UInt32(database, FirstFragment + 12, 0x4800);
        WithDestination(destination => {
            using var input = new MemoryStream(database, writable: false);
            EmailStoreSalvageReport report = EmailStoreSalvage.RecoverDbxToDirectory(input, "Inbox.dbx", destination,
                new EmailStoreSalvageOptions(format: EmailFileFormat.OutlookMsg));

            Assert.Collection(report.Entries,
                first => {
                    Assert.Equal("dbx:item:00003100", first.ItemId);
                    Assert.Equal(EmailStoreSalvageStatus.Partial, first.Status);
                    Assert.Equal(512, first.RecoveredBytes);
                    Assert.Equal(0x4800, first.DamageOffset);
                    using EmailReadResult message = new EmailDocumentReader().ReadDocument(first.DestinationPath!);
                    Assert.Equal("Archived receipt", message.Document.Subject);
                    Assert.Equal(true, message.Document.MessageMetadata.IsRead);
                },
                second => {
                    Assert.Equal("dbx:item:00003140", second.ItemId);
                    Assert.Equal(EmailStoreSalvageStatus.Recovered, second.Status);
                    Assert.EndsWith(".msg", second.DestinationPath, StringComparison.Ordinal);
                    using EmailReadResult message = new EmailDocumentReader().ReadDocument(second.DestinationPath!);
                    Assert.Equal("Other message", message.Document.Subject);
                });
            Assert.Equal(SecondFragment, Assert.Single(report.NativeDiagnostics).ByteOffset);
        });
    }

    [Fact]
    public void ReadsUnorderedIndexedRecordsAndReportsIndexedRecordsWithoutABody() {
        byte[] database = CreateDatabase();
        // Swap the first record's two direct entries and retag the second record's body pointer.
        uint flags = BinaryPrimitives.ReadUInt32LittleEndian(database.AsSpan(FirstRecord + 12));
        uint body = BinaryPrimitives.ReadUInt32LittleEndian(database.AsSpan(FirstRecord + 16));
        UInt32(database, FirstRecord + 12, (int)body);
        UInt32(database, FirstRecord + 16, (int)flags);
        database[SecondRecord + 16] = 0x03;
        WithDestination(destination => {
            using var input = new MemoryStream(database, writable: false);
            EmailStoreSalvageReport report = EmailStoreSalvage.RecoverDbxToDirectory(input, "Inbox.dbx", destination);

            Assert.Collection(report.Entries,
                first => Assert.Equal(EmailStoreSalvageStatus.Recovered, first.Status),
                second => {
                    Assert.Equal("dbx:item:00003140", second.ItemId);
                    Assert.Equal(EmailStoreSalvageStatus.NotRecovered, second.Status);
                    Assert.Equal("EMAIL_STORE_DBX_SALVAGE_NO_PAYLOAD", second.DamageCode);
                },
                orphan => {
                    // The body the second record no longer points at is rebuilt from its fragments.
                    Assert.Equal("dbx:chain:00003720", orphan.ItemId);
                    Assert.Equal(EmailStoreSalvageStatus.Recovered, orphan.Status);
                    Assert.Equal("Other message", orphan.Subject);
                });
        });

        WithDestination(destination => {
            using var input = new MemoryStream(CreateDatabase(), writable: false);
            EmailStoreSalvageReport report = EmailStoreSalvage.RecoverDbxToDirectory(input, "Inbox.dbx", destination,
                new EmailStoreSalvageOptions(readerOptions: new EmailStoreReaderOptions(maxItemCount: 1)));
            Assert.Equal("dbx:item:00003100", Assert.Single(report.Entries).ItemId);
            Assert.True(report.StoppedAtLimit);
        });

        WithDestination(destination => {
            using var input = new MemoryStream(CreateDatabase(), writable: false);
            EmailStoreSalvageReport report = EmailStoreSalvage.RecoverDbxToDirectory(input, "Inbox.dbx", destination,
                new EmailStoreSalvageOptions(readerOptions: new EmailStoreReaderOptions(maxNodeCount: 2)));
            Assert.Equal(2, report.Entries.Count);
            Assert.False(report.DiscoveryCompleted);
            Assert.False(report.IsComplete);
        });
    }

    [Fact]
    public void KeepsExistingSalvageOutputAndHonoursTheItemBound() {
        WithDestination(destination => {
            string source = DamagedFixture("Inbox-lost-record.dbx");
            EmailStoreSalvageReport bounded = EmailStoreSalvage.RecoverDbxToDirectory(source, destination,
                new EmailStoreSalvageOptions(maxRecoveredItems: 1));
            Assert.Single(bounded.Entries);
            Assert.True(bounded.StoppedAtLimit);
            Assert.False(bounded.IsComplete);

            EmailStoreSalvageReport again = EmailStoreSalvage.RecoverDbxToDirectory(
                Path.Combine(CorpusRoot, "dbx-metadata-unknown-producer", "Inbox.dbx"), destination,
                new EmailStoreSalvageOptions(writeManifest: false));
            EmailStoreSalvageReport repeated = EmailStoreSalvage.RecoverDbxToDirectory(
                Path.Combine(CorpusRoot, "dbx-metadata-unknown-producer", "Inbox.dbx"), destination);
            Assert.Equal(EmailStoreSalvageStatus.Recovered, Assert.Single(again.Entries).Status);
            EmailStoreSalvageEntry blocked = Assert.Single(repeated.Entries);
            Assert.Equal(EmailStoreSalvageStatus.NotRecovered, blocked.Status);
            Assert.Contains(blocked.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_STORE_SALVAGE_DESTINATION_EXISTS");
            Assert.Null(repeated.ManifestPath);
            Assert.Contains(repeated.NativeDiagnostics, diagnostic => diagnostic.Code == "EMAIL_STORE_SALVAGE_MANIFEST_EXISTS");
            Assert.Empty(Directory.GetFiles(destination, ".officeimo-*", SearchOption.AllDirectories));
        });
    }

    [Fact]
    public void NeverReplacesASourceFileNamedLikeTheManifest() {
        WithDestination(destination => {
            Directory.CreateDirectory(destination);
            string source = Path.Combine(destination, "officeimo-email-store-salvage.tsv");
            File.Copy(Path.Combine(CorpusRoot, "dbx-metadata-unknown-producer", "Inbox.dbx"), source);
            byte[] original = File.ReadAllBytes(source);

            EmailStoreSalvageReport report = EmailStoreSalvage.RecoverDbxToDirectory(source, destination,
                new EmailStoreSalvageOptions(overwriteExisting: true));

            Assert.Equal(EmailStoreSalvageStatus.Recovered, Assert.Single(report.Entries).Status);
            Assert.Null(report.ManifestPath);
            Assert.Contains(report.NativeDiagnostics, diagnostic => diagnostic.Code == "EMAIL_STORE_SALVAGE_MANIFEST_IS_SOURCE");
            Assert.Equal(original, File.ReadAllBytes(source));
        });
    }

    [Fact]
    public void RefusesToSalvageAFolderCatalog() {
        WithDestination(destination => Assert.Throws<InvalidDataException>(() =>
            EmailStoreSalvage.RecoverDbxToDirectory(
                Path.Combine(CorpusRoot, "dbx-metadata-unknown-producer", "Folders.dbx"), destination)));
    }

    [Fact]
    public void StopsAChainThatLinksIntoAnotherMessageInsteadOfSplicingIt() {
        byte[] database = CreateDatabase();
        // The first message's link is damaged so it points at the body the second record owns.
        UInt32(database, FirstFragment + 12, OtherMessage);
        WithDestination(destination => {
            using var input = new MemoryStream(database, writable: false);
            EmailStoreSalvageReport report = EmailStoreSalvage.RecoverDbxToDirectory(input, "Inbox.dbx", destination);

            Assert.Collection(report.Entries,
                first => {
                    Assert.Equal(EmailStoreSalvageStatus.Partial, first.Status);
                    Assert.Equal(512, first.RecoveredBytes);
                    Assert.Equal(OtherMessage, first.DamageOffset);
                    Assert.Equal("EMAIL_STORE_DBX_SALVAGE_FRAGMENT_CONFLICT", first.DamageCode);
                    Assert.Equal(database.AsSpan(FirstFragment + 16, 512).ToArray(), File.ReadAllBytes(first.DestinationPath!));
                },
                second => {
                    Assert.Equal(EmailStoreSalvageStatus.Recovered, second.Status);
                    Assert.Equal("Subject: Other message\r\n\r\nBody\r\n", File.ReadAllText(second.DestinationPath!));
                });
        });
    }

    [Fact]
    public void ReadsAnIndexedBodyWhoseOptionalPropertyIsDamaged() {
        byte[] database = CreateDatabase();
        // Give the second record a subject whose value offset lies outside the record.
        UInt32(database, SecondRecord + 4, 16);
        database[SecondRecord + 10] = 3;
        UInt32(database, SecondRecord + 20, (0xffff << 8) | 0x05);
        UInt32(database, SecondRecord + 24, OtherMessage);
        WithDestination(destination => {
            using var input = new MemoryStream(database, writable: false);
            EmailStoreSalvageReport report = EmailStoreSalvage.RecoverDbxToDirectory(input, "Inbox.dbx", destination,
                new EmailStoreSalvageOptions(includeUnindexedFragmentChains: false));

            EmailStoreSalvageEntry second = report.Entries[1];
            Assert.Equal("dbx:item:00003140", second.ItemId);
            Assert.Equal(EmailStoreSalvageStatus.Recovered, second.Status);
            Assert.Equal(true, second.IsInActiveIndex);
            Assert.Equal("Subject: Other message\r\n\r\nBody\r\n", File.ReadAllText(second.DestinationPath!));
            Assert.Contains(second.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_STORE_DBX_SALVAGE_PROPERTY_UNREADABLE");
        });
    }

    [Fact]
    public void RecoversAnOrphanWhoseLinkPointsAtItself() {
        byte[] database = CreateDatabase();
        database[SecondRecord + 16] = 0x03;
        UInt32(database, OtherMessage + 12, OtherMessage);
        WithDestination(destination => {
            using var input = new MemoryStream(database, writable: false);
            EmailStoreSalvageReport report = EmailStoreSalvage.RecoverDbxToDirectory(input, "Inbox.dbx", destination);

            EmailStoreSalvageEntry orphan = report.Entries[2];
            Assert.Equal("dbx:chain:00003720", orphan.ItemId);
            Assert.Equal(EmailStoreSalvageStatus.Partial, orphan.Status);
            Assert.Equal("EMAIL_STORE_DBX_SALVAGE_FRAGMENT_CYCLE", orphan.DamageCode);
            Assert.Equal("Subject: Other message\r\n\r\nBody\r\n", File.ReadAllText(orphan.DestinationPath!));
        });
    }

    [Fact]
    public void RecoversOrphansFoundBeforeTheScanStopsAtTheNodeBound() {
        byte[] database = CreateDatabase();
        database[SecondRecord + 16] = 0x03;
        byte[] filler = Encoding.ASCII.GetBytes("freed space");
        WriteFragment(database, 0x3a00, filler, 0, filler.Length, 0);
        WithDestination(destination => {
            using var input = new MemoryStream(database, writable: false);
            // Two records and three fragments fit; the fragment after them stops the scan.
            EmailStoreSalvageReport report = EmailStoreSalvage.RecoverDbxToDirectory(input, "Inbox.dbx", destination,
                new EmailStoreSalvageOptions(readerOptions: new EmailStoreReaderOptions(maxNodeCount: 5)));

            Assert.Equal("dbx:chain:00003720", report.Entries[2].ItemId);
            Assert.Equal(EmailStoreSalvageStatus.Recovered, report.Entries[2].Status);
            Assert.False(report.DiscoveryCompleted);
            Assert.Contains(report.NativeDiagnostics, diagnostic =>
                diagnostic.Code == "EMAIL_STORE_SALVAGE_NODE_LIMIT" && diagnostic.ByteOffset == 0x3a00);
            Assert.Contains(File.ReadAllLines(report.ManifestPath!), line =>
                line.StartsWith("source:EMAIL_STORE_SALVAGE_NODE_LIMIT\tWarning\t", StringComparison.Ordinal));
        });
    }

    [Fact]
    public void MarksAConvertedMessageWithUndecodableContentAsPartial() {
        byte[] database = CreateDatabase();
        byte[] damaged = Encoding.ASCII.GetBytes("Subject: Other message\r\nMIME-Version: 1.0\r\n" +
            "Content-Type: multipart/mixed; boundary=part\r\n\r\n--part\r\nContent-Type: text/plain\r\n\r\nBody\r\n" +
            "--part\r\nContent-Type: application/octet-stream\r\nContent-Transfer-Encoding: base64\r\n\r\n" +
            "@@@@ not base64 @@@@\r\n--part--\r\n");
        WriteFragment(database, OtherMessage, damaged, 0, damaged.Length, 0);
        WithDestination(destination => {
            using var input = new MemoryStream(database, writable: false);
            EmailStoreSalvageReport report = EmailStoreSalvage.RecoverDbxToDirectory(input, "Inbox.dbx", destination,
                new EmailStoreSalvageOptions(format: EmailFileFormat.OutlookMsg));

            EmailStoreSalvageEntry second = report.Entries[1];
            Assert.Equal(EmailStoreSalvageStatus.Partial, second.Status);
            Assert.Equal("EMAIL_STORE_DBX_SALVAGE_MIME_DAMAGED", second.DamageCode);
            Assert.NotNull(second.DestinationPath);
            Assert.Contains(second.Diagnostics, diagnostic => diagnostic.Severity == EmailStoreDiagnosticSeverity.Error &&
                diagnostic.Location!.StartsWith("item/dbx:item:00003140", StringComparison.Ordinal));
        });
    }

    internal static string CorpusRoot => Path.Combine(EmailTestRepository.FindRoot(), "OfficeIMO.Email.Tests", "Corpora");

    internal static string DamagedFixture(string name) => Path.Combine(CorpusRoot, "damaged-stores", name);

    internal static void WithDestination(Action<string> test) {
        string parent = Path.Combine(Path.GetTempPath(), "officeimo-salvage-" + Guid.NewGuid().ToString("N"));
        string destination = Path.Combine(parent, "out");
        Directory.CreateDirectory(parent);
        try {
            test(destination);
            Assert.Equal(new[] { destination }, Directory.GetFileSystemEntries(parent));
        } finally {
            Directory.Delete(parent, recursive: true);
        }
    }

    internal static void AssertOnlyArtifactsInside(string destination, EmailStoreSalvageReport report) {
        string root = Path.GetFullPath(destination) + Path.DirectorySeparatorChar;
        string[] expected = report.Entries.Where(entry => entry.DestinationPath != null)
            .Select(entry => entry.DestinationPath!)
            .Concat(report.ManifestPath == null ? Array.Empty<string>() : new[] { report.ManifestPath })
            .OrderBy(path => path, StringComparer.Ordinal).ToArray();
        Assert.All(expected, path => Assert.StartsWith(root, path, StringComparison.Ordinal));
        Assert.Equal(expected, Directory.GetFiles(destination, "*", SearchOption.AllDirectories)
            .OrderBy(path => path, StringComparer.Ordinal));
    }

    private static string Sha256(byte[] bytes) {
        using SHA256 algorithm = SHA256.Create();
        return string.Concat(algorithm.ComputeHash(bytes).Select(value => value.ToString("x2")));
    }
}
