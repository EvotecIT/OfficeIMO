using OfficeIMO.Email;
using System.Globalization;
using System.Security.Cryptography;
using System.Threading;

namespace OfficeIMO.Email.Store.Tests;

public sealed class PstWriterOutlookCompatibilityTests {
    [Fact]
    public void Populated_store_writes_consistent_outlook_message_and_attachment_rows() {
        string path = TemporaryPstPath();
        try {
            using (EmailStorePstWriter writer = EmailStorePstWriter.Create(path,
                new EmailStorePstWriterOptions("Outlook-compatible store"))) {
                string folder = writer.AddFolder("Inbox", EmailStoreSpecialFolderKind.Inbox);
                writer.AddItem(folder, CreateDocument());
                writer.Complete();
            }

            using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            PstHeader header = PstHeader.Read(stream, EmailStoreFormat.Pst);
            var ndb = new PstNdbReader(stream, header, EmailStoreReaderOptions.Default,
                CancellationToken.None);
            ndb.LoadIndexes();
            PstNodeReference message = Assert.Single(ndb.Nodes.Values,
                node => node.Type == 0x04);
            IReadOnlyList<MapiProperty> messageProperties = ReadPropertyContext(ndb,
                message.DataBid, message.SubnodeBid);
            int messageSize = GetInt32(messageProperties, 0x0E08);

            Assert.Equal(GetContextDataLength(ndb, message.DataBid, message.SubnodeBid),
                messageSize);
            byte[] expectedConversationId = ComputeConversationId(
                "Outlook interoperability subject");
            Assert.DoesNotContain(messageProperties, property => property.PropertyId == 0x3013);
            Assert.Equal("Recipient", GetString(messageProperties, 0x0E04));

            IReadOnlyDictionary<uint, PstSubnodeReference> messageSubnodes =
                ndb.ReadSubnodes(message.SubnodeBid);
            PstSubnodeReference attachment = Assert.Single(messageSubnodes.Values,
                node => node.Type == 0x05);
            IReadOnlyList<MapiProperty> attachmentProperties = ReadPropertyContext(ndb,
                attachment.DataBid, attachment.SubnodeBid);
            int attachmentSize = GetInt32(attachmentProperties, 0x0E20);
            long attachmentLogicalSize = attachmentProperties.Sum(property =>
                PstPropertyValueWriter.GetLogicalValueSize(property, 65001));
            Assert.Equal(attachmentLogicalSize, attachmentSize);

            PstNodeReference contents = Assert.Single(ndb.Nodes.Values,
                node => node.Type == 0x0E && ReadTableRows(ndb, node).Count == 1);
            IReadOnlyList<uint> columns = ReadTableColumns(ndb, contents);
            Assert.Contains(0x0E300003U, columns);
            Assert.Contains(0x0E300102U, columns);
            IReadOnlyList<MapiProperty> row = Assert.Single(ReadTableRows(ndb, contents));
            Assert.Equal(messageSize, GetInt32(row, 0x0E08));
            Assert.Equal("Recipient", GetString(row, 0x0E04));
            Assert.DoesNotContain(row, property => property.PropertyTag == 0x0E300003U);
            Assert.Equal(16, Assert.Single(row, property =>
                property.PropertyTag == 0x0E300102U).RawData?.Length ??
                Assert.IsType<byte[]>(Assert.Single(row, property =>
                    property.PropertyTag == 0x0E300102U).Value).Length);
            Assert.Equal(24, GetBinary(row, 0x0E34).Length);
            Assert.Equal(expectedConversationId, GetBinary(row, 0x3013));
        } finally {
            TryDelete(path);
        }
    }

    [Fact]
    public void Legacy_zero_reference_checkpoint_rebuilds_internal_block_references() {
        string directory = Path.Combine(Path.GetTempPath(),
            "officeimo-pst-legacy-refcounts-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        string path = Path.Combine(directory, "resumed.pst");
        string checkpoint = Path.Combine(directory, "resumed.checkpoint");
        try {
            using (EmailStorePstWriter writer = EmailStorePstWriter.Create(path,
                new EmailStorePstWriterOptions(checkpointPath: checkpoint,
                    checkpointIntervalItems: 1))) {
                string folder = writer.AddFolder("Inbox");
                var document = new EmailDocument { Subject = "Before resume" };
                document.Attachments.Add(new EmailAttachment {
                    FileName = "multi-block.bin",
                    Content = Enumerable.Range(0, 20_000).Select(index =>
                        checked((byte)(index % 251))).ToArray(),
                    Length = 20_000
                });
                writer.AddItem(folder, document);
            }

            string blockJournal = Assert.Single(Directory.EnumerateFiles(directory),
                candidate => candidate.EndsWith(".blocks", StringComparison.Ordinal));
            using (var journal = new FileStream(blockJournal, FileMode.Open,
                FileAccess.ReadWrite, FileShare.None)) {
                var zero = new byte[4];
                for (long offset = 20; offset < journal.Length; offset += 24) {
                    journal.Position = offset;
                    journal.Write(zero, 0, zero.Length);
                }
            }

            using (EmailStorePstWriter resumed = EmailStorePstWriter.Resume(checkpoint)) {
                resumed.Complete();
            }

            using EmailStoreSession session = EmailStoreSession.Open(path);
            EmailStoreValidationReport validation = session.Validate(
                new EmailStoreValidationOptions(
                    mode: EmailStoreValidationMode.FullItems,
                    verifyStructuralIntegrity: true,
                    maxStructuralPages: 10_000,
                    maxStructuralBlocks: 10_000,
                    maxStructuralBytes: 128 * 1024 * 1024));
            Assert.Equal(0, validation.StructuralFailures);
            Assert.Equal(20_000, Assert.Single(session.ReadItem(
                Assert.Single(session.EnumerateItems())).Document.Attachments).Length);
        } finally {
            try { if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true); }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }
    }

    [Fact]
    public void Interrupted_legacy_reference_migration_resets_and_rebuilds_idempotently() {
        string directory = Path.Combine(Path.GetTempPath(),
            "officeimo-pst-interrupted-refcounts-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        string path = Path.Combine(directory, "resumed.pst");
        string checkpoint = Path.Combine(directory, "resumed.checkpoint");
        try {
            using (EmailStorePstWriter writer = EmailStorePstWriter.Create(path,
                new EmailStorePstWriterOptions(checkpointPath: checkpoint,
                    checkpointIntervalItems: 1))) {
                string folder = writer.AddFolder("Inbox");
                var document = new EmailDocument { Subject = "Before interrupted migration" };
                document.Attachments.Add(new EmailAttachment {
                    FileName = "multi-block.bin",
                    Content = Enumerable.Range(0, 20_000).Select(index =>
                        checked((byte)(index % 251))).ToArray(),
                    Length = 20_000
                });
                writer.AddItem(folder, document);
            }

            string blockJournal = Assert.Single(Directory.EnumerateFiles(directory),
                candidate => candidate.EndsWith(".blocks", StringComparison.Ordinal));
            File.WriteAllBytes(string.Concat(blockJournal, ".refcounts"), new byte[] { 1 });
            using (var journal = new FileStream(blockJournal, FileMode.Open,
                FileAccess.ReadWrite, FileShare.None))
            using (var writer = new BinaryWriter(journal, Encoding.UTF8, leaveOpen: true)) {
                int index = 0;
                for (long offset = 20; offset < journal.Length; offset += 24, index++) {
                    journal.Position = offset;
                    writer.Write(index % 3 == 0 ? 7 : index % 3 == 1 ? 0 : 2);
                }
            }

            using (EmailStorePstWriter resumed = EmailStorePstWriter.Resume(checkpoint)) {
                Assert.False(File.Exists(string.Concat(blockJournal, ".refcounts")));
                AssertJournalReferenceCounts(blockJournal);
                resumed.Complete();
            }

            using EmailStoreSession session = EmailStoreSession.Open(path);
            Assert.Equal(20_000, Assert.Single(session.ReadItem(
                Assert.Single(session.EnumerateItems())).Document.Attachments).Length);
        } finally {
            try { if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true); }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }
    }

    [Fact]
    public void Contents_rows_follow_final_mapi_patch_and_conversation_index_tracking() {
        string path = TemporaryPstPath();
        try {
            byte[] retainedConversationId = Enumerable.Range(1, 16)
                .Select(value => checked((byte)value)).ToArray();
            byte[] patchedSearchKey = Encoding.ASCII.GetBytes("PATCHED-SEARCH-KEY");
            EmailDocument retained = CreateDocument();
            retained.Subject = "Retained conversation id";
            retained.MapiWritePatch
                .Set(MapiKnownProperties.PidTag.DisplayTo, "Patched display")
                .Set(MapiKnownProperties.PidTag.MessageStatus, 7)
                .Set(MapiKnownProperties.PidTag.SearchKey, patchedSearchKey)
                .Set(MapiKnownProperties.PidTag.ConversationTopic, "Patched topic")
                .Set(MapiKnownProperties.PidTag.ConversationId, retainedConversationId)
                .Remove(MapiKnownProperties.PidTag.MessageSize);

            byte[] conversationIndex = new byte[22];
            conversationIndex[0] = 0x01;
            for (int index = 0; index < 16; index++) {
                conversationIndex[index + 6] = checked((byte)(0xA0 + index));
            }
            EmailDocument tracked = CreateDocument();
            tracked.Subject = "Tracked conversation index";
            tracked.MapiWritePatch
                .Remove(MapiKnownProperties.PidTag.ConversationId)
                .Set(MapiKnownProperties.PidTag.ConversationIndexTracking, true)
                .Set(MapiKnownProperties.PidTag.ConversationIndex, conversationIndex);

            using (EmailStorePstWriter writer = EmailStorePstWriter.Create(path)) {
                string folder = writer.AddFolder("Inbox");
                writer.AddItem(folder, retained);
                writer.AddItem(folder, tracked);
                writer.Complete();
            }

            using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            PstHeader header = PstHeader.Read(stream, EmailStoreFormat.Pst);
            var ndb = new PstNdbReader(stream, header, EmailStoreReaderOptions.Default,
                CancellationToken.None);
            ndb.LoadIndexes();
            IReadOnlyList<IReadOnlyList<MapiProperty>> messages = ndb.Nodes.Values
                .Where(node => node.Type == 0x04)
                .Select(node => ReadPropertyContext(ndb, node.DataBid, node.SubnodeBid)).ToArray();
            IReadOnlyList<MapiProperty> retainedProperties = Assert.Single(messages,
                properties => GetString(properties, 0x0037) == retained.Subject);
            Assert.Equal("Patched display", GetString(retainedProperties, 0x0E04));
            Assert.Equal(7, GetInt32(retainedProperties, 0x0E17));
            Assert.Equal(patchedSearchKey, GetBinary(retainedProperties, 0x300B));
            Assert.Equal(retainedConversationId, GetBinary(retainedProperties, 0x3013));
            Assert.True(GetInt32(retainedProperties, 0x0E08) > 0);

            IReadOnlyList<MapiProperty> trackedProperties = Assert.Single(messages,
                properties => GetString(properties, 0x0037) == tracked.Subject);
            Assert.DoesNotContain(trackedProperties, property => property.PropertyId == 0x3013);
            Assert.True(Assert.Single(trackedProperties, property =>
                property.PropertyId == 0x3016).Value is true);

            PstNodeReference contents = Assert.Single(ndb.Nodes.Values,
                node => node.Type == 0x0E && ReadTableRows(ndb, node).Count == 2);
            IReadOnlyList<IReadOnlyList<MapiProperty>> rows = ReadTableRows(ndb, contents);
            IReadOnlyList<MapiProperty> retainedRow = Assert.Single(rows,
                row => GetString(row, 0x0037) == retained.Subject);
            Assert.Equal("Patched display", GetString(retainedRow, 0x0E04));
            Assert.Equal(7, GetInt32(retainedRow, 0x0E17));
            Assert.Equal(retainedConversationId, GetBinary(retainedRow, 0x3013));
            IReadOnlyList<MapiProperty> trackedRow = Assert.Single(rows,
                row => GetString(row, 0x0037) == tracked.Subject);
            Assert.Equal(conversationIndex.Skip(6).Take(16).ToArray(),
                GetBinary(trackedRow, 0x3013));
        } finally {
            TryDelete(path);
        }
    }

    private static EmailDocument CreateDocument() {
        var document = new EmailDocument {
            Subject = "Outlook interoperability subject",
            MessageClass = "IPM.Note",
            From = new EmailAddress("sender@example.test", "Sender")
        };
        document.Body.Text = "Outlook interoperability body";
        document.Recipients.Add(new EmailRecipient(EmailRecipientKind.To,
            new EmailAddress("recipient@example.test", "Recipient")));
        byte[] content = Encoding.UTF8.GetBytes("Outlook attachment evidence");
        document.Attachments.Add(new EmailAttachment {
            FileName = "evidence.txt",
            ContentType = "text/plain",
            Content = content,
            Length = content.LongLength
        });
        return document;
    }

    private static void AssertJournalReferenceCounts(string blockJournal) {
        var records = new List<(ulong Bid, long Offset, int Length, int ReferenceCount)>();
        using (var input = new FileStream(blockJournal, FileMode.Open,
            FileAccess.Read, FileShare.ReadWrite))
        using (var reader = new BinaryReader(input, Encoding.UTF8, leaveOpen: false)) {
            while (input.Position < input.Length) {
                records.Add((reader.ReadUInt64(), reader.ReadInt64(),
                    reader.ReadInt32(), reader.ReadInt32()));
            }
        }
        var expected = records.ToDictionary(record => record.Bid & ~3UL, _ => 1);
        string workingPath = blockJournal.Substring(0,
            blockJournal.Length - ".blocks".Length);
        using (var working = new FileStream(workingPath, FileMode.Open,
            FileAccess.Read, FileShare.ReadWrite)) {
            foreach (var record in records.Where(record => (record.Bid & 0x02UL) != 0)) {
                var payload = new byte[record.Length];
                working.Position = record.Offset;
                Assert.Equal(payload.Length, working.Read(payload, 0, payload.Length));
                int count = PstBinary.UInt16(payload, 2);
                if (payload[0] == 0x01) {
                    for (int index = 0; index < count; index++) {
                        AddExpectedReference(expected, PstBinary.UInt64(payload, 8 + index * 8));
                    }
                } else if (payload[0] == 0x02 && payload[1] == 0) {
                    for (int index = 0; index < count; index++) {
                        int offset = 8 + index * 24;
                        AddExpectedReference(expected, PstBinary.UInt64(payload, offset + 8));
                        AddExpectedReference(expected, PstBinary.UInt64(payload, offset + 16));
                    }
                } else if (payload[0] == 0x02 && payload[1] == 1) {
                    for (int index = 0; index < count; index++) {
                        AddExpectedReference(expected, PstBinary.UInt64(payload, 16 + index * 16));
                    }
                }
            }
        }
        Assert.All(records, record => Assert.Equal(expected[record.Bid & ~3UL],
            record.ReferenceCount));
    }

    private static void AddExpectedReference(IDictionary<ulong, int> references, ulong bid) {
        if (bid == 0) return;
        ulong normalized = bid & ~3UL;
        references[normalized] = checked(references[normalized] + 1);
    }

    private static IReadOnlyList<MapiProperty> ReadPropertyContext(PstNdbReader ndb,
        ulong dataBid, ulong subnodeBid) {
        PstDataTree tree = ndb.ReadDataTree(dataBid, 64 * 1024 * 1024);
        IReadOnlyDictionary<uint, PstSubnodeReference> subnodes = ndb.ReadSubnodes(subnodeBid);
        var heap = new PstHeap(tree, subnodes, ndb,
            EmailStoreReaderOptions.Default, CancellationToken.None);
        return new PstPropertyContextReader(heap, EmailStoreReaderOptions.Default,
            CancellationToken.None).ReadProperties();
    }

    private static IReadOnlyList<IReadOnlyList<MapiProperty>> ReadTableRows(
        PstNdbReader ndb, PstNodeReference node) {
        PstDataTree tree = ndb.ReadDataTree(node.DataBid, 64 * 1024 * 1024);
        IReadOnlyDictionary<uint, PstSubnodeReference> subnodes = ndb.ReadSubnodes(node.SubnodeBid);
        var heap = new PstHeap(tree, subnodes, ndb,
            EmailStoreReaderOptions.Default, CancellationToken.None);
        return new PstTableContextReader(heap, true, EmailStoreReaderOptions.Default,
            CancellationToken.None).ReadRows();
    }

    private static IReadOnlyList<uint> ReadTableColumns(PstNdbReader ndb,
        PstNodeReference node) {
        PstDataTree tree = ndb.ReadDataTree(node.DataBid, 64 * 1024 * 1024);
        var heap = new PstHeap(tree, ndb.ReadSubnodes(node.SubnodeBid), ndb,
            EmailStoreReaderOptions.Default, CancellationToken.None);
        byte[] info = heap.GetAllocation(heap.UserRoot);
        int count = info[1];
        return Enumerable.Range(0, count).Select(index =>
            PstBinary.UInt32(info, 22 + index * 8)).ToArray();
    }

    private static long GetContextDataLength(PstNdbReader ndb,
        ulong dataBid, ulong subnodeBid) {
        long length = ndb.GetDataTreeLength(dataBid);
        foreach (PstSubnodeReference subnode in ndb.ReadSubnodes(subnodeBid).Values) {
            length = checked(length + GetContextDataLength(ndb,
                subnode.DataBid, subnode.SubnodeBid));
        }
        return length;
    }

    private static byte[] ComputeConversationId(string topic) {
        using MD5 md5 = MD5.Create();
        return md5.ComputeHash(Encoding.Unicode.GetBytes(topic.ToUpperInvariant()));
    }

    private static int GetInt32(IEnumerable<MapiProperty> properties, ushort propertyId) =>
        Convert.ToInt32(Assert.Single(properties, property =>
            property.PropertyId == propertyId).Value, CultureInfo.InvariantCulture);

    private static string GetString(IEnumerable<MapiProperty> properties, ushort propertyId) =>
        Assert.IsType<string>(Assert.Single(properties, property =>
            property.PropertyId == propertyId).Value);

    private static byte[] GetBinary(IEnumerable<MapiProperty> properties, ushort propertyId) =>
        Assert.IsType<byte[]>(Assert.Single(properties, property =>
            property.PropertyId == propertyId).Value);

    private static string TemporaryPstPath() => Path.Combine(Path.GetTempPath(),
        "officeimo-pst-outlook-compat-" + Guid.NewGuid().ToString("N") + ".pst");

    private static void TryDelete(string path) {
        try { if (File.Exists(path)) File.Delete(path); }
        catch (IOException) { }
        catch (UnauthorizedAccessException) { }
    }
}
