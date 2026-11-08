using System.Buffers.Binary;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using OfficeIMO.Email;
using OfficeIMO.Email.Store;

namespace OfficeIMO.Migration.Benchmarks;

/// <summary>Synthetic indexed DBX transport with complete MIME and attachment verification.</summary>
internal sealed class DbxWorkload : MigrationWorkload {
    private const string FolderId = "dbx:folder:root";
    private const string SearchNeedle = "dbx-body-proof";
    private static readonly DateTimeOffset SentAt = new(2026, 11, 6, 12, 0, 0, TimeSpan.Zero);
    private readonly int[] _recordAddresses;
    private MemoryStream? _source;
    private EmailStoreSession? _session;
    private EmailStoreReadResult? _read;
    private EmailStoreContentSearchReport? _search;
    private List<byte[]>? _messages;

    public DbxWorkload(int units) : this(CreateInput(units), units) { }

    private DbxWorkload((byte[] Bytes, int[] Records) fixture, int units) : base(fixture.Bytes, units) =>
        _recordAddresses = fixture.Records;

    public override long OutputBytes => _messages?.Sum(message => message.LongLength) ?? 0;

    public override void Execute(string operation) {
        ReleaseResults();
        VerifiedUnits = 0;
        DiagnosticDigest = string.Empty;
        if (operation is not ("LoadInspect" or "ConvertSave" or "SearchContent"))
            throw new ArgumentException("Unknown DBX benchmark operation.", nameof(operation));
        _source = new MemoryStream(Input, writable: false);
        _session = EmailStoreSession.Open(_source, "Migration.dbx", leaveOpen: false);
        if (operation == "SearchContent") {
            _search = _session.SearchContent(new EmailStoreContentQuery(new[] { SearchNeedle },
                fields: EmailStoreContentSearchFields.TextBody,
                maxItemsScanned: Units, maxResults: Units + 1, continueOnItemError: false));
            return;
        }

        _read = _session.ReadAll();
        if (operation == "ConvertSave") {
            var writer = new EmailDocumentWriter();
            _messages = new List<byte[]>(Units);
            foreach (EmailStoreItem item in _read.Store.Folders.SelectMany(folder => folder.Items))
                _messages.Add(writer.ToBytes(item.Document, EmailFileFormat.Eml));
        }
    }

    public override void Validate() {
        if (_session == null) throw new InvalidOperationException("No DBX operation has run.");
        var diagnostics = new List<string>();
        AddDiagnostics(_session.Diagnostics, diagnostics);
        if (_session.Format != EmailStoreFormat.Dbx || _session.Folders.Count != 1 ||
            _session.Folders[0].Id != FolderId || _session.Folders[0].ParentId != null ||
            _session.Folders[0].Name != "Migration")
            throw new InvalidDataException("The DBX folder identity changed.");

        if (_search != null) {
            AddDiagnostics(_search.Diagnostics, diagnostics);
            if (_search.ItemsScanned != Units || _search.ItemsSkipped != 0 || _search.Results.Count != Units ||
                !_search.IsComplete || _search.StoppedAtItemLimit || _search.StoppedAtResultLimit)
                throw new InvalidDataException("The DBX search did not exhaust the complete input.");
            for (int index = 0; index < Units; index++) {
                EmailStoreContentSearchResult result = _search.Results[index];
                ValidateReference(result.Reference.Id, result.Reference.FolderId, index);
                ValidateSummary(result.Summary, index);
                if (result.MatchedFields != EmailStoreContentSearchFields.TextBody ||
                    result.Snippet?.Contains(SearchNeedle, StringComparison.Ordinal) != true)
                    throw new InvalidDataException("The DBX search did not match the requested message body.");
                // This selected read is outside measurement and verifies all content behind every result.
                ValidateDocument(_session.ReadItem(result.Reference).Document, index);
            }
        } else {
            if (_read == null || _read.Store.Format != EmailStoreFormat.Dbx ||
                _read.Store.Folders.Count != 1 || _read.Store.ItemCount != Units)
                throw new InvalidDataException("The DBX read did not materialize the complete input.");
            EmailStoreFolder folder = _read.Store.Folders[0];
            if (folder.Id != FolderId || folder.AssociatedItems.Count != 0 || folder.Items.Count != Units)
                throw new InvalidDataException("The DBX materialized folder scope changed.");
            AddDiagnostics(_read.Diagnostics, diagnostics);
            for (int index = 0; index < Units; index++) {
                EmailStoreItem item = folder.Items[index];
                ValidateReference(item.Id, item.FolderId, index);
                if (item.IsAssociated || item.IsOrphaned)
                    throw new InvalidDataException("The DBX item classification changed.");
                ValidateDocument(item.Document, index);
            }
        }

        if (_messages != null) {
            if (_messages.Count != Units) throw new InvalidDataException("The DBX conversion omitted messages.");
            for (int index = 0; index < Units; index++) {
                using EmailReadResult reopened = new EmailDocumentReader().ReadDocument(_messages[index]);
                if (reopened.Diagnostics.Count != 0)
                    throw new InvalidDataException("An exported DBX message reported an unexpected diagnostic: " +
                        reopened.Diagnostics[0].Code);
                ValidateDocument(reopened.Document, index);
            }
        }
        AddDiagnostics(_session.Diagnostics, diagnostics);
        VerifiedUnits = Units;
        DiagnosticDigest = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(string.Join("\n", diagnostics))));
    }

    public override void ReleaseResults() {
        _read = null;
        _search = null;
        _messages = null;
        try { _session?.Dispose(); }
        finally {
            _session = null;
            _source?.Dispose();
            _source = null;
            base.ReleaseResults();
        }
    }

    private void ValidateReference(string id, string folderId, int index) {
        string expected = "dbx:item:" + _recordAddresses[index].ToString("x8", CultureInfo.InvariantCulture);
        if (id != expected || folderId != FolderId) throw new InvalidDataException("DBX item identity or order changed.");
    }

    private static void ValidateSummary(EmailStoreItemSummary summary, int index) {
        if (summary.Subject != Subject(index) || summary.MessageId != MessageId(index) ||
            summary.From?.Address != Sender(index) || summary.SentAt != SentAt || summary.ReceivedAt != null ||
            summary.IsRead != (index % 2 == 0) || summary.HasAttachments != true ||
            summary.NormalizedSubject != null || summary.DeclaredSize != null)
            throw new InvalidDataException("The DBX summary lost message metadata.");
    }

    private static void ValidateDocument(EmailDocument document, int index) {
        if (document.Format != EmailFileFormat.Eml || document.OutlookItemKind != OutlookItemKind.Message ||
            document.Subject != Subject(index) || document.MessageId != MessageId(index) ||
            document.From?.Address != Sender(index) || document.From?.DisplayName != "Sender " + Number(index) ||
            document.Date != SentAt || document.ReceivedDate != null ||
            document.MessageMetadata.IsRead != (index % 2 == 0) ||
            document.MessageMetadata.IsDraft != (index % 3 == 0) ||
            document.MessageMetadata.NormalizedSubject != null || document.MessageMetadata.Importance != null ||
            document.MessageMetadata.DeclaredSize != null || document.Body.Text != Body(index) ||
            document.Body.Html != null || document.Body.Rtf != null || document.Recipients.Count != 1 ||
            document.Recipients[0].Kind != EmailRecipientKind.To ||
            document.Recipients[0].Address.Address != "recipient" + Number(index) + "@example.test" ||
            document.Attachments.Count != 1)
            throw new InvalidDataException("The DBX message lost body, recipient or metadata content.");
        EmailAttachment attachment = document.Attachments[0];
        byte[] expected = Attachment(index);
        if (attachment.FileName != "proof.txt" || attachment.ContentType != "application/octet-stream" ||
            attachment.IsInline || attachment.EmbeddedDocument != null || attachment.Length != expected.Length)
            throw new InvalidDataException("The DBX attachment metadata changed.");
        using Stream content = attachment.OpenContentStream();
        using var bytes = new MemoryStream();
        content.CopyTo(bytes);
        if (!bytes.ToArray().AsSpan().SequenceEqual(expected))
            throw new InvalidDataException("The DBX attachment bytes changed.");
    }

    private static void AddDiagnostics(IEnumerable<EmailStoreDiagnostic> source, List<string> result) {
        foreach (EmailStoreDiagnostic diagnostic in source) {
            result.Add(diagnostic.Code + ":" + diagnostic.Severity);
            throw new InvalidDataException("DBX validation reported an unexpected diagnostic: " + diagnostic.Code);
        }
    }

    private static string Number(int index) => index.ToString("D6", CultureInfo.InvariantCulture);
    private static string Subject(int index) => "DBX migration message " + Number(index);
    private static string Sender(int index) => "sender" + Number(index) + "@example.test";
    private static string MessageId(int index) => "dbx-" + Number(index) + "@example.test";
    private static string Body(int index) => SearchNeedle + " " + Number(index) + "\r\n" +
        new string((char)('a' + index % 26), 1024) + "\r\n";
    private static byte[] Attachment(int index) => Encoding.ASCII.GetBytes("proof bytes " + Number(index));

    private static byte[] Message(int index) {
        string number = Number(index);
        string boundary = "dbx-benchmark-" + number;
        return Encoding.ASCII.GetBytes("From: Sender " + number + " <" + Sender(index) + ">\r\n" +
            "To: Recipient " + number + " <recipient" + number + "@example.test>\r\n" +
            "Subject: " + Subject(index) + "\r\nDate: Fri, 06 Nov 2026 12:00:00 +0000\r\n" +
            "Message-ID: <" + MessageId(index) + ">\r\nMIME-Version: 1.0\r\n" +
            "Content-Type: multipart/mixed; boundary=" + boundary + "\r\n\r\n" +
            "--" + boundary + "\r\nContent-Type: text/plain; charset=utf-8\r\n\r\n" + Body(index) +
            "\r\n--" + boundary + "\r\nContent-Type: application/octet-stream\r\n" +
            "Content-Disposition: attachment; filename=proof.txt\r\nContent-Transfer-Encoding: base64\r\n\r\n" +
            Convert.ToBase64String(Attachment(index)) + "\r\n--" + boundary + "--\r\n");
    }

    private static (byte[] Bytes, int[] Records) CreateInput(int units) {
        if (units < 1 || units > 512) throw new ArgumentOutOfRangeException(nameof(units));
        using var output = new MemoryStream();
        output.SetLength(0x2600);
        output.Position = output.Length;
        var records = new int[units];
        for (int index = 0; index < units; index++) {
            records[index] = Allocate(output, 20);
            byte[] message = Message(index);
            int firstFragment = checked((int)output.Length);
            for (int start = 0; start < message.Length; start += 512) {
                int address = Allocate(output, 528);
                int length = Math.Min(512, message.Length - start);
                int next = start + length < message.Length ? address + 528 : 0;
                WriteInteger(output, address, address);
                WriteInteger(output, address + 4, 512);
                WriteInteger(output, address + 8, length);
                WriteInteger(output, address + 12, next);
                output.Position = address + 16;
                output.Write(message, start, length);
            }
            int record = records[index];
            int flags = 0x4001 | (index % 2 == 0 ? 0x80 : 0) | (index % 3 == 0 ? 0x08 : 0);
            WriteInteger(output, record, record);
            WriteInteger(output, record + 4, 8);
            output.Position = record + 10;
            output.WriteByte(2);
            WriteInteger(output, record + 12, (flags << 8) | 0x81);
            WriteInteger(output, record + 16, (firstFragment << 8) | 0x84);
        }
        int root = WriteTree(output, records, 0, records.Length, 0);
        byte[] bytes = output.ToArray();
        byte[] signature = { 0xcf, 0xad, 0x12, 0xfe, 0xc5, 0xfd, 0x74, 0x6f,
            0x66, 0xe3, 0xd1, 0x11, 0x9a, 0x4e, 0x00, 0xc0, 0x4f, 0xa3, 0x09, 0xd4, 0x05, 0, 0, 0 };
        signature.CopyTo(bytes, 0);
        BinaryPrimitives.WriteUInt32LittleEndian(bytes.AsSpan(0x1c, 4), 5);
        BinaryPrimitives.WriteUInt32LittleEndian(bytes.AsSpan(0x24bc, 4), 1);
        bytes[0x24c0] = 3;
        BinaryPrimitives.WriteUInt32LittleEndian(bytes.AsSpan(0xc4, 4), (uint)units);
        BinaryPrimitives.WriteUInt32LittleEndian(bytes.AsSpan(0xe4, 4), (uint)root);
        return (bytes, records);
    }

    private static int WriteTree(MemoryStream output, int[] records, int start, int count, int parent) {
        int address = Allocate(output, 0x27c);
        WriteInteger(output, address, address);
        WriteInteger(output, address + 12, parent);
        int entries = count <= 51 ? count : Math.Min(51, count / 52);
        output.Position = address + 17;
        output.WriteByte((byte)entries);
        if (count <= 51) {
            for (int index = 0; index < count; index++)
                WriteInteger(output, address + 24 + index * 12, records[start + index]);
            return address;
        }
        int children = entries + 1;
        int remaining = count - entries;
        int cursor = start;
        for (int childIndex = 0; childIndex < children; childIndex++) {
            int childCount = remaining / children + (childIndex < remaining % children ? 1 : 0);
            int child = WriteTree(output, records, cursor, childCount, address);
            int childPointer = childIndex == 0 ? address + 8 : address + 24 + (childIndex - 1) * 12 + 4;
            WriteInteger(output, childPointer, child);
            WriteInteger(output, childIndex == 0 ? address + 20 : childPointer + 4, childCount);
            cursor += childCount;
            if (childIndex < entries) WriteInteger(output, address + 24 + childIndex * 12, records[cursor++]);
        }
        return address;
    }

    private static int Allocate(MemoryStream output, int count) {
        int address = checked((int)output.Length);
        output.SetLength(checked(output.Length + count));
        return address;
    }

    private static void WriteInteger(MemoryStream output, int offset, int value) {
        Span<byte> bytes = stackalloc byte[4];
        BinaryPrimitives.WriteUInt32LittleEndian(bytes, (uint)value);
        output.Position = offset;
        output.Write(bytes);
    }
}
