using System.IO.Compression;
using OfficeIMO.Email.Store;

namespace OfficeIMO.Email.Tests;

public sealed class EmailStoreSelectedSourceTests {
    [Theory]
    [InlineData("olm", false, false)]
    [InlineData("olm", false, true)]
    [InlineData("olm", true, false)]
    [InlineData("mbox", false, false)]
    [InlineData("mbox", false, true)]
    [InlineData("mbox", true, false)]
    public void SelectedReadsRejectChangedSourcesAndSnapshotsRetainOriginalBytes(string format, bool snapshot, bool duringRead) {
        byte[] original = Artifact(format, "Original");
        byte[] modified = Artifact(format, "Modified");
        Assert.Equal(original.Length, modified.Length);
        using var input = new MutatingInput();
        input.Write(original, 0, original.Length);
        input.Position = 2;
        using (EmailStoreSession session = snapshot ? EmailStoreSession.OpenSnapshot(input, "mutable." + format) :
                   EmailStoreSession.Open(input, "mutable." + format, leaveOpen: true)) {
            EmailStoreItemReference reference = Assert.Single(session.EnumerateItems());
            Assert.Equal("Original", session.ReadSummary(reference).Subject);
            Assert.Equal("Original", session.ReadItem(reference).Document.Subject);
            void Mutate() {
                long position = input.Position;
                input.Position = 0;
                input.Write(modified, 0, modified.Length);
                input.Position = position;
            }
            if (duringRead) input.AfterEndOfRead = Mutate;
            else Mutate();
            if (snapshot) Assert.Equal("Original", session.ReadItem(reference).Document.Subject);
            else {
                Assert.Throws<InvalidDataException>(() => session.ReadItem(reference));
                // A failed identity check permanently invalidates the session even if bytes are restored.
                input.Position = 0;
                input.Write(original, 0, original.Length);
                Assert.Throws<InvalidDataException>(() => session.ReadItem(reference));
            }
        }
        Assert.True(input.CanRead);
        Assert.Equal(2, input.Position);
    }

    [Theory]
    [InlineData("eml")]
    [InlineData("emlx")]
    [InlineData("mbox")]
    public void DirectoryProjectionPinsTheSelectedFileContent(string format) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-source-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string folder = format == "mbox" ? Path.Combine(root, "Inbox.mbox") : root;
            Directory.CreateDirectory(folder);
            string file = Path.Combine(folder, format == "mbox" ? "mbox" : "message." + format);
            byte[] original = Artifact(format, "Original");
            byte[] modified = Artifact(format, "Modified");
            Assert.Equal(original.Length, modified.Length);
            File.WriteAllBytes(file, original);
            using EmailStoreSession session = EmailStoreSession.Open(root);
            EmailStoreItemReference reference = Assert.Single(session.EnumerateItems());
            Assert.Equal("Original", session.ReadSummary(reference).Subject);
            File.WriteAllBytes(file, modified);
            Assert.Throws<InvalidDataException>(() => session.ReadItem(reference));
        } finally { Directory.Delete(root, recursive: true); }
    }

    [Theory]
    [InlineData("olm")]
    [InlineData("mbox")]
    public void ChangedArchiveExpiresPreviouslyReturnedAttachmentReaders(string format) {
        byte[] original = Artifact(format, "Original", attachment: true);
        byte[] modified = Artifact(format, "Modified", attachment: true);
        Assert.Equal(original.Length, modified.Length);
        using var input = new MemoryStream(original, writable: true);
        using EmailStoreSession session = EmailStoreSession.Open(input, "mutable." + format, leaveOpen: true);
        EmailStoreItemReference reference = Assert.Single(session.EnumerateItems());
        EmailStoreItem first = session.ReadItem(reference, new EmailStoreItemReadOptions(preferStreamingAttachmentContent: true));
        IEmailContentSource content = Assert.Single(first.Document.Attachments).ContentSource!;
        using Stream outstanding = content.OpenRead();
        Assert.Equal((int)'O', outstanding.ReadByte());
        input.Position = 0;
        input.Write(modified, 0, modified.Length);
        Assert.Throws<InvalidDataException>(() => session.ReadItem(reference));
        Assert.Throws<ObjectDisposedException>(() => content.OpenRead());
        Assert.Throws<ObjectDisposedException>(() => outstanding.ReadByte());
    }

    private static byte[] Artifact(string format, string subject, bool attachment = false) {
        string message = "Subject: " + subject + (attachment ?
            "\r\nMIME-Version: 1.0\r\nContent-Type: multipart/mixed; boundary=part\r\n\r\n" +
            "--part\r\nContent-Type: text/plain\r\n\r\nBody\r\n--part\r\nContent-Type: application/octet-stream\r\n" +
            "Content-Disposition: attachment; filename=payload.bin\r\nContent-Transfer-Encoding: base64\r\n\r\nT3JpZ2luYWw=\r\n--part--\r\n" : "\r\n\r\nBody\r\n");
        if (format == "eml") return Encoding.ASCII.GetBytes(message);
        if (format == "emlx") return Encoding.ASCII.GetBytes(Encoding.ASCII.GetByteCount(message) + "\n" + message);
        if (format == "mbox") return Encoding.ASCII.GetBytes("From sender@example.com Sat Oct  3 12:00:00 2026\r\n" + message);
        using var output = new MemoryStream();
        using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
            ZipArchiveEntry entry = archive.CreateEntry("Local/Inbox/items.xml", CompressionLevel.NoCompression);
            entry.LastWriteTime = new DateTimeOffset(2026, 10, 3, 12, 0, 0, TimeSpan.Zero);
            using (Stream xml = entry.Open()) {
                byte[] bytes = Encoding.UTF8.GetBytes("<emails><email><OPFMessageCopySubject>" + subject +
                    "</OPFMessageCopySubject><OPFMessageCopyBody>Body</OPFMessageCopyBody>" + (attachment ?
                    "<OPFMessageCopyAttachmentList><messageAttachment OPFAttachmentName='payload.bin' OPFAttachmentURL='Local/attachment' /></OPFMessageCopyAttachmentList>" : "") + "</email></emails>");
                xml.Write(bytes, 0, bytes.Length);
            }
            if (attachment) {
                using Stream payload = archive.CreateEntry("Local/attachment", CompressionLevel.NoCompression).Open();
                byte[] bytes = Encoding.ASCII.GetBytes("Original");
                payload.Write(bytes, 0, bytes.Length);
            }
        }
        return output.ToArray();
    }

    private sealed class MutatingInput : MemoryStream {
        internal Action? AfterEndOfRead { get; set; }
        public override int Read(byte[] buffer, int offset, int count) {
            int read = base.Read(buffer, offset, count);
            if (read == 0 && AfterEndOfRead is { } change) { AfterEndOfRead = null; change(); }
            return read;
        }
    }
}
