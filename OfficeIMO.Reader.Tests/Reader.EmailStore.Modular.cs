using OfficeIMO.Email.Store;
using OfficeIMO.Reader.EmailStore;
using System.Globalization;
using System.IO.Compression;
using Xunit;

namespace OfficeIMO.Reader.Tests;

public sealed class ReaderEmailStoreModularTests {
    [Fact]
    public void HandlerAdvertisesEveryStoreFormatAndRichReadSurface() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddEmailStoreHandler()
            .Build();

        ReaderHandlerCapability capability = Assert.Single(reader.GetCapabilities(), item =>
            item.Id == OfficeDocumentReaderBuilderEmailStoreExtensions.HandlerId);

        Assert.Equal(ReaderInputKind.Email, capability.Kind);
        Assert.Equal(new[] { ".emlx", ".olm", ".ost", ".pst" }, capability.Extensions);
        Assert.True(capability.SupportsPath);
        Assert.True(capability.SupportsStream);
        Assert.True(capability.SupportsDocumentPath);
        Assert.True(capability.SupportsDocumentStream);
        Assert.True(capability.DeterministicOutput);
        Assert.Equal(EmailStoreReaderOptions.Default.MaxInputBytes, capability.DefaultMaxInputBytes);
    }

    [Fact]
    public void EmlxUsesSharedEmailChunksMetadataDiagnosticsAndAttachmentAssets() {
        byte[] emlx = CreateEmlx(CreateMultipartMessage(),
            "<plist><dict><key>remote-id</key><string>remote-42</string></dict></plist>");
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddEmailStoreHandler()
            .Build();
        using var stream = new MemoryStream(emlx, writable: false);
        stream.Position = 3;

        OfficeDocumentReadResult result = reader.ReadDocument(stream, "42.emlx");

        Assert.Equal(3, stream.Position);
        Assert.Equal(ReaderInputKind.Email, result.Kind);
        Assert.Equal("42", result.Source.Title);
        Assert.Contains("officeimo.reader.emailstore", result.CapabilitiesUsed);
        Assert.Contains("officeimo.email.store", result.CapabilitiesUsed);
        Assert.Contains("officeimo.email.store.emlx", result.CapabilitiesUsed);
        Assert.Contains(result.Chunks, chunk =>
            chunk.Location.Path == "42.emlx!/Apple Mail/item-000000" &&
            chunk.Text.Contains("EMLX Reader contract", StringComparison.Ordinal));
        Assert.Contains(result.Chunks, chunk => chunk.Text.Contains("Portable body", StringComparison.Ordinal));
        Assert.Contains(result.Metadata, item =>
            item.Name == "StoreFormat" && item.Value == "Emlx");
        Assert.Contains(result.Metadata, item =>
            item.Name == "FolderCount" && item.Value == "1");
        OfficeDocumentAsset asset = Assert.Single(result.Assets);
        Assert.Equal("payload.bin", asset.FileName);
        Assert.Equal("application/octet-stream", asset.MediaType);
        Assert.Equal(new byte[] { 1, 2, 3, 4 }, asset.PayloadBytes);
    }

    [Fact]
    public void OlmPreservesFolderHierarchyInLogicalReaderPaths() {
        const string attachmentPath = "Local/com.microsoft.__Messages/Account/Inbox/com.microsoft.__Attachments/logo";
        const string xml = "<emails><email>" +
            "<OPFMessageCopySubject>Nested OLM message</OPFMessageCopySubject>" +
            "<OPFMessageCopyBody>OLM body</OPFMessageCopyBody>" +
            "<OPFMessageCopyAttachmentList><messageAttachment OPFAttachmentName=\"logo.bin\" " +
            "OPFAttachmentURL=\"" + attachmentPath + "\" OPFAttachmentContentType=\"application/octet-stream\" />" +
            "</OPFMessageCopyAttachmentList></email></emails>";
        byte[] archive = CreateOlmArchive(new Dictionary<string, byte[]> {
            ["Local/com.microsoft.__Messages/Account/Inbox/message_00000.xml"] = Encoding.UTF8.GetBytes(xml),
            [attachmentPath] = new byte[] { 9, 8, 7 }
        });
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddEmailStoreHandler()
            .Build();

        OfficeDocumentReadResult result = reader.ReadDocument(archive, "mailbox.olm");

        Assert.Contains(result.Chunks, chunk =>
            chunk.Location.Path == "mailbox.olm!/Local/Account/Inbox/item-000000");
        Assert.Contains("officeimo.email.store.olm", result.CapabilitiesUsed);
        Assert.Contains(result.Metadata, item =>
            item.Name == "FolderCount" && item.Value == "3");
        Assert.Equal(new byte[] { 9, 8, 7 }, Assert.Single(result.Assets).PayloadBytes);
    }

    [Fact]
    public void ReaderLimitNarrowsButCannotWidenRegisteredStoreLimit() {
        byte[] emlx = CreateEmlx(CreateMultipartMessage(), null);
        var adapterOptions = new ReaderEmailStoreOptions {
            StoreOptions = new EmailStoreReaderOptions(maxInputBytes: emlx.Length - 1L)
        };
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddEmailStoreHandler(adapterOptions)
            .Build();

        Exception registeredLimit = Assert.ThrowsAny<Exception>(() => reader.ReadDocument(
            emlx, "message.emlx", new ReaderOptions { MaxInputBytes = emlx.Length + 100L }));
        Assert.Contains("MaxInputBytes", registeredLimit.Message, StringComparison.OrdinalIgnoreCase);

        OfficeDocumentReader defaultReader = new OfficeDocumentReaderBuilder()
            .AddEmailStoreHandler()
            .Build();
        Exception readerLimit = Assert.ThrowsAny<Exception>(() => defaultReader.ReadDocument(
            emlx, "message.emlx", new ReaderOptions { MaxInputBytes = emlx.Length - 1L }));
        Assert.Contains("MaxInputBytes", readerLimit.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void RegistrationCapturesStoreOptionsDefensively() {
        byte[] emlx = CreateEmlx(CreateMultipartMessage(), null);
        var options = new ReaderEmailStoreOptions {
            StoreOptions = new EmailStoreReaderOptions(maxInputBytes: emlx.Length + 100L)
        };
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddEmailStoreHandler(options)
            .Build();
        options.StoreOptions = new EmailStoreReaderOptions(maxInputBytes: 1);

        OfficeDocumentReadResult result = reader.ReadDocument(emlx, "message.emlx");

        Assert.Contains(result.Chunks, chunk =>
            chunk.Text.Contains("EMLX Reader contract", StringComparison.Ordinal));
    }

    private static byte[] CreateEmlx(byte[] message, string? plist) {
        byte[] prefix = Encoding.ASCII.GetBytes(message.Length.ToString(CultureInfo.InvariantCulture) + "\n");
        byte[] metadata = plist == null ? Array.Empty<byte>() : Encoding.UTF8.GetBytes(plist);
        var result = new byte[prefix.Length + message.Length + metadata.Length];
        Buffer.BlockCopy(prefix, 0, result, 0, prefix.Length);
        Buffer.BlockCopy(message, 0, result, prefix.Length, message.Length);
        Buffer.BlockCopy(metadata, 0, result, prefix.Length + message.Length, metadata.Length);
        return result;
    }

    private static byte[] CreateMultipartMessage() {
        const string message =
            "From: Sender <sender@example.test>\r\n" +
            "To: Receiver <receiver@example.test>\r\n" +
            "Subject: EMLX Reader contract\r\n" +
            "MIME-Version: 1.0\r\n" +
            "Content-Type: multipart/mixed; boundary=reader-boundary\r\n\r\n" +
            "--reader-boundary\r\nContent-Type: text/plain; charset=utf-8\r\n\r\nPortable body\r\n" +
            "--reader-boundary\r\nContent-Type: application/octet-stream; name=payload.bin\r\n" +
            "Content-Disposition: attachment; filename=payload.bin\r\n" +
            "Content-Transfer-Encoding: base64\r\n\r\nAQIDBA==\r\n" +
            "--reader-boundary--\r\n";
        return Encoding.ASCII.GetBytes(message);
    }

    private static byte[] CreateOlmArchive(IReadOnlyDictionary<string, byte[]> entries) {
        using var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true)) {
            foreach (KeyValuePair<string, byte[]> pair in entries) {
                ZipArchiveEntry entry = archive.CreateEntry(pair.Key, CompressionLevel.Optimal);
                using Stream output = entry.Open();
                output.Write(pair.Value, 0, pair.Value.Length);
            }
        }
        return stream.ToArray();
    }
}
