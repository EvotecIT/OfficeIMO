using OfficeIMO.Email;
using OfficeIMO.Email.Store;
using OfficeIMO.Reader.Email;
using OfficeIMO.Reader.Html;
using OfficeIMO.Reader.Rtf;
using OfficeIMO.Rtf;
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
        Assert.Contains(OfficeDocumentReaderBuilderEmailStoreExtensions.HandlerId, result.CapabilitiesUsed);
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

    [Fact]
    public void Selective_query_projects_only_matching_store_items() {
        const string xml = "<emails>" +
            "<email><OPFMessageCopySubject>Keep this message</OPFMessageCopySubject>" +
            "<OPFMessageCopyBody>Selected body</OPFMessageCopyBody></email>" +
            "<email><OPFMessageCopySubject>Skip this message</OPFMessageCopySubject>" +
            "<OPFMessageCopyBody>Unselected body</OPFMessageCopyBody></email>" +
            "</emails>";
        byte[] archive = CreateOlmArchive(new Dictionary<string, byte[]> {
            ["Local/com.microsoft.__Messages/Inbox/messages.xml"] = Encoding.UTF8.GetBytes(xml)
        });
        var adapterOptions = new ReaderEmailStoreOptions {
            Query = new EmailStoreQuery(subjectContains: "keep"),
            MaxItems = 10
        };
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddEmailStoreHandler(adapterOptions)
            .Build();

        OfficeDocumentReadResult result = reader.ReadDocument(archive, "mailbox.olm");

        Assert.Contains(result.Chunks, chunk => chunk.Text.Contains("Keep this message", StringComparison.Ordinal));
        Assert.DoesNotContain(result.Chunks, chunk => chunk.Text.Contains("Skip this message", StringComparison.Ordinal));
        Assert.Contains(result.Metadata, item => item.Name == "ItemCount" && item.Value == "1");
        Assert.Contains(result.Metadata,
            item => item.Name == "SelectionLimitReached" && item.Value == "False");
    }

    [Fact]
    public void Reader_item_bound_is_reported_instead_of_materializing_the_whole_store() {
        const string xml = "<emails>" +
            "<email><OPFMessageCopySubject>First</OPFMessageCopySubject></email>" +
            "<email><OPFMessageCopySubject>Second</OPFMessageCopySubject></email>" +
            "</emails>";
        byte[] archive = CreateOlmArchive(new Dictionary<string, byte[]> {
            ["Local/com.microsoft.__Messages/Inbox/messages.xml"] = Encoding.UTF8.GetBytes(xml)
        });
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddEmailStoreHandler(new ReaderEmailStoreOptions { MaxItems = 1 })
            .Build();

        OfficeDocumentReadResult result = reader.ReadDocument(archive, "mailbox.olm");

        Assert.Contains(result.Metadata,
            item => item.Name == "SelectionLimitReached" && item.Value == "True");
        Assert.Contains(result.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_STORE_READER_SELECTION_LIMIT");
        Assert.Contains(result.Chunks, chunk => chunk.Text.Contains("First", StringComparison.Ordinal));
        Assert.DoesNotContain(result.Chunks, chunk => chunk.Text.Contains("Second", StringComparison.Ordinal));
    }

    [Fact]
    public void Store_source_hash_is_opt_in_while_chunk_hashes_remain_available() {
        byte[] emlx = CreateEmlx(CreateMultipartMessage(), null);
        string path = Path.Combine(Path.GetTempPath(), "officeimo-reader-store-" + Guid.NewGuid().ToString("N") + ".emlx");
        try {
            File.WriteAllBytes(path, emlx);
            OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
                .AddEmailStoreHandler()
                .Build();

            OfficeDocumentReadResult result = reader.ReadDocument(
                path, new ReaderOptions { ComputeHashes = true });

            Assert.Null(result.Source.SourceHash);
            Assert.All(result.Chunks, chunk => Assert.False(string.IsNullOrWhiteSpace(chunk.ChunkHash)));
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void Item_at_a_time_reader_uses_configured_semantic_body_handlers() {
        string folder = Path.Combine(Path.GetTempPath(),
            "officeimo-reader-store-items-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        try {
            var html = new EmailDocument { Subject = "HTML item" };
            html.Body.Html = "<html><body><p>Visible semantic HTML</p>" +
                "<script>hidden-script-marker</script></body></html>";
            File.WriteAllBytes(Path.Combine(folder, "01-html.eml"),
                new EmailDocumentWriter().ToBytes(html, EmailFileFormat.Eml));

            RtfDocument rtf = RtfDocument.Create();
            rtf.AddParagraph("Visible semantic RTF");
            var rtfMessage = new EmailDocument { Subject = "RTF item" };
            rtfMessage.Body.Rtf = rtf.ToRtf();
            File.WriteAllBytes(Path.Combine(folder, "02-rtf.eml"),
                new EmailDocumentWriter().ToBytes(rtfMessage, EmailFileFormat.Eml));

            OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
                .AddHtmlHandler()
                .AddRtfHandler()
                .AddEmailStoreHandler()
                .Build();

            ReaderEmailStoreItemResult[] results = reader.ReadEmailStoreItems(
                folder, emailStoreOptions: new ReaderEmailStoreOptions { MaxItems = 10 }).ToArray();

            Assert.Equal(2, results.Length);
            Assert.All(results, result => Assert.True(result.Succeeded));
            Assert.Contains(results.SelectMany(result => result.Chunks), chunk =>
                chunk.Kind == ReaderInputKind.Html &&
                chunk.Location.SourceBlockKind == "email-body-html" &&
                chunk.Text.Contains("Visible semantic HTML", StringComparison.Ordinal));
            Assert.DoesNotContain(results.SelectMany(result => result.Chunks), chunk =>
                chunk.Text.Contains("hidden-script-marker", StringComparison.Ordinal));
            Assert.Contains(results.SelectMany(result => result.Chunks), chunk =>
                chunk.Kind == ReaderInputKind.Rtf &&
                chunk.Location.SourceBlockKind == "email-body-rtf" &&
                chunk.Text.Contains("Visible semantic RTF", StringComparison.Ordinal));
            string[] chunkIds = results.SelectMany(result => result.Chunks)
                .Select(chunk => chunk.Id).ToArray();
            Assert.Equal(chunkIds.Length, chunkIds.Distinct(StringComparer.Ordinal).Count());
        } finally {
            Directory.Delete(folder, recursive: true);
        }
    }

    [Fact]
    public void Item_success_is_not_negated_by_store_wide_diagnostics() {
        byte[] archive = CreateOlmArchive(new Dictionary<string, byte[]> {
            ["Local/com.microsoft.__Messages/Inbox/valid.xml"] = Encoding.UTF8.GetBytes(
                "<emails><email><OPFMessageCopySubject>Valid item</OPFMessageCopySubject>" +
                "<OPFMessageCopyBody>Valid body</OPFMessageCopyBody></email></emails>"),
            ["Local/com.microsoft.__Messages/Inbox/invalid.xml"] = Encoding.UTF8.GetBytes(
                "<appointments><appointment>")
        });
        string path = Path.Combine(Path.GetTempPath(),
            "officeimo-reader-store-diagnostics-" + Guid.NewGuid().ToString("N") + ".olm");
        try {
            File.WriteAllBytes(path, archive);
            OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
                .AddEmailStoreHandler()
                .Build();

            ReaderEmailStoreItemResult result = Assert.Single(reader.ReadEmailStoreItems(path));

            Assert.True(result.Succeeded);
            Assert.Contains(result.StoreDiagnostics,
                diagnostic => diagnostic.Code == "EMAIL_STORE_OLM_XML_INVALID" &&
                    diagnostic.Severity == EmailDiagnosticSeverity.Error);
            Assert.DoesNotContain(result.ItemDiagnostics,
                diagnostic => diagnostic.Severity == EmailDiagnosticSeverity.Error);
            Assert.Equal(result.StoreDiagnostics.Count + result.ItemDiagnostics.Count,
                result.Diagnostics.Count);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
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
