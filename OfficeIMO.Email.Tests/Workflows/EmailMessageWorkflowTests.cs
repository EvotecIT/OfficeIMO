using OfficeIMO.Email.Store;

namespace OfficeIMO.Email.Tests;

public sealed class EmailMessageWorkflowTests : IDisposable {
    private readonly string _root = Path.Combine(Path.GetTempPath(), "OfficeIMO.Email.Workflow." + Guid.NewGuid().ToString("N"));
    public EmailMessageWorkflowTests() => Directory.CreateDirectory(_root);
    public void Dispose() => Directory.Delete(_root, true);

    [Theory]
    [InlineData("eml")]
    [InlineData("msg")]
    public void FileViewReleasesSourceAndExportsSelectedPayloads(string extension) {
        string path = Path.Combine(_root, "invoice." + extension);
        CreateMessage("Invoice").Save(path);
        EmailMessage view = Assert.Single(EmailMessageReader.Read(path).Messages);
        Assert.Equal("Invoice", view.Subject);
        Assert.Equal("Invoice body", view.TextBody);
        Assert.Equal("alice@example.test", view.From!.Address);
        Assert.Equal(2, view.Attachments.Count);
        using (new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None)) { }
        var saved = view.SaveAttachments(Path.Combine(_root, "files"), new[] { 1 });
        var entry = Assert.Single(saved.Entries);
        Assert.Equal("1", entry.SourcePath);
        Assert.Equal(new byte[] { 1, 2, 3 }, File.ReadAllBytes(entry.OutputPath!));
        Assert.Equal("Invoice", view.Subject);
        Assert.False(view.Save(Path.Combine(_root, "copy." + extension)).HasErrors);
        Assert.False(view.Save(Path.Combine(_root, "portable.eml"), options: new EmailWriterOptions(EmailConversionLossPolicy.Warn)).HasErrors);
    }

    [Fact]
    public void ScopedPstReadsBodiesAndReturnedMessagesSurviveClosingScope() {
        string path = Path.Combine(_root, "archive.pst");
        using (var writer = EmailStorePstWriter.Create(path)) {
            string folder = writer.AddFolder("Inbox");
            writer.AddItem(folder, CreateMessage("Unrelated"));
            writer.AddItem(folder, CreateMessage("Invoice"));
            writer.Complete();
        }
        EmailMessage view;
        using (var scope = EmailMessageStore.Open(path)) {
            var limited = scope.Read(new EmailMessageQuery { SubjectContains = "Invoice", MaxItemsScanned = 1 });
            Assert.Empty(limited.Messages);
            Assert.True(limited.StoppedAtScanLimit);
            view = Assert.Single(scope.Read(new EmailMessageQuery { Folder = "Inbox", SubjectContains = "Invoice", First = 1 }).Messages);
            Assert.EndsWith("Inbox", view.FolderPath);
            Assert.Equal("Invoice body", view.TextBody);
        }
        using (new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None)) { }
        var saved = view.SaveAttachments(Path.Combine(_root, "pst-files"), new[] { 0, 1 });
        Assert.Equal(2, saved.Entries.Count);
        Assert.All(saved.Entries, e => Assert.NotNull(e.OutputPath));
    }

    [Fact]
    public void ChangedSourceFailsBeforeAnyAttachmentIsWritten() {
        string path = Path.Combine(_root, "message.eml");
        CreateMessage("Invoice").Save(path);
        EmailMessage view = Assert.Single(EmailMessageReader.Read(path).Messages);
        File.AppendAllText(path, "changed");
        string destination = Path.Combine(_root, "changed-files");
        Assert.Throws<IOException>(() => view.SaveAttachments(destination, new[] { 1 }));
        Assert.False(Directory.Exists(destination));
    }

    [Fact]
    public void PstSenderQueriesUseAddressesBeyondTheContentsTableRow() {
        string path = Path.Combine(_root, "senders.pst");
        var document = CreateMessage("Invoice");
        document.Sender = new EmailAddress("delegate@example.test", "Delegate");
        using (var writer = EmailStorePstWriter.Create(path)) {
            writer.AddItem(writer.AddFolder("Inbox"), document);
            writer.Complete();
        }
        using (var store = EmailStoreSession.Open(path)) {
            var reference = Assert.Single(store.EnumerateItems());
            var summary = store.ReadSummary(reference);
            Assert.Equal("alice@example.test", summary.From!.Address);
            Assert.Equal("delegate@example.test", summary.Sender!.Address);
            Assert.Single(store.Search(new EmailStoreQuery(senderContains: "alice@example.test")));
            Assert.Single(store.SearchWithReport(new EmailStoreQuery(senderContains: "delegate@example.test")).Results);
        }
        Assert.Single(EmailMessageReader.Read(path, new EmailMessageQuery {
            Folder = "Inbox", SenderContains = "alice@example.test", SubjectContains = "invoice",
            Since = new DateTimeOffset(2026, 10, 1, 0, 0, 0, TimeSpan.Zero),
            Before = new DateTimeOffset(2026, 11, 1, 0, 0, 0, TimeSpan.Zero)
        }).Messages);
    }

    [Fact]
    public void NativeExportCancellationAfterStagingStartsPreservesDestination() {
        string source = Path.Combine(_root, "large.eml");
        var document = CreateMessage("Large attachment");
        document.Attachments[1].Content = new byte[16 * 1024 * 1024];
        document.Attachments[1].Length = document.Attachments[1].Content!.Length;
        document.Save(source);
        var message = Assert.Single(EmailMessageReader.Read(source).Messages);
        string destination = Path.Combine(_root, "existing.msg");
        File.WriteAllText(destination, "keep existing destination");
        using var cancellation = new CancellationTokenSource();
        using var watcher = new FileSystemWatcher(_root, ".officeimo-*.tmp");
        watcher.Created += (_, _) => cancellation.Cancel();
        watcher.EnableRaisingEvents = true;
        Assert.ThrowsAny<OperationCanceledException>(() => message.Save(destination, overwrite: true,
            cancellationToken: cancellation.Token));
        Assert.Equal("keep existing destination", File.ReadAllText(destination));
        Assert.Empty(Directory.GetFiles(_root, ".officeimo-*.tmp"));
        using (new FileStream(source, FileMode.Open, FileAccess.ReadWrite, FileShare.None)) { }
    }

    [Fact]
    public void HtmlCopyRewritesCidAndKeepsRegularAttachmentBytes() {
        var document = CreateMessage("Invoice <2026>");
        document.Body.Html += "<script>alert(1)</script><img src=\"https://example.test/tracker.png\">";
        document.Body.Html += "<style>.banner{background-image:url(cid:logo)}</style>" +
            "<div style=\"background-image:url(cid:logo)\">Inline background</div>" +
            "<picture><source srcset=\"cid:logo 2x\"><img src=\"cid:logo\" srcset=\"cid:logo 1x, cid:logo 2x\"></picture>" +
            "<img src=\"data:image/png;base64,iVBORw0KGgo=\">";
        string path = Path.Combine(_root, "message.html");
        var result = EmailHtmlExporter.Export(document, path, includeAttachments: true);
        string html = File.ReadAllText(path);
        Assert.DoesNotContain("cid:logo", html);
        Assert.DoesNotContain("<script", html, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("https://example.test/tracker.png", html);
        Assert.Contains("Invoice &lt;2026&gt;", html);
        Assert.Contains("Attachments", html);
        Assert.Contains("invoice.bin", html);
        Assert.NotNull(result.AssetsDirectory);
        Assert.Contains("background-image:url(", html);
        Assert.Contains("srcset=\"message.assets-", html);
        Assert.Contains("data:image/png;base64,iVBORw0KGgo=", html);
        Assert.DoesNotContain(result.Diagnostics, d => d.Code == "EMAIL_HTML_RESOURCE_UNRESOLVED");
        Assert.Equal(2, result.Extraction!.Entries.Count);
        Assert.Contains(result.Extraction.Entries, e => e.OriginalName == "invoice.bin" &&
            File.ReadAllBytes(e.OutputPath!).SequenceEqual(new byte[] { 1, 2, 3 }));
        Assert.Throws<IOException>(() => EmailHtmlExporter.Export(document, path));
        string firstAssets = result.AssetsDirectory!;
        var replacement = EmailHtmlExporter.Export(document, path, overwrite: true);
        Assert.NotEqual(firstAssets, replacement.AssetsDirectory);
        Assert.True(Directory.Exists(firstAssets));
    }

    [Theory]
    [InlineData(EmailFileFormat.Eml)]
    [InlineData(EmailFileFormat.OutlookMsg)]
    public void HtmlContentEmbedsCidImagesAsDataUrisInMemory(EmailFileFormat format) {
        var source = CreateMessage("Invoice <2026>");
        source.Recipients.Add(new EmailRecipient(EmailRecipientKind.Cc, new EmailAddress("carol@example.test", "Carol")));
        source.Attachments[0].ContentId = "<LOGO>";
        source.Body.Html += "<script>alert(1)</script><img src=\"https://example.test/tracker.png\">" +
            "<div style=\"background-image:url(cid:logo)\">Inline background</div>" +
            "<img srcset=\"cid:logo 1x, cid:logo 2x\"><p>Résumé \U0001f642</p>";
        using var stream = new MemoryStream();
        Assert.False(source.Save(stream, format).HasErrors);
        stream.Position = 0;
        EmailDocument document = EmailDocument.Load(stream);
        string? authoredHtml = document.Body.Html;
        string[] authoredAttachmentIdentities = document.Attachments.Select(a =>
            a.FileName + "|" + a.ContentId + "|" + a.ContentType + "|" + a.Length).ToArray();

        EmailHtmlContent content = EmailHtmlExporter.ExportContent(document);

        string html = content.Html;
        Assert.Equal(authoredHtml, document.Body.Html);
        Assert.Equal(authoredAttachmentIdentities, document.Attachments.Select(a =>
            a.FileName + "|" + a.ContentId + "|" + a.ContentType + "|" + a.Length).ToArray());
        string png = Convert.ToBase64String(source.Attachments[0].Content!);
        Assert.Contains("<img src=\"data:image/png;base64," + png + "\"", html);
        Assert.Contains("background-image:url(&quot;data:image/png;base64," + png + "&quot;)", html);
        Assert.Contains("srcset=\"data:image/png;base64," + png + " 1x, data:image/png;base64," + png + " 2x\"", html);
        Assert.Contains("Résumé \U0001f642", html);
        Assert.DoesNotContain("cid:", html, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("<script", html, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("https://example.test/tracker.png", html);
        Assert.Contains("http-equiv=\"Content-Security-Policy\"", html);
        Assert.Contains("<h1>Invoice &lt;2026&gt;</h1>", html);
        Assert.Contains("From: Alice &lt;alice@example.test&gt;", html);
        Assert.Contains("Date: Thursday, 1 October 2026 12:00 +00:00", html);
        Assert.Contains("To: bob@example.test", html);
        Assert.Contains("Cc: Carol &lt;carol@example.test&gt;", html);
        Assert.DoesNotContain("invoice.bin", html);
        Assert.DoesNotContain(content.Diagnostics, _ => _.Code == "EMAIL_HTML_RESOURCE_UNRESOLVED");

        using var output = new MemoryStream();
        content.WriteTo(output);
        Assert.True(output.CanWrite);
        Assert.Equal(new UTF8Encoding(false).GetBytes(html), output.ToArray());
    }

    [Fact]
    public void HtmlContentReadsReferencedResourcesOnceAndClosesTheirStreams() {
        var document = CreateMessage("Repeated inline image");
        var inline = new ObservedContentSource(document.Attachments[0].Content!);
        var regular = new ObservedContentSource(new byte[] { 1, 2, 3 });
        document.Attachments[0].Content = null;
        document.Attachments[0].ContentSource = inline;
        document.Attachments[1].Content = null;
        document.Attachments[1].ContentSource = regular;
        document.Body.Html += "<img srcset=\"cid:logo 1x, cid:logo 2x\">" +
            "<div style=\"background-image:url(cid:logo)\"></div>";

        EmailHtmlContent content = EmailHtmlExporter.ExportContent(document);

        Assert.Contains("data:image/png;base64,", content.Html);
        Assert.DoesNotContain("cid:", content.Html, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(1, inline.OpenCount);
        Assert.Equal(1, inline.DisposeCount);
        Assert.Equal(0, regular.OpenCount);
    }

    [Fact]
    public void HtmlContentEncodesActualBytesWhenDeclaredLengthOverstatesContent() {
        var document = CreateMessage("Declared metadata");
        byte[] bytes = document.Attachments[0].Content!;
        var contentSource = new ObservedContentSource(bytes);
        document.Attachments[0].Content = null;
        document.Attachments[0].ContentSource = contentSource;
        document.Attachments[0].Length = 64L * 1024 * 1024;

        EmailHtmlContent content = EmailHtmlExporter.ExportContent(document);

        Assert.Contains("data:image/png;base64," + Convert.ToBase64String(bytes), content.Html);
        Assert.Equal(1, contentSource.OpenCount);
        Assert.Equal(1, contentSource.DisposeCount);
        Assert.Equal(64L * 1024 * 1024, document.Attachments[0].Length);
    }

    [Theory]
    [InlineData("eml")]
    [InlineData("msg")]
    public void HtmlContentFromMessageDoesNotRetainItsSource(string extension) {
        string path = Path.Combine(_root, "html-lifetime." + extension);
        Assert.False(CreateMessage("Scoped content").Save(path).HasErrors);
        EmailMessage message = Assert.Single(EmailMessageReader.Read(path).Messages);

        EmailHtmlContent content = EmailHtmlExporter.ExportContent(message);

        using (new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None)) { }
        File.Delete(path);
        using var destination = new MemoryStream();
        content.WriteTo(destination);
        Assert.True(destination.CanWrite);
        Assert.Equal(content.Html, Encoding.UTF8.GetString(destination.ToArray()));
        Assert.Contains("data:image/png;base64,", content.Html);
    }

    [Fact]
    public void FlatExtractionPrefixesPreventDifferentMessagesFromColliding() {
        string destination = Path.Combine(_root, "flat");
        var document = CreateMessage("Invoice");
        var first = EmailAttachmentExtractor.Extract(document, destination,
            new EmailAttachmentExtractionOptions(selectedAttachmentIndexes: new[] { 1 }, fileNamePrefix: "first"));
        var second = EmailAttachmentExtractor.Extract(document, destination,
            new EmailAttachmentExtractionOptions(selectedAttachmentIndexes: new[] { 1 }, fileNamePrefix: "second"));
        Assert.NotNull(Assert.Single(first.Entries).OutputPath);
        Assert.NotNull(Assert.Single(second.Entries).OutputPath);
        Assert.NotEqual(first.Entries[0].OutputPath, second.Entries[0].OutputPath);
    }

    [Fact]
    public void FlatNamesRetainIdentityAndFitPortableAtomicSegmentBounds() {
        string destination = Path.Combine(_root, "long-flat");
        var document = CreateMessage(new string('S', 110));
        document.Attachments[1].FileName = new string('\u754c', 90) + ".bin";
        string firstSource = Path.Combine(_root, "first.eml");
        string secondSource = Path.Combine(_root, "second.eml");
        document.Save(firstSource);
        document.Save(secondSource);
        var first = Assert.Single(EmailMessageReader.Read(firstSource).Messages);
        var second = Assert.Single(EmailMessageReader.Read(secondSource).Messages);
        var firstFile = Assert.Single(first.SaveAttachments(destination, new[] { 1 }, fileNamePrefix: first.ExportName).Entries);
        var secondFile = Assert.Single(second.SaveAttachments(destination, new[] { 1 }, fileNamePrefix: second.ExportName).Entries);
        Assert.NotNull(firstFile.OutputPath);
        Assert.NotNull(secondFile.OutputPath);
        Assert.NotEqual(firstFile.OutputPath, secondFile.OutputPath);
        Assert.All(new[] { firstFile, secondFile }, entry => {
            string name = Path.GetFileName(entry.OutputPath!);
            Assert.True(System.Text.Encoding.UTF8.GetByteCount(name) <= 218);
            Assert.EndsWith(".bin", name);
            Assert.Equal(new byte[] { 1, 2, 3 }, File.ReadAllBytes(entry.OutputPath!));
        });
    }

    [Fact]
    public void ExtractionTreatsAttachmentNamesAsTextOnEveryRuntime() {
        var document = CreateMessage("Untrusted names");
        document.Attachments[1].FileName = "<invoice>\"\0.bin";
        var entry = Assert.Single(EmailAttachmentExtractor.Extract(document, Path.Combine(_root, "unsafe-names"),
            new EmailAttachmentExtractionOptions(selectedAttachmentIndexes: new[] { 1 }, fileNamePrefix: "message")).Entries);
        Assert.NotNull(entry.OutputPath);
        Assert.Equal(new byte[] { 1, 2, 3 }, File.ReadAllBytes(entry.OutputPath!));
        Assert.EndsWith(".bin", entry.OutputPath);
        Assert.Equal("<invoice>\"\0.bin", entry.OriginalName);
    }

    private sealed class ObservedContentSource : IEmailContentSource {
        private readonly byte[] _bytes;
        internal ObservedContentSource(byte[] bytes) { _bytes = bytes; }
        public long? Length => _bytes.LongLength;
        internal int OpenCount { get; private set; }
        internal int DisposeCount { get; private set; }
        public Stream OpenRead() {
            OpenCount++;
            return new ObservedStream(_bytes, () => DisposeCount++);
        }
        public Task<Stream> OpenReadAsync(CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            return Task.FromResult(OpenRead());
        }
        private sealed class ObservedStream : MemoryStream {
            private readonly Action _closed;
            private bool _disposed;
            internal ObservedStream(byte[] bytes, Action closed) : base(bytes, writable: false) { _closed = closed; }
            protected override void Dispose(bool disposing) {
                if (disposing && !_disposed) { _disposed = true; _closed(); }
                base.Dispose(disposing);
            }
        }
    }

    private static EmailDocument CreateMessage(string subject) {
        var document = new EmailDocument { Subject = subject, From = new EmailAddress("alice@example.test", "Alice"),
            Date = new DateTimeOffset(2026, 10, 1, 12, 0, 0, TimeSpan.Zero), MessageId = "invoice@example.test" };
        document.Recipients.Add(new EmailRecipient(EmailRecipientKind.To, new EmailAddress("bob@example.test")));
        document.Body.Text = "Invoice body";
        document.Body.Html = "<p>Invoice body</p><img src=\"cid:logo\">";
        byte[] png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jf0sAAAAASUVORK5CYII=");
        document.Attachments.Add(new EmailAttachment { FileName = "logo.png", ContentType = "image/png", ContentId = "logo", IsInline = true, Content = png, Length = png.Length });
        document.Attachments.Add(new EmailAttachment { FileName = "invoice.bin", ContentType = "application/octet-stream", Content = new byte[] { 1, 2, 3 }, Length = 3 });
        return document;
    }
}
