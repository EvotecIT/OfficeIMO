using OfficeIMO.Email;
using OfficeIMO.Email.Store;
using OfficeIMO.Pdf;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace OfficeIMO.Workflows.Tests;

public sealed class EmailEvidenceWorkflowTests {
    [Fact]
    public void ConversationRejectsFirstOversizedBodyBeforeReadingTheNextBody() {
        using var scope = new MailScope();
        var first = Message("first@example.test", 1); first.Body.Text = "12345";
        var next = Message("next@example.test", 2); next.Body.Text = new string('x', 1000);
        scope.Write("a.eml", first); scope.Write("b.eml", next);
        using var session = EmailStoreSession.Open(scope.Root);
        var exceeded = Assert.Throws<ArgumentException>(() => EmailEvidenceWorkflow.CreateConversation(session, session.EnumerateItems().First().Key,
            new EmailEvidenceOptions { IncludePdf = false, MaxBodySourceCharacters = 4,
                ReaderOptions = new EmailReaderOptions(maxDecodedPropertyBytes: 10) }));
        Assert.Contains("MaxSourceChars", exceeded.Message);
    }

    [Fact]
    public void PstDossierClassifiesEmbeddedMessageFromMetadataWithoutReadingItsBody() {
        using var scope = new MailScope();
        string path = Path.Combine(scope.Root, "source.pst");
        var message = Message("parent@example.test", 1);
        var embedded = Message("embedded@example.test", 2);
        message.Attachments.Add(new EmailAttachment { FileName = "forwarded.msg", EmbeddedDocument = embedded });
        using (var writer = EmailStorePstWriter.Create(path)) {
            var folder = writer.AddFolder("Inbox");
            writer.AddItem(folder, message);
            writer.Complete();
        }
        using var session = EmailStoreSession.Open(path);
        var reference = Assert.Single(session.EnumerateItems());
        var metadata = session.ReadItem(reference, new EmailStoreItemReadOptions(EmailStoreItemReadParts.AttachmentMetadata));
        Assert.Null(Assert.Single(metadata.Document.Attachments).EmbeddedDocument);
        var result = EmailEvidenceWorkflow.CreateConversation(session, reference.Key, new EmailEvidenceOptions { IncludePdf = false });
        Assert.True(Assert.Single(result.Manifest.Messages[0].Attachments).EmbeddedMessage);
        Assert.Contains("Embedded message: True", result.Html);
    }

    [Fact]
    public void EvidenceZipReopensWithSourceAndAttachmentHashesReadablePdfAndEscapedHtml() {
        using var scope = new MailScope();
        var message = new EmailDocument { Subject = "Evidence <script>bad()</script>", From = new EmailAddress("sender@example.test") };
        message.Body.Html = "<p>Readable evidence text</p><script>bad()</script><img src='https://outside.example.test/tracker.png'>";
        byte[] attachment = Encoding.UTF8.GetBytes("attachment bytes");
        message.Attachments.Add(new EmailAttachment { FileName = "../notes.txt", Content = attachment, ContentType = "text/plain" });
        string path = scope.Write("source.eml", message);
        byte[] original = File.ReadAllBytes(path);
        var result = EmailEvidenceWorkflow.Create(path);
        Assert.Equal(Convert.ToHexString(SHA256.HashData(original)).ToLowerInvariant(), result.Manifest.SourceFingerprint);
        Assert.Equal(Convert.ToHexString(SHA256.HashData(attachment)).ToLowerInvariant(), result.Manifest.Messages[0].Attachments[0].Sha256);
        Assert.Equal("Unverified", result.Manifest.Messages[0].IntegrityStatus);
        Assert.Contains("Readable evidence text", PdfDocument.Load(result.PdfBytes!).Reader.Text());
        using var archive = new ZipArchive(new MemoryStream(result.ToZipBytes()), ZipArchiveMode.Read);
        Assert.Equal(new[] { "manifest.json", "report.html", "report.md", "report.pdf" }, archive.Entries.Select(entry => entry.FullName).OrderBy(value => value).ToArray());
        using var htmlReader = new StreamReader(archive.GetEntry("report.html")!.Open());
        string html = htmlReader.ReadToEnd();
        Assert.DoesNotContain("<script>", html); Assert.DoesNotContain("tracker.png", html);
        Assert.Contains("&lt;script&gt;", html);
        using var manifestReader = new StreamReader(archive.GetEntry("manifest.json")!.Open());
        using var manifest = JsonDocument.Parse(manifestReader.ReadToEnd());
        Assert.Equal(result.Manifest.SourceFingerprint, manifest.RootElement.GetProperty("SourceFingerprint").GetString());
        Assert.Equal(original, File.ReadAllBytes(path));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void DossierPreservesChronologyAuthoritativeLinksAndMissingOrAmbiguousParents(bool duplicateParent) {
        using var scope = new MailScope();
        var root = Message("root@example.test", 3);
        var reply = Message("reply@example.test", 4); reply.MessageMetadata.InReplyToId = "root@example.test";
        var other = Message(duplicateParent ? "root@example.test" : "orphan@example.test", 1);
        if (!duplicateParent) other.MessageMetadata.InReplyToId = "missing@example.test";
        scope.Write("a.eml", root); scope.Write("b.eml", reply); scope.Write("c.eml", other);
        using var session = EmailStoreSession.Open(scope.Root);
        var selected = session.EnumerateItems().First().Key;
        var result = EmailEvidenceWorkflow.CreateConversation(session, selected, new EmailEvidenceOptions { IncludePdf = false });
        Assert.True(result.Manifest.GraphComplete);
        Assert.Equal(duplicateParent ? 3 : 2, result.Manifest.Messages.Count);
        Assert.Equal(duplicateParent ? new[] { 1, 3, 4 } : new[] { 3, 4 }, result.Manifest.Messages.Select(message => message.Date!.Value.Day).ToArray());
        if (duplicateParent) {
            Assert.Contains(result.Manifest.MissingParents, parent => parent.Reason == "AmbiguousParent");
            Assert.Contains("AmbiguousParent", result.Html);
        } else {
            Assert.Contains(result.Manifest.ThreadLinks, link => !link.IsHeuristic && link.Reasons.Contains("InReplyTo"));
            var orphan = EmailEvidenceWorkflow.CreateConversation(session, session.EnumerateItems().Last().Key,
                new EmailEvidenceOptions { IncludePdf = false });
            Assert.Contains(orphan.Manifest.MissingParents, parent => parent.Reason == "MissingParent");
            Assert.Contains("MissingParent", orphan.Html);
        }
        Assert.Throws<InvalidDataException>(() => EmailEvidenceWorkflow.CreateConversation(session, selected,
            new EmailEvidenceOptions { IncludePdf = false, MaxMessages = 1 }));
    }

    [Fact]
    public void BoundedDossierDisclosesIncompleteGraphAndBodyClippingAndRejectsOversizedOutput() {
        using var scope = new MailScope();
        scope.Write("a.eml", Message("first@example.test", 1)); scope.Write("b.eml", Message("second@example.test", 2));
        using var session = EmailStoreSession.Open(scope.Root);
        var result = EmailEvidenceWorkflow.CreateConversation(session, session.EnumerateItems().First().Key,
            new EmailEvidenceOptions { IncludePdf = false, MaxItemsScanned = 1, MaxBodyTextCharacters = 3 });
        Assert.False(result.Manifest.GraphComplete); Assert.True(result.Manifest.Messages[0].BodyTextTruncated);
        Assert.Contains("Incomplete", result.Html);
        Assert.Null(result.PdfBytes);
        Assert.Throws<InvalidDataException>(() => EmailEvidenceWorkflow.Create(Path.Combine(scope.Root, "a.eml"),
            new EmailEvidenceOptions { IncludePdf = false, MaxOutputBytes = 1 }));
        Assert.Throws<OperationCanceledException>(() => result.ToZipBytes(new CancellationToken(true)));
    }

    private static EmailDocument Message(string id, int day) {
        var document = new EmailDocument { Subject = "Discussion", MessageId = id, Date = new DateTimeOffset(2026, 1, day, 9, 0, 0, TimeSpan.Zero) };
        document.Body.Text = "Discussion body";
        return document;
    }
    private sealed class MailScope : IDisposable {
        internal string Root { get; } = Path.Combine(Path.GetTempPath(), "officeimo-evidence-" + Guid.NewGuid().ToString("N"));
        internal MailScope() => Directory.CreateDirectory(Root);
        internal string Write(string name, EmailDocument message) { string path = Path.Combine(Root, name); File.WriteAllBytes(path, message.ToBytes()); return path; }
        public void Dispose() => Directory.Delete(Root, recursive: true);
    }
}
