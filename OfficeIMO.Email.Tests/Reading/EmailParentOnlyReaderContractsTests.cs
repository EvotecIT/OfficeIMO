using OfficeIMO.Email.Store;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Email;

namespace OfficeIMO.Email.Tests;

public sealed class EmailParentOnlyReaderContractsTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task MailboxSpecificMessageLimitSurvivesDirectMessageDefaults(bool asynchronous) {
        byte[] mailbox = Encoding.ASCII.GetBytes("From sender@example.test Wed Sep 30 12:00:00 2026\nSubject: bounded\n\n" + new string('x', 100));
        var reader = new OfficeDocumentReaderBuilder().AddEmailHandler(new ReaderEmailOptions {
            MailboxOptions = new EmailMailboxReaderOptions(new EmailReaderOptions(maxInputBytes: 32))
        }).Build();
        using var source = new MemoryStream(mailbox);
        EmailLimitExceededException exception = asynchronous
            ? await Assert.ThrowsAsync<EmailLimitExceededException>(() => reader.ReadDocumentAsync(source, "mail.mbox"))
            : Assert.Throws<EmailLimitExceededException>(() => reader.ReadDocument(source, "mail.mbox"));
        Assert.Equal(nameof(EmailReaderOptions.MaxInputBytes), exception.LimitName);
    }

    [Theory]
    [InlineData("emlx")]
    [InlineData("mbox")]
    public async Task ParentOnlyPolicyExcludesOpaqueMailArtifactsAcrossProjectionRoutes(string extension) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-parent-artifact-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            var child = new EmailDocument();
            child.Body.Text = "opaque child scope marker";
            byte[] childPayload = child.ToBytes();
            byte[] artifact = Encoding.ASCII.GetBytes(extension == "emlx" ? childPayload.Length + "\n"
                : "From sender@example.test Wed Sep 30 12:00:00 2026\n").Concat(childPayload).ToArray();
            var parent = new EmailDocument();
            parent.Body.Text = "parent scope marker";
            parent.Attachments.Add(new EmailAttachment {
                FileName = "child." + extension, ContentType = "application/octet-stream", Content = artifact
            });
            byte[] payload = parent.ToBytes();
            var defaultReader = new OfficeDocumentReaderBuilder().AddEmailHandler().AddEmailStoreHandler().Build();
            using (var source = new MemoryStream(payload)) {
                Assert.Contains(defaultReader.ReadDocument(source, "parent.eml").Chunks,
                    chunk => chunk.Text.Contains("opaque child scope marker"));
            }
            var policy = new ReaderEmailStoreOptions { ItemReadOptions = new EmailStoreItemReadOptions(
                EmailStoreItemReadParts.All & ~EmailStoreItemReadParts.EmbeddedItems) };
            // The nested store handler keeps its defaults; exclusion belongs to the parent projection.
            var reader = new OfficeDocumentReaderBuilder().AddEmailHandler(new ReaderEmailOptions {
                MessageOptions = new EmailReaderOptions(includeEmbeddedMessages: false)
            }).AddEmailStoreHandler().Build();
            using (var source = new MemoryStream(payload)) AssertParentOnly(reader.ReadDocument(source, "parent.eml").Chunks);
            using (var source = new MemoryStream(payload)) AssertParentOnly((await reader.ReadDocumentAsync(source, "parent.eml")).Chunks);
            File.WriteAllBytes(Path.Combine(root, "parent.eml"), payload);
            AssertParentOnly(Assert.Single(EmailStoreItemReader.Read(reader, root, emailStoreOptions: policy)).Chunks);
            var aggregateReader = new OfficeDocumentReaderBuilder().AddEmailHandler().AddEmailStoreHandler(policy).Build();
            byte[] emlx = Encoding.ASCII.GetBytes(payload.Length + "\n").Concat(payload).ToArray();
            using var aggregateSource = new MemoryStream(emlx);
            AssertParentOnly(aggregateReader.ReadDocument(aggregateSource, "parent.emlx").Chunks);
        } finally { Directory.Delete(root, recursive: true); }
    }

    private static void AssertParentOnly(IReadOnlyList<ReaderChunk> chunks) {
        Assert.Contains(chunks, chunk => chunk.Text.Contains("parent scope marker"));
        Assert.DoesNotContain(chunks, chunk => chunk.Text.Contains("opaque child scope marker"));
        Assert.Contains(chunks, chunk => chunk.Text.Contains("child."));
        Assert.Contains(chunks, chunk => chunk.Location.SourceBlockKind == "email-attachment");
    }
}
