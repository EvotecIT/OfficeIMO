using MimeKit;
using OfficeIMO.Email;
using System.Security.Cryptography;
using Xunit;

namespace OfficeIMO.Email.Tests;

/// <summary>Explicit independently-produced MIME evidence; MimeKit remains a test-only dependency.</summary>
public sealed class MimeKitAdversarialInteropTests {
    private static readonly IReadOnlyDictionary<string, string[]> ScenarioFiles =
        new Dictionary<string, string[]>(StringComparer.Ordinal) {
            ["malformed-recovery"] = new[] {
                "messages/missing-subtype.txt", "messages/empty-multipart.txt",
                "messages/delivery-status-no-blank-line.txt"
            },
            ["nested-multiparts"] = new[] { "messages/multipart-digest.txt", "messages/rfc2060.txt" },
            ["legacy-encoding"] = new[] { "messages/japanese.txt" },
            ["embedded-messages"] = new[] { "messages/bounce.txt", "messages/rfc2060.txt" },
            ["reports"] = new[] {
                "messages/delivery-status.txt", "messages/disposition-notification.txt"
            },
            ["protected-entities"] = new[] { "smime/thunderbird-signed.txt" }
        };

    [ProducerCorpusFact("MimeKit")]
    public void PinnedProducerCorpus_covers_the_adversarial_scenario_matrix() {
        string repository = Assert.IsType<string>(ExternalEmailCorpusHarness.FindRepository("MimeKit"));
        string data = Path.Combine(repository, "UnitTests", "TestData");
        foreach (KeyValuePair<string, string[]> scenario in ScenarioFiles) {
            Assert.NotEmpty(scenario.Value);
            foreach (string relativePath in scenario.Value) {
                string path = Path.Combine(data, relativePath.Replace('/', Path.DirectorySeparatorChar));
                Assert.True(File.Exists(path), scenario.Key + ": missing " + relativePath);
                using (SHA256 sha = SHA256.Create())
                using (FileStream stream = File.OpenRead(path)) {
                    Assert.Equal(32, sha.ComputeHash(stream).Length);
                }
                bool matched = scenario.Key == "malformed-recovery"
                    ? ExternalEmailCorpusHarness.ValidateMalformedMimeArtifact(path)
                    : ExternalEmailCorpusHarness.ValidateMimeArtifact(path);
                Assert.True(matched,
                    scenario.Key + ": semantic comparison did not apply to " + relativePath);
            }
        }
    }

    [ProducerCorpusFact("MimeKit")]
    public void Producer_fixture_preserves_duplicate_header_order() {
        string repository = Assert.IsType<string>(ExternalEmailCorpusHarness.FindRepository("MimeKit"));
        string path = Path.Combine(repository, "UnitTests", "TestData", "messages", "delivery-status.txt");
        using MimeMessage oracle = MimeMessage.Load(path);
        string[] oracleReceived = oracle.Headers.Where(header => header.Field.Equals("Received",
                StringComparison.OrdinalIgnoreCase)).Select(header => header.Value).ToArray();
        using EmailReadResult read = new EmailDocumentReader().Read(path);
        string[] officeReceived = read.Document.Headers.Where(header => header.Name.Equals("Received",
                StringComparison.OrdinalIgnoreCase)).Select(header => header.Value).ToArray();

        Assert.True(oracleReceived.Length > 1);
        Assert.Equal(oracleReceived.Length, officeReceived.Length);
        Assert.Contains("maleman.mcom.com", officeReceived[0], StringComparison.OrdinalIgnoreCase);
        Assert.Contains("mm1", officeReceived[officeReceived.Length - 1],
            StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void MimeKit_generated_Rfc2231_continuations_and_embedded_message_reopen_semantically() {
        var embedded = new MimeMessage();
        embedded.From.Add(new MailboxAddress("Nested sender", "nested@example.test"));
        embedded.To.Add(new MailboxAddress("Nested recipient", "recipient@example.test"));
        embedded.Subject = "Nested producer message";
        embedded.Body = new TextPart("plain") { Text = "Nested body" };
        var attachment = new MimePart("application", "octet-stream") {
            Content = new MimeContent(new MemoryStream(new byte[] { 1, 2, 3, 4 })),
            ContentDisposition = new ContentDisposition(ContentDisposition.Attachment),
            FileName = string.Concat(Enumerable.Repeat("Zażółć-gęślą-", 12)) + ".bin"
        };
        var root = new Multipart("mixed") {
            new TextPart("plain") { Text = "Outer body" },
            new MessagePart("rfc822") { Message = embedded },
            attachment
        };
        var message = new MimeMessage();
        message.From.Add(new MailboxAddress("Producer", "producer@example.test"));
        message.To.Add(new MailboxAddress("Consumer", "consumer@example.test"));
        message.Subject = "Producer continuations";
        message.Headers.Add("X-Trace", "first");
        message.Headers.Add("X-Trace", "second");
        message.Body = root;
        var format = FormatOptions.Default.Clone();
        format.ParameterEncodingMethod = ParameterEncodingMethod.Rfc2231;
        format.MaxLineLength = 78;
        byte[] produced;
        using (var stream = new MemoryStream()) {
            message.WriteTo(format, stream);
            produced = stream.ToArray();
        }
        string wire = Encoding.ASCII.GetString(produced);
        Assert.Contains("filename*0*=", wire, StringComparison.OrdinalIgnoreCase);

        using EmailReadResult read = new EmailDocumentReader().Read(produced);
        Assert.Equal(new[] { "first", "second" }, read.Document.Headers
            .Where(header => header.Name == "X-Trace").Select(header => header.Value).ToArray());
        EmailAttachment nested = Assert.Single(read.Document.Attachments,
            value => value.EmbeddedDocument != null);
        Assert.Equal("Nested producer message", nested.EmbeddedDocument?.Subject);
        Assert.Equal("Nested body", nested.EmbeddedDocument?.Body.Text?.TrimEnd());
        EmailAttachment file = Assert.Single(read.Document.Attachments,
            value => value.EmbeddedDocument == null);
        Assert.Equal(attachment.FileName, file.FileName);
        Assert.Equal(new byte[] { 1, 2, 3, 4 }, file.Content);
    }
}
