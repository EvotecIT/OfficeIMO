using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class EmailSemanticComparerTests {
    [Fact]
    public void IdenticalPortableDocumentsProduceStableFingerprint() {
        EmailDocument source = CreateDocument(Encoding.UTF8.GetBytes("attachment payload"));
        EmailDocument destination = CreateDocument(Encoding.UTF8.GetBytes("attachment payload"));

        EmailSemanticComparisonReport report = EmailSemanticComparer.Compare(source, destination);

        Assert.True(report.IsMatch);
        Assert.Empty(report.Differences);
        Assert.Equal("SHA-256", report.Source.Algorithm);
        Assert.Equal(report.Source.HexDigest, report.Destination.HexDigest);
        Assert.Equal(1, report.Source.RecipientCount);
        Assert.Equal(1, report.Source.AttachmentCount);
        Assert.Equal(18, report.Source.AttachmentBytesHashed);
        Assert.True(report.Source.EntryCount > 10);
    }

    [Fact]
    public void DifferencesExposePathsAndLengthsWithoutValues() {
        EmailDocument source = CreateDocument(Encoding.UTF8.GetBytes("private source bytes"));
        EmailDocument destination = CreateDocument(Encoding.UTF8.GetBytes("different"));
        destination.Subject = "private destination subject";

        EmailSemanticComparisonReport report = EmailSemanticComparer.Compare(source, destination);

        Assert.False(report.IsMatch);
        Assert.Contains(report.Differences, difference =>
            difference.Path.Contains("attachments/00000000/content", StringComparison.Ordinal));
        Assert.DoesNotContain(report.Differences, difference =>
            difference.Path.Contains("private source bytes", StringComparison.Ordinal) ||
            difference.Path.Contains("private destination subject", StringComparison.Ordinal));
    }

    [Fact]
    public void Difference_paths_digest_arbitrary_named_property_names() {
        const string privatePropertyName = "private-property-name-do-not-disclose";
        EmailDocument source = CreateDocument(Array.Empty<byte>());
        EmailDocument destination = CreateDocument(Array.Empty<byte>());
        source.MapiProperties.Add(new MapiProperty(0x8001, MapiPropertyType.Unicode, "source",
            name: new MapiNamedProperty(Guid.NewGuid(), privatePropertyName)));

        EmailSemanticComparisonReport report = EmailSemanticComparer.Compare(source, destination);

        Assert.False(report.IsMatch);
        Assert.DoesNotContain(report.Differences, difference =>
            difference.Path.Contains(privatePropertyName, StringComparison.Ordinal));
        Assert.Contains(report.Differences, difference =>
            difference.Path.Contains("name-digest-", StringComparison.Ordinal));
    }

    [Fact]
    public void MigrationProfileNormalizesStoreIdentityWhileStrictProfileDetectsIt() {
        EmailDocument source = CreateDocument(Array.Empty<byte>());
        EmailDocument destination = CreateDocument(Array.Empty<byte>());
        source.Format = EmailFileFormat.OutlookMsg;
        destination.Format = EmailFileFormat.Tnef;
        source.MapiProperties.Add(new MapiProperty(0x0FFF, MapiPropertyType.Binary, new byte[] { 1, 2, 3 }));
        destination.MapiProperties.Add(new MapiProperty(0x0FFF, MapiPropertyType.Binary, new byte[] { 9, 8, 7 }));

        EmailSemanticComparisonReport migration = EmailSemanticComparer.Compare(source, destination);
        EmailSemanticComparisonReport strict = EmailSemanticComparer.Compare(source, destination,
            new EmailSemanticComparisonOptions(EmailSemanticComparisonProfile.Strict));

        Assert.True(migration.IsMatch);
        Assert.False(strict.IsMatch);
        Assert.Contains(strict.Differences, difference => difference.Path.EndsWith("/strict/format", StringComparison.Ordinal) ||
            difference.Path.Contains("p-0FFF", StringComparison.Ordinal));
    }

    [Fact]
    public void KeyedFingerprintsAreStableAndDistinctFromUnkeyedDigests() {
        EmailDocument document = CreateDocument(Encoding.UTF8.GetBytes("secret"));
        byte[] key = Enumerable.Range(1, 32).Select(value => checked((byte)value)).ToArray();
        var options = new EmailSemanticComparisonOptions(digestKey: key);

        EmailSemanticFingerprint first = EmailSemanticComparer.CreateFingerprint(document, options);
        EmailSemanticFingerprint second = EmailSemanticComparer.CreateFingerprint(document, options);
        EmailSemanticFingerprint unkeyed = EmailSemanticComparer.CreateFingerprint(document);

        Assert.Equal("HMAC-SHA-256", first.Algorithm);
        Assert.Equal(first.HexDigest, second.HexDigest);
        Assert.NotEqual(first.HexDigest, unkeyed.HexDigest);
    }

    [Fact]
    public async Task AsyncComparisonStreamsReopenableAttachmentSources() {
        byte[] payload = Enumerable.Range(0, 100_000).Select(value => unchecked((byte)value)).ToArray();
        var sourceContent = new CountingContentSource(payload);
        var destinationContent = new CountingContentSource(payload);
        EmailDocument source = CreateDocument(Array.Empty<byte>());
        EmailDocument destination = CreateDocument(Array.Empty<byte>());
        source.Attachments[0].Content = null;
        destination.Attachments[0].Content = null;
        source.Attachments[0].ContentSource = sourceContent;
        destination.Attachments[0].ContentSource = destinationContent;
        source.Attachments[0].Length = payload.Length;
        destination.Attachments[0].Length = payload.Length;

        EmailSemanticComparisonReport report = await EmailSemanticComparer.CompareAsync(source, destination);

        Assert.True(report.IsMatch);
        Assert.Equal(1, sourceContent.AsyncOpenCount);
        Assert.Equal(1, destinationContent.AsyncOpenCount);
        Assert.Equal(0, sourceContent.SyncOpenCount);
        Assert.Equal(payload.Length, report.Source.AttachmentBytesHashed);
    }

    [Fact]
    public void Attachment_content_matches_when_only_one_source_declares_its_length() {
        byte[] payload = Encoding.UTF8.GetBytes("same streamed bytes");
        EmailDocument source = CreateDocument(Array.Empty<byte>());
        EmailDocument destination = CreateDocument(Array.Empty<byte>());
        source.Attachments[0].Content = null;
        source.Attachments[0].ContentSource = new LengthOptionalContentSource(payload, null);
        source.Attachments[0].Length = payload.Length;
        destination.Attachments[0].Content = null;
        destination.Attachments[0].ContentSource = new LengthOptionalContentSource(payload, payload.Length);
        destination.Attachments[0].Length = payload.Length;

        EmailSemanticComparisonReport report = EmailSemanticComparer.Compare(source, destination);

        Assert.True(report.IsMatch);
    }

    [Fact]
    public void AttachmentDigestHonorsConfiguredLimit() {
        EmailDocument document = CreateDocument(new byte[17]);
        var options = new EmailSemanticComparisonOptions(maxAttachmentBytes: 16);

        EmailLimitExceededException exception = Assert.Throws<EmailLimitExceededException>(() =>
            EmailSemanticComparer.CreateFingerprint(document, options));

        Assert.Equal(nameof(EmailSemanticComparisonOptions.MaxAttachmentBytes), exception.LimitName);
    }

    private static EmailDocument CreateDocument(byte[] payload) {
        var document = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            OutlookItemKind = OutlookItemKind.Message,
            MessageClass = "IPM.Note",
            Subject = "semantic subject",
            MessageId = "semantic@example.test",
            Date = new DateTimeOffset(2026, 7, 17, 5, 0, 0, TimeSpan.Zero),
            From = new EmailAddress("sender@example.test", "Sender")
        };
        document.Body.Text = "semantic body";
        document.Headers.Add(new EmailHeader("X-Test", "value", "raw value"));
        document.Recipients.Add(new EmailRecipient(EmailRecipientKind.To,
            new EmailAddress("recipient@example.test", "Recipient")));
        document.MapiProperties.Add(new MapiProperty(0x8000, MapiPropertyType.MultipleUnicode,
            new object[] { "one", "two" }, name: new MapiNamedProperty(
                new Guid("00020329-0000-0000-C000-000000000046"), "CustomValues")));
        document.Attachments.Add(new EmailAttachment {
            FileName = "payload.bin",
            ContentType = "application/octet-stream",
            Content = payload,
            Length = payload.LongLength
        });
        return document;
    }

    private sealed class CountingContentSource : IEmailContentSource {
        private readonly byte[] _content;
        internal CountingContentSource(byte[] content) { _content = content; }
        public long? Length => _content.LongLength;
        internal int SyncOpenCount { get; private set; }
        internal int AsyncOpenCount { get; private set; }
        public Stream OpenRead() {
            SyncOpenCount++;
            return new MemoryStream(_content, writable: false);
        }
        public Task<Stream> OpenReadAsync(CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            AsyncOpenCount++;
            return Task.FromResult<Stream>(new MemoryStream(_content, writable: false));
        }
    }

    private sealed class LengthOptionalContentSource : IEmailContentSource {
        private readonly byte[] _content;
        internal LengthOptionalContentSource(byte[] content, long? length) {
            _content = content;
            Length = length;
        }
        public long? Length { get; }
        public Stream OpenRead() => new MemoryStream(_content, writable: false);
        public Task<Stream> OpenReadAsync(CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            return Task.FromResult<Stream>(OpenRead());
        }
    }
}
