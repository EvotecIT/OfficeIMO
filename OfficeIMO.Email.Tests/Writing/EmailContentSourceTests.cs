using OfficeIMO.Email;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class EmailContentSourceTests {
    [Theory]
    [InlineData(EmailFileFormat.Eml)]
    [InlineData(EmailFileFormat.OutlookMsg)]
    [InlineData(EmailFileFormat.Tnef)]
    public void WritersConsumeReopenableAttachmentSources(EmailFileFormat format) {
        byte[] payload = Encoding.UTF8.GetBytes("lazy attachment");
        var source = new CountingContentSource(payload);
        var document = new EmailDocument { Format = format, Subject = "lazy" };
        document.Attachments.Add(new EmailAttachment {
            FileName = "lazy.txt",
            ContentType = "text/plain",
            ContentSource = source,
            Length = payload.Length
        });

        byte[] artifact = new EmailDocumentWriter().ToBytes(document, format);
        EmailAttachment attachment = Assert.Single(new EmailDocumentReader().Read(artifact).Document.Attachments);

        Assert.Equal(payload, attachment.Content);
        Assert.Equal(1, source.OpenCount);
    }

    [Fact]
    public async Task AttachmentCanOpenSourceSynchronouslyAndAsynchronously() {
        byte[] payload = new byte[] { 1, 2, 3 };
        var attachment = new EmailAttachment { ContentSource = new CountingContentSource(payload) };

        using Stream synchronous = attachment.OpenContentStream();
        using Stream asynchronous = await attachment.OpenContentStreamAsync();

        Assert.Equal(payload, ReadAll(synchronous));
        Assert.Equal(payload, ReadAll(asynchronous));
    }

    private static byte[] ReadAll(Stream stream) {
        using var output = new MemoryStream();
        stream.CopyTo(output);
        return output.ToArray();
    }

    private sealed class CountingContentSource : IEmailContentSource {
        private readonly byte[] _content;
        internal CountingContentSource(byte[] content) { _content = content; }
        public long? Length => _content.LongLength;
        internal int OpenCount { get; private set; }
        public Stream OpenRead() { OpenCount++; return new MemoryStream(_content, writable: false); }
        public Task<Stream> OpenReadAsync(CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            OpenCount++;
            return Task.FromResult<Stream>(new MemoryStream(_content, writable: false));
        }
    }
}