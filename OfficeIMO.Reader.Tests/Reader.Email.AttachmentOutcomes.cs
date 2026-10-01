using OfficeIMO.Email;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Email;
using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests;

[Collection("ReaderRegistryNonParallel")]
public sealed class ReaderEmailAttachmentOutcomeTests {
    [Fact]
    public void UnsupportedPayloadIsNotOpenedAndUnavailablePayloadHasItsOwnOutcome() {
        var source = new RefusingContentSource();
        var document = new EmailDocument();
        document.Attachments.Add(new EmailAttachment { FileName = "opaque.unsupported", ContentSource = source });
        document.Attachments.Add(new EmailAttachment { FileName = "unavailable.txt" });
        using var input = new MemoryStream();
        var result = EmailReaderProjection.ProjectEmailDocumentsToStreamResult(new[] { document }, new string?[] { "source.eml" },
            Array.Empty<EmailDiagnostic>(), EmailFileFormat.Eml, "source.eml", input, new ReaderOptions(), CancellationToken.None);
        Assert.Equal(0, source.OpenCount);
        Assert.Contains(result.Diagnostics, value => value.Code == "EMAIL_ATTACHMENT_READER_UNSUPPORTED" && value.Location?.Path == "source.eml!/opaque.unsupported");
        Assert.Contains(result.Diagnostics, value => value.Code == "EMAIL_ATTACHMENT_READER_CONTENT_UNAVAILABLE");
        Assert.Equal("2", result.Metadata.Single(value => value.Name == "AttachmentExtractionSkipped").Value);
        Assert.Equal("0", result.Metadata.Single(value => value.Name == "AttachmentExtractionAttempted").Value);
    }

    [Fact]
    public void HandlerOutcomesReachDocumentCountsChunksAndSelectiveStoreItemDiagnostics() {
        var reader = new OfficeDocumentReaderBuilder().AddEmailHandlers().AddHandler(new ReaderHandlerRegistration {
            Id = "attachment-outcome-test", Kind = ReaderInputKind.Text, Extensions = new[] { ".outcome" },
            ReadStream = (stream, name, options, token) => {
                if (name == "failed.outcome") throw new InvalidDataException("Sensitive decoder details");
                if (name == "empty.outcome") return Array.Empty<ReaderChunk>();
                return new[] { new ReaderChunk { Text = "Extracted body", Location = new ReaderLocation { SourceBlockKind = "text" } } };
            }
        }).Build();
        var document = new EmailDocument { Subject = "Outcome evidence" };
        foreach (string name in new[] { "succeeded", "failed", "empty" })
            document.Attachments.Add(new EmailAttachment { FileName = name + ".outcome", Content = new byte[] { 1 }, ContentType = "application/octet-stream" });
        byte[] bytes = document.ToBytes();
        using var input = new MemoryStream(bytes);
        var result = reader.ReadDocument(input, "source.eml");
        Assert.Equal("3", result.Metadata.Single(value => value.Name == "AttachmentExtractionAttempted").Value);
        foreach (string outcome in new[] { "Succeeded", "Failed", "Empty" })
            Assert.Equal("1", result.Metadata.Single(value => value.Name == "AttachmentExtraction" + outcome).Value);
        Assert.Contains(result.Diagnostics, value => value.Code == "EMAIL_ATTACHMENT_READER_FAILED" && value.Location?.Path == "source.eml!/failed.outcome");
        Assert.Contains(result.Chunks, chunk => chunk.Warnings?.Any(value => value.StartsWith("EMAIL_ATTACHMENT_READER_EMPTY")) == true);
        Assert.DoesNotContain(result.Diagnostics, value => value.Message.Contains("Sensitive decoder details"));
        string root = Path.Combine(Path.GetTempPath(), "officeimo-reader-outcomes-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            File.WriteAllBytes(Path.Combine(root, "source.eml"), bytes);
            using var session = OfficeIMO.Email.Store.EmailStoreSession.Open(root);
            string id = Assert.Single(session.EnumerateItems()).Id;
            var item = reader.ReadEmailStoreItem(root, id);
            Assert.Contains(item.ItemDiagnostics, value => value.Code == "EMAIL_ATTACHMENT_READER_FAILED" && value.Location!.EndsWith("!/failed.outcome"));
            Assert.Contains(item.ItemDiagnostics, value => value.Code == "EMAIL_ATTACHMENT_READER_EMPTY");
            Assert.Contains(item.ItemDiagnostics, value => value.Code == "EMAIL_ATTACHMENT_READER_SUCCEEDED");
        } finally { Directory.Delete(root, recursive: true); }
    }

    [Fact]
    public void AttachmentHandlerCancellationPropagates() {
        var reader = new OfficeDocumentReaderBuilder().AddEmailHandlers().AddHandler(new ReaderHandlerRegistration {
            Id = "attachment-cancellation-test", Kind = ReaderInputKind.Text, Extensions = new[] { ".cancel" },
            ReadStream = (stream, name, options, token) => throw new OperationCanceledException()
        }).Build();
        var document = new EmailDocument();
        document.Attachments.Add(new EmailAttachment { FileName = "attachment.cancel", Content = new byte[] { 1 } });
        using var input = new MemoryStream(document.ToBytes());
        Assert.Throws<OperationCanceledException>(() => reader.ReadDocument(input, "source.eml"));
    }

    private sealed class RefusingContentSource : IEmailContentSource {
        internal int OpenCount { get; private set; }
        public long? Length => 10;
        public Stream OpenRead() { OpenCount++; throw new IOException("Payload must remain deferred"); }
        public Task<Stream> OpenReadAsync(CancellationToken cancellationToken = default) => Task.FromResult(OpenRead());
    }
}
