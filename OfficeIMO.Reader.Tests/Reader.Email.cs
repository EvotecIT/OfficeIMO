using OfficeIMO.Email;
using OfficeIMO.Pdf;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Email;
using OfficeIMO.Reader.Rtf;
using OfficeIMO.Rtf;
using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests;

[Collection("ReaderRegistryNonParallel")]
public sealed class ReaderEmailTests {
    [Fact]
    public void EmailKindAndCapabilities_AreBuiltInWithoutChangingEarlierEnumValues() {
        Assert.Equal(16, (int)ReaderInputKind.OpenDocument);
        Assert.Equal(17, (int)ReaderInputKind.AsciiDoc);
        Assert.Equal(18, (int)ReaderInputKind.Latex);
        Assert.Equal(19, (int)ReaderInputKind.Email);
        Assert.Equal(ReaderInputKind.Email, OfficeIMO.Reader.Tests.ReaderTestReaders.All.DetectKind("message.eml"));
        Assert.Equal(ReaderInputKind.Email, OfficeIMO.Reader.Tests.ReaderTestReaders.All.DetectKind("outlook.msg"));
        Assert.Equal(ReaderInputKind.Email, OfficeIMO.Reader.Tests.ReaderTestReaders.All.DetectKind("archive.mbox"));
        Assert.Equal(ReaderInputKind.Email, OfficeIMO.Reader.Tests.ReaderTestReaders.All.DetectKind("winmail.dat"));

        ReaderHandlerCapability capability = Assert.Single(
            OfficeIMO.Reader.Tests.ReaderTestReaders.All.GetCapabilities(), item => item.Id == "officeimo.reader.email");
        Assert.Equal(ReaderInputKind.Email, capability.Kind);
        Assert.Contains(".tnef", capability.Extensions);
        Assert.True(capability.SupportsPath);
        Assert.True(capability.SupportsStream);
    }

    [Fact]
    public void EmlRead_MapsEnvelopeBodyAssetsAndReusableAttachmentContent() {
        byte[] bytes = BuildEmlWithAttachment();

        ReaderChunk[] chunks = OfficeIMO.Reader.Tests.ReaderTestReaders.All.Read(bytes, "sample.eml").ToArray();

        Assert.Contains(chunks, chunk => chunk.Kind == ReaderInputKind.Email &&
            chunk.Location.SourceBlockKind == "email-message" && chunk.Text.Contains("Reader subject", StringComparison.Ordinal));
        Assert.Contains(chunks, chunk => chunk.Kind == ReaderInputKind.Email &&
            chunk.Location.SourceBlockKind == "email-body" && chunk.Text.Contains("Body for retrieval", StringComparison.Ordinal));
        Assert.Contains(chunks, chunk => chunk.Kind == ReaderInputKind.Email &&
            chunk.Location.SourceBlockKind == "email-attachment" && chunk.Text.Contains("notes.txt", StringComparison.Ordinal));
        Assert.Contains(chunks, chunk => chunk.Kind == ReaderInputKind.Text &&
            chunk.Location.Path != null && chunk.Location.Path.EndsWith("!/notes.txt", StringComparison.Ordinal) &&
            chunk.Text.Contains("attachment text", StringComparison.Ordinal));
        Assert.All(chunks, chunk => Assert.False(string.IsNullOrWhiteSpace(chunk.SourceId)));
    }

    [Fact]
    public void ByValueEmlAttachment_ProducesSearchableNestedEmailChunks() {
        var forwarded = new EmailDocument { Subject = "Forwarded subject" };
        forwarded.Body.Text = "Forwarded searchable body";
        byte[] forwardedBytes = new EmailDocumentWriter().ToBytes(forwarded, EmailFileFormat.Eml);
        var parent = new EmailDocument { Subject = "Parent" };
        parent.Attachments.Add(new EmailAttachment {
            FileName = "forwarded.eml",
            ContentType = "application/octet-stream",
            Content = forwardedBytes,
            Length = forwardedBytes.Length
        });
        byte[] bytes = new EmailDocumentWriter().ToBytes(parent, EmailFileFormat.Eml);

        ReaderChunk[] chunks = OfficeIMO.Reader.Tests.ReaderTestReaders.All.Read(bytes, "parent.eml").ToArray();

        Assert.Contains(chunks, chunk => chunk.Location.Path != null &&
            chunk.Location.Path.EndsWith("!/forwarded.eml", StringComparison.Ordinal) &&
            chunk.Text.Contains("Forwarded subject", StringComparison.Ordinal));
        Assert.Contains(chunks, chunk => chunk.Location.Path != null &&
            chunk.Location.Path.EndsWith("!/forwarded.eml", StringComparison.Ordinal) &&
            chunk.Text.Contains("Forwarded searchable body", StringComparison.Ordinal));
    }

    [Fact]
    public void EmlRichResult_ContainsTypedMetadataMaterializableAssetsAndHtml() {
        byte[] bytes = BuildEmlWithAttachment();

        OfficeDocumentReadResult result = OfficeIMO.Reader.Tests.ReaderTestReaders.All.ReadDocument(bytes, "sample.eml");

        Assert.Equal(ReaderInputKind.Email, result.Kind);
        Assert.Equal("Reader subject", result.Source.Title);
        Assert.Equal("Sender <sender@example.test>", result.Source.Author);
        Assert.Equal("<p>Body for retrieval</p>", result.Html);
        Assert.Contains("officeimo.email.eml", result.CapabilitiesUsed);
        Assert.Contains(result.Metadata, item => item.Name == "MessageCount" && item.Value == "1");
        Assert.Contains(result.Metadata, item => item.Name == "To" && item.Value!.Contains("recipient@example.test", StringComparison.Ordinal));
        OfficeDocumentAsset asset = Assert.Single(result.Assets);
        Assert.Equal("notes.txt", asset.FileName);
        Assert.Equal("text/plain", asset.MediaType);
        Assert.Equal("attachment text", Encoding.UTF8.GetString(Assert.IsType<byte[]>(asset.PayloadBytes)));
        Assert.True(asset.PayloadHashMatches(out string? actualHash));
        Assert.Equal(asset.PayloadHash, actualHash);
    }

    [Fact]
    public void NonSeekableUnnamedStream_IsDetectedAndMappedInOneRichResult() {
        using var stream = new NonSeekableReadStream(BuildEmlWithAttachment());

        OfficeDocumentReadResult result = OfficeIMO.Reader.Tests.ReaderTestReaders.All.ReadDocument(stream);

        Assert.Equal(ReaderInputKind.Email, result.Kind);
        Assert.Equal("Reader subject", result.Source.Title);
        Assert.Single(result.Assets);
        Assert.Contains(result.Chunks, chunk => chunk.Location.SourceBlockKind == "email-body");
    }

    [Fact]
    public async Task ContentDetection_RecognizesRenamedMsgButDoesNotClaimArbitraryCompoundSignature() {
        var appointment = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            OutlookItemKind = OutlookItemKind.Appointment,
            MessageClass = "IPM.Appointment",
            Subject = "Planning"
        };
        appointment.Attachments.Add(new EmailAttachment {
            FileName = "large.bin",
            Content = new byte[128 * 1024],
            Length = 128 * 1024
        });
        appointment.Appointment = new OutlookAppointment {
            Start = new DateTimeOffset(2026, 7, 10, 8, 0, 0, TimeSpan.Zero),
            End = new DateTimeOffset(2026, 7, 10, 9, 0, 0, TimeSpan.Zero),
            Location = "Room 1"
        };
        byte[] msg = new EmailDocumentWriter().ToBytes(appointment, EmailFileFormat.OutlookMsg);

        ReaderDetectionResult syncDetection = OfficeIMO.Reader.Tests.ReaderTestReaders.All.Detect(msg, "renamed.bin");
        ReaderDetectionResult asyncDetection = await OfficeIMO.Reader.Tests.ReaderTestReaders.All.DetectAsync(msg, "renamed.bin");
        Assert.Equal(ReaderInputKind.Email, syncDetection.Kind);
        Assert.Equal(ReaderInputKind.Email, asyncDetection.Kind);
        Assert.True(syncDetection.ContainerInspected);
        Assert.Contains("container:msg-properties-stream", syncDetection.Evidence);

        ReaderDetectionResult boundedDetection = OfficeIMO.Reader.Tests.ReaderTestReaders.All.Detect(msg, "renamed.bin",
            new ReaderDetectionOptions { MaxProbeBytes = 256 });
        ReaderDetectionResult boundedAsyncDetection = await OfficeIMO.Reader.Tests.ReaderTestReaders.All.DetectAsync(msg, "renamed.bin",
            new ReaderDetectionOptions { MaxProbeBytes = 256 });
        Assert.Equal(ReaderInputKind.Email, boundedDetection.Kind);
        Assert.Equal(ReaderInputKind.Email, boundedAsyncDetection.Kind);
        Assert.Equal(256, boundedDetection.InspectedBytes);

        OfficeDocumentReadResult detected = OfficeIMO.Reader.Tests.ReaderTestReaders.All.ReadDocument(msg, "renamed.bin");
        Assert.Equal(ReaderInputKind.Email, detected.Kind);
        Assert.Contains(detected.Metadata, item => item.Category == "email.appointment" &&
            item.Name == "Location" && item.Value == "Room 1");

        byte[] signatureOnly = { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };
        Assert.Equal(ReaderInputKind.Unknown, OfficeIMO.Reader.Tests.ReaderTestReaders.All.Detect(signatureOnly, "legacy.bin").Kind);
        Assert.DoesNotContain(OfficeIMO.Reader.Tests.ReaderTestReaders.All.Read(signatureOnly, "legacy.bin"), chunk => chunk.Kind == ReaderInputKind.Email);
    }

    [Fact]
    public void MboxRead_EmitsEveryMessageAndEnvelopeMetadata() {
        var mailbox = new EmailMailbox();
        mailbox.Messages.Add(new EmailMailboxEntry(new EmailDocument {
            Format = EmailFileFormat.Eml,
            Subject = "First",
            From = new EmailAddress("first@example.test"),
            Date = new DateTimeOffset(2026, 7, 10, 10, 0, 0, TimeSpan.Zero)
        }) { EnvelopeSender = "first@example.test" });
        mailbox.Messages.Add(new EmailMailboxEntry(new EmailDocument {
            Format = EmailFileFormat.Eml,
            Subject = "Second",
            From = new EmailAddress("second@example.test"),
            Date = new DateTimeOffset(2026, 7, 10, 11, 0, 0, TimeSpan.Zero)
        }) { EnvelopeSender = "second@example.test" });
        byte[] bytes = new EmailMailboxWriter().ToBytes(mailbox);

        OfficeDocumentReadResult result = OfficeIMO.Reader.Tests.ReaderTestReaders.All.ReadDocument(bytes, "archive.mbox");

        Assert.Equal(ReaderInputKind.Email, result.Kind);
        Assert.Contains(result.Metadata, item => item.Name == "MessageCount" && item.Value == "2");
        Assert.Contains(result.Metadata, item => item.Name == "EnvelopeSender" && item.Value == "first@example.test");
        Assert.Contains(result.Chunks, chunk => chunk.Text.Contains("# First", StringComparison.Ordinal));
        Assert.Contains(result.Chunks, chunk => chunk.Text.Contains("# Second", StringComparison.Ordinal));
        Assert.Contains(result.Chunks, chunk => chunk.Location.Path!.Contains("message-000002.eml", StringComparison.Ordinal));
    }

    [Fact]
    public void InvalidNamedEmail_ProducesStructuredDiagnosticsAndInputBoundsRemainEffective() {
        OfficeDocumentReadResult invalid = OfficeIMO.Reader.Tests.ReaderTestReaders.All.ReadDocument(Encoding.ASCII.GetBytes("not an email"), "broken.eml");
        Assert.Equal(ReaderInputKind.Email, invalid.Kind);
        Assert.Contains(invalid.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_FORMAT_UNKNOWN" &&
            diagnostic.Severity == OfficeDocumentDiagnosticSeverity.Error);

        byte[] bytes = BuildEmlWithAttachment();
        Assert.Throws<IOException>(() => OfficeIMO.Reader.Tests.ReaderTestReaders.All.Read(bytes, "bounded.eml", new ReaderOptions {
            MaxInputBytes = 32
        }).ToArray());
    }

    [Fact]
    public void DefaultEmailSnapshotLimitRejectsOversizedSeekableStreamsBeforeReading() {
        using var stream = new LengthOnlySeekableStream(EmailReaderOptions.Default.MaxInputBytes + 1);

        IOException exception = Assert.Throws<IOException>(() =>
            OfficeIMO.Reader.Tests.ReaderTestReaders.All.ReadDocument(stream, "oversized.eml"));

        Assert.Contains("MaxInputBytes", exception.Message, StringComparison.Ordinal);
        Assert.Equal(0, stream.ReadCount);
    }

    [Fact]
    public void OversizedUnknownSeekableStreamsFallBackWithoutEmailProbeFailure() {
        using var chunkStream = new LengthOnlySeekableStream(EmailReaderOptions.Default.MaxInputBytes + 1);
        using var documentStream = new LengthOnlySeekableStream(EmailReaderOptions.Default.MaxInputBytes + 1);

        ReaderChunk[] chunks = OfficeIMO.Reader.Tests.ReaderTestReaders.All.Read(chunkStream, "oversized.bin").ToArray();
        OfficeDocumentReadResult document = OfficeIMO.Reader.Tests.ReaderTestReaders.All.ReadDocument(documentStream, "oversized.bin");

        Assert.Single(chunks);
        Assert.Equal(ReaderInputKind.Unknown, chunks[0].Kind);
        Assert.Equal(ReaderInputKind.Unknown, document.Kind);
    }

    [Fact]
    public void DefaultFolderIngestionIncludesWinmailDatWithoutClaimingOtherDatFiles() {
        string folder = Path.Combine(Path.GetTempPath(), "officeimo-reader-winmail-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        try {
            var document = new EmailDocument { Format = EmailFileFormat.Tnef, Subject = "Folder TNEF" };
            document.Body.Text = "winmail body";
            File.WriteAllBytes(Path.Combine(folder, "winmail.dat"),
                new EmailDocumentWriter().ToBytes(document, EmailFileFormat.Tnef));
            File.WriteAllText(Path.Combine(folder, "other.dat"), "not an email");

            ReaderChunk[] chunks = OfficeIMO.Reader.Tests.ReaderTestReaders.All.ReadFolder(folder,
                new ReaderFolderOptions { Recurse = false, MaxFiles = 10 }, new ReaderOptions()).ToArray();

            Assert.Contains(chunks, chunk => chunk.Kind == ReaderInputKind.Email &&
                chunk.Text.Contains("Folder TNEF", StringComparison.Ordinal));
            Assert.DoesNotContain(chunks, chunk => chunk.Location.Path != null &&
                chunk.Location.Path.EndsWith("other.dat", StringComparison.OrdinalIgnoreCase));
        } finally {
            Directory.Delete(folder, true);
        }
    }

    [Fact]
    public void RtfOnlyMsg_UsesConfiguredSemanticRtfHandler() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddEmailHandler()
            .AddRtfHandler()
            .Build();
        RtfDocument rtf = RtfDocument.Create();
        rtf.AddParagraph("Semantic RTF email body");
        var document = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            Subject = "RTF body"
        };
        document.Body.Rtf = rtf.ToRtf();
        byte[] bytes = new EmailDocumentWriter().ToBytes(document, EmailFileFormat.OutlookMsg);

        ReaderChunk[] chunks = reader.Read(bytes, "rtf-body.msg").ToArray();

        Assert.Contains(chunks, chunk => chunk.Kind == ReaderInputKind.Rtf &&
            chunk.Location.SourceBlockKind == "email-body-rtf" &&
            chunk.Text.Contains("Semantic RTF email body", StringComparison.Ordinal));
        Assert.DoesNotContain(chunks.SelectMany(chunk => chunk.Warnings ?? Array.Empty<string>()),
            warning => warning.StartsWith("EMAIL_RTF_BODY_PRESERVED", StringComparison.Ordinal));
    }

    [Fact]
    public void AttachmentNamesRemainLogicalWhenInvalidAsWindowsPaths() {
        Assert.Equal(ReaderInputKind.Text, OfficeIMO.Reader.Tests.ReaderTestReaders.All.DetectKind("report|draft.txt"));
        var document = new EmailDocument { Subject = "Logical attachment name" };
        document.Attachments.Add(new EmailAttachment {
            FileName = "report|draft.txt",
            ContentType = "text/plain",
            Content = Encoding.UTF8.GetBytes("content"),
            Length = 7
        });
        byte[] bytes = new EmailDocumentWriter().ToBytes(document);

        OfficeDocumentReadResult result = OfficeIMO.Reader.Tests.ReaderTestReaders.All.ReadDocument(bytes, "sample.eml");

        OfficeDocumentAsset asset = Assert.Single(result.Assets);
        Assert.Equal("report|draft.txt", asset.FileName);
        Assert.Equal("content", Encoding.UTF8.GetString(Assert.IsType<byte[]>(asset.PayloadBytes)));
    }

    [Fact]
    public void ExtensionlessPdfAttachmentProducesSearchableChildChunks() {
        string pdfPath = Path.Combine(Path.GetTempPath(), string.Concat(Guid.NewGuid().ToString("N"), ".pdf"));
        try {
            PdfDocument pdf = PdfDocument.Create();
            pdf.Paragraph(paragraph => paragraph.Text("Searchable invoice content"));
            pdf.Save(pdfPath);
            byte[] payload = File.ReadAllBytes(pdfPath);
            var document = new EmailDocument { Subject = "invoice" };
            document.Attachments.Add(new EmailAttachment {
                FileName = "invoice",
                ContentType = "application/pdf",
                Content = payload,
                Length = payload.Length
            });

            ReaderChunk[] chunks = OfficeIMO.Reader.Tests.ReaderTestReaders.All.Read(
                new EmailDocumentWriter().ToBytes(document), "invoice.eml").ToArray();

            Assert.Contains(chunks, chunk => chunk.Kind == ReaderInputKind.Pdf &&
                chunk.Location.Path != null && chunk.Location.Path.EndsWith("!/invoice", StringComparison.Ordinal) &&
                chunk.Text.Contains("Searchable invoice content", StringComparison.Ordinal));
        } finally {
            if (File.Exists(pdfPath)) File.Delete(pdfPath);
        }
    }

    [Fact]
    public void DeferredAttachmentContentProducesSearchableChildChunksWithoutResidentPayload() {
        var source = new CountingContentSource(Encoding.UTF8.GetBytes("deferred attachment text"));
        var document = new EmailDocument { Subject = "Deferred attachment" };
        document.Attachments.Add(new EmailAttachment {
            FileName = "deferred.txt",
            ContentType = "text/plain",
            ContentSource = source,
            Length = source.Length!.Value
        });
        using var input = new MemoryStream(Array.Empty<byte>());

        OfficeDocumentReadResult result = OfficeIMO.Reader.Email.EmailReaderProjection.ProjectEmailDocumentsToStreamResult(
            new[] { document },
            new string?[] { "mailbox.pst!/Inbox/item-000001" },
            Array.Empty<EmailDiagnostic>(),
            EmailFileFormat.OutlookMsg,
            "mailbox.pst",
            input,
            new ReaderOptions(),
            CancellationToken.None);

        Assert.Equal(1, source.OpenCount);
        Assert.Contains(result.Chunks, chunk =>
            chunk.Location.Path == "mailbox.pst!/Inbox/item-000001!/deferred.txt" &&
            chunk.Text.Contains("deferred attachment text", StringComparison.Ordinal));
        Assert.Null(Assert.Single(result.Assets).PayloadBytes);
    }

    private static byte[] BuildEmlWithAttachment() {
        var document = new EmailDocument {
            Format = EmailFileFormat.Eml,
            Subject = "Reader subject",
            From = new EmailAddress("sender@example.test", "Sender"),
            Date = new DateTimeOffset(2026, 7, 10, 9, 30, 0, TimeSpan.Zero)
        };
        document.Recipients.Add(new EmailRecipient(EmailRecipientKind.To,
            new EmailAddress("recipient@example.test", "Recipient")));
        document.Body.Text = "Body for retrieval";
        document.Body.Html = "<p>Body for retrieval</p>";
        document.Attachments.Add(new EmailAttachment {
            FileName = "notes.txt",
            ContentType = "text/plain",
            Length = Encoding.UTF8.GetByteCount("attachment text"),
            Content = Encoding.UTF8.GetBytes("attachment text")
        });
        return new EmailDocumentWriter().ToBytes(document);
    }

    private sealed class LengthOnlySeekableStream : Stream {
        private long _position;

        internal LengthOnlySeekableStream(long length) {
            Length = length;
        }

        internal int ReadCount { get; private set; }
        public override bool CanRead => true;
        public override bool CanSeek => true;
        public override bool CanWrite => false;
        public override long Length { get; }
        public override long Position { get => _position; set => _position = value; }
        public override void Flush() { }
        public override int Read(byte[] buffer, int offset, int count) {
            ReadCount++;
            return 0;
        }
        public override long Seek(long offset, SeekOrigin origin) {
            _position = origin == SeekOrigin.Begin ? offset :
                origin == SeekOrigin.Current ? _position + offset : Length + offset;
            return _position;
        }
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }

    private sealed class CountingContentSource : IEmailContentSource {
        private readonly byte[] _content;

        internal CountingContentSource(byte[] content) {
            _content = content;
        }

        internal int OpenCount { get; private set; }
        public long? Length => _content.Length;

        public Stream OpenRead() {
            OpenCount++;
            return new MemoryStream(_content, writable: false);
        }

        public Task<Stream> OpenReadAsync(CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            return Task.FromResult(OpenRead());
        }
    }
}
