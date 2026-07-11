using OfficeIMO.Email;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Rtf;
using OfficeIMO.Rtf;
using System;
using System.IO;
using System.Linq;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests;

[Collection("ReaderRegistryNonParallel")]
public sealed class ReaderEmailTests {
    [Fact]
    public void EmailKindAndCapabilities_AreBuiltInWithoutChangingEarlierEnumValues() {
        Assert.Equal(17, (int)ReaderInputKind.Latex);
        Assert.Equal(18, (int)ReaderInputKind.Email);
        Assert.Equal(19, (int)ReaderInputKind.OpenDocument);
        Assert.Equal(ReaderInputKind.Email, DocumentReader.DetectKind("message.eml"));
        Assert.Equal(ReaderInputKind.Email, DocumentReader.DetectKind("outlook.msg"));
        Assert.Equal(ReaderInputKind.Email, DocumentReader.DetectKind("archive.mbox"));
        Assert.Equal(ReaderInputKind.Email, DocumentReader.DetectKind("winmail.dat"));

        ReaderHandlerCapability capability = Assert.Single(
            DocumentReader.GetCapabilities(), item => item.Id == "officeimo.reader.email");
        Assert.Equal(ReaderInputKind.Email, capability.Kind);
        Assert.Contains(".tnef", capability.Extensions);
        Assert.True(capability.SupportsPath);
        Assert.True(capability.SupportsStream);
    }

    [Fact]
    public void EmlRead_MapsEnvelopeBodyAssetsAndReusableAttachmentContent() {
        byte[] bytes = BuildEmlWithAttachment();

        ReaderChunk[] chunks = DocumentReader.Read(bytes, "sample.eml").ToArray();

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
    public void EmlRichResult_ContainsTypedMetadataMaterializableAssetsAndHtml() {
        byte[] bytes = BuildEmlWithAttachment();

        OfficeDocumentReadResult result = DocumentReader.ReadDocument(bytes, "sample.eml");

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

        OfficeDocumentReadResult result = OfficeDocumentReader.Default.ReadDocument(stream);

        Assert.Equal(ReaderInputKind.Email, result.Kind);
        Assert.Equal("Reader subject", result.Source.Title);
        Assert.Single(result.Assets);
        Assert.Contains(result.Chunks, chunk => chunk.Location.SourceBlockKind == "email-body");
    }

    [Fact]
    public void ContentDetection_RecognizesRenamedMsgButDoesNotClaimArbitraryCompoundSignature() {
        var appointment = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            OutlookItemKind = OutlookItemKind.Appointment,
            MessageClass = "IPM.Appointment",
            Subject = "Planning"
        };
        appointment.Appointment = new OutlookAppointment {
            Start = new DateTimeOffset(2026, 7, 10, 8, 0, 0, TimeSpan.Zero),
            End = new DateTimeOffset(2026, 7, 10, 9, 0, 0, TimeSpan.Zero),
            Location = "Room 1"
        };
        byte[] msg = new EmailDocumentWriter().WriteToBytes(appointment, EmailFileFormat.OutlookMsg);

        OfficeDocumentReadResult detected = DocumentReader.ReadDocument(msg, "renamed.bin");
        Assert.Equal(ReaderInputKind.Email, detected.Kind);
        Assert.Contains(detected.Metadata, item => item.Category == "email.appointment" &&
            item.Name == "Location" && item.Value == "Room 1");

        byte[] signatureOnly = { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };
        Assert.DoesNotContain(DocumentReader.Read(signatureOnly, "legacy.bin"), chunk => chunk.Kind == ReaderInputKind.Email);
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
        byte[] bytes = new EmailMailboxWriter().WriteToBytes(mailbox);

        OfficeDocumentReadResult result = DocumentReader.ReadDocument(bytes, "archive.mbox");

        Assert.Equal(ReaderInputKind.Email, result.Kind);
        Assert.Contains(result.Metadata, item => item.Name == "MessageCount" && item.Value == "2");
        Assert.Contains(result.Metadata, item => item.Name == "EnvelopeSender" && item.Value == "first@example.test");
        Assert.Contains(result.Chunks, chunk => chunk.Text.Contains("# First", StringComparison.Ordinal));
        Assert.Contains(result.Chunks, chunk => chunk.Text.Contains("# Second", StringComparison.Ordinal));
        Assert.Contains(result.Chunks, chunk => chunk.Location.Path!.Contains("message-000002.eml", StringComparison.Ordinal));
    }

    [Fact]
    public void InvalidNamedEmail_ProducesStructuredDiagnosticsAndInputBoundsRemainEffective() {
        OfficeDocumentReadResult invalid = DocumentReader.ReadDocument(Encoding.ASCII.GetBytes("not an email"), "broken.eml");
        Assert.Equal(ReaderInputKind.Email, invalid.Kind);
        Assert.Contains(invalid.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_FORMAT_UNKNOWN" &&
            diagnostic.Severity == OfficeDocumentDiagnosticSeverity.Error);

        byte[] bytes = BuildEmlWithAttachment();
        Assert.Throws<IOException>(() => DocumentReader.Read(bytes, "bounded.eml", new ReaderOptions {
            MaxInputBytes = 32
        }).ToArray());
    }

    [Fact]
    public void DefaultFolderIngestionIncludesWinmailDatWithoutClaimingOtherDatFiles() {
        string folder = Path.Combine(Path.GetTempPath(), "officeimo-reader-winmail-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        try {
            var document = new EmailDocument { Format = EmailFileFormat.Tnef, Subject = "Folder TNEF" };
            document.Body.Text = "winmail body";
            File.WriteAllBytes(Path.Combine(folder, "winmail.dat"),
                new EmailDocumentWriter().WriteToBytes(document, EmailFileFormat.Tnef));
            File.WriteAllText(Path.Combine(folder, "other.dat"), "not an email");

            ReaderChunk[] chunks = DocumentReader.ReadFolder(folder,
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
    public void RtfOnlyMsg_UsesRegisteredSemanticRtfHandler() {
        try {
            DocumentReaderRtfRegistrationExtensions.RegisterRtfHandler();
            RtfDocument rtf = RtfDocument.Create();
            rtf.AddParagraph("Semantic RTF email body");
            var document = new EmailDocument {
                Format = EmailFileFormat.OutlookMsg,
                Subject = "RTF body"
            };
            document.Body.Rtf = rtf.ToRtf();
            byte[] bytes = new EmailDocumentWriter().WriteToBytes(document, EmailFileFormat.OutlookMsg);

            ReaderChunk[] chunks = DocumentReader.Read(bytes, "rtf-body.msg").ToArray();

            Assert.Contains(chunks, chunk => chunk.Kind == ReaderInputKind.Rtf &&
                chunk.Location.SourceBlockKind == "email-body-rtf" &&
                chunk.Text.Contains("Semantic RTF email body", StringComparison.Ordinal));
            Assert.DoesNotContain(chunks.SelectMany(chunk => chunk.Warnings ?? Array.Empty<string>()),
                warning => warning.StartsWith("EMAIL_RTF_BODY_PRESERVED", StringComparison.Ordinal));
        } finally {
            DocumentReaderRtfRegistrationExtensions.UnregisterRtfHandler();
        }
    }

    [Fact]
    public void AttachmentNamesRemainLogicalWhenInvalidAsWindowsPaths() {
        var document = new EmailDocument { Subject = "Logical attachment name" };
        document.Attachments.Add(new EmailAttachment {
            FileName = "report|draft.txt",
            ContentType = "text/plain",
            Content = Encoding.UTF8.GetBytes("content"),
            Length = 7
        });
        byte[] bytes = new EmailDocumentWriter().WriteToBytes(document);

        OfficeDocumentReadResult result = DocumentReader.ReadDocument(bytes, "sample.eml");

        OfficeDocumentAsset asset = Assert.Single(result.Assets);
        Assert.Equal("report|draft.txt", asset.FileName);
        Assert.Equal("content", Encoding.UTF8.GetString(Assert.IsType<byte[]>(asset.PayloadBytes)));
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
        return new EmailDocumentWriter().WriteToBytes(document);
    }
}
