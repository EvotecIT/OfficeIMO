using OfficeIMO.Email;
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class MsgProtectedContentTests {
    [Theory]
    [InlineData("IPM.Note.SMIME", EmailProtectionKind.SmimeOpaque)]
    [InlineData("IPM.Note.SMIME.MultipartSigned", EmailProtectionKind.SmimeClearSigned)]
    public void ExposesProtectedPayloadForExternalCryptographicProcessing(string messageClass,
        EmailProtectionKind expectedKind) {
        byte[] cms = new byte[] { 0x30, 0x03, 0x02, 0x01, 0x01 };
        var source = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            MessageClass = messageClass,
            Subject = "Protected"
        };
        source.Attachments.Add(new EmailAttachment {
            FileName = "smime.p7m",
            ContentType = "application/pkcs7-mime",
            Content = cms,
            Length = cms.Length
        });

        byte[] bytes = new EmailDocumentWriter().WriteToBytes(source, EmailFileFormat.OutlookMsg);
        EmailReadResult result = new EmailDocumentReader().Read(bytes);

        Assert.Equal(expectedKind, result.Document.Protection.Kind);
        Assert.True(result.Document.Protection.IsProtected);
        Assert.Equal(messageClass, result.Document.Protection.MessageClass);
        Assert.Equal(cms, result.Document.Protection.PayloadAttachment!.Content);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_PROTECTED_PAYLOAD_MISSING");
    }

    [Fact]
    public void DiagnosesProtectedMessageWithoutPayloadWithoutAttemptingCryptography() {
        var source = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            MessageClass = "IPM.Note.SMIME",
            Subject = "Missing payload"
        };

        EmailReadResult result = new EmailDocumentReader().Read(
            new EmailDocumentWriter().WriteToBytes(source, EmailFileFormat.OutlookMsg));

        Assert.Equal(EmailProtectionKind.SmimeOpaque, result.Document.Protection.Kind);
        Assert.Null(result.Document.Protection.PayloadAttachment);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_PROTECTED_PAYLOAD_MISSING" &&
            diagnostic.Severity == EmailDiagnosticSeverity.Warning);
    }

    [Fact]
    public void ProjectsWinmailDatInsideMsgWithoutDiscardingOriginalAttachment() {
        var tnef = new EmailDocument {
            Format = EmailFileFormat.Tnef,
            Subject = "Encapsulated"
        };
        tnef.Attachments.Add(new EmailAttachment {
            FileName = "inside.txt",
            ContentType = "text/plain",
            Content = Encoding.UTF8.GetBytes("inside"),
            Length = 6
        });
        byte[] winmail = new EmailDocumentWriter().WriteToBytes(tnef, EmailFileFormat.Tnef);
        var source = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            Subject = "Outer"
        };
        source.Attachments.Add(new EmailAttachment {
            FileName = "winmail.dat",
            ContentType = "application/ms-tnef",
            MapiAttachMethod = 1,
            Content = winmail,
            Length = winmail.Length
        });

        EmailDocument result = new EmailDocumentReader().Read(
            new EmailDocumentWriter().WriteToBytes(source, EmailFileFormat.OutlookMsg)).Document;

        EmailAttachment original = Assert.Single(result.Attachments);
        Assert.Equal(winmail, original.Content);
        Assert.Equal(EmailFileFormat.Tnef, original.EmbeddedDocument!.Format);
        Assert.Equal("Encapsulated", original.EmbeddedDocument.Subject);
        Assert.Equal("inside.txt", Assert.Single(original.EmbeddedDocument.Attachments).FileName);
    }
}
