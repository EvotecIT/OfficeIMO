using OfficeIMO.Email;
using OfficeIMO.Drawing.Internal;
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class MsgMalformedInputTests {
    [Fact]
    public void CorruptNamedPropertyGuidIsDiagnosedWithoutDroppingStandardMessageData() {
        var source = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            Subject = "survives"
        };
        source.MapiProperties.Add(new MapiProperty(0x8000, MapiPropertyType.Unicode, "custom",
            name: new MapiNamedProperty(MsgProjection.PsetidCommon, 0x85FF)));
        byte[] valid = new EmailDocumentWriter().ToBytes(source, EmailFileFormat.OutlookMsg);
        Assert.True(OfficeCompoundFileReader.TryRead(valid, out OfficeCompoundFile? compound, out string? error), error);
        byte[] entries = compound!.Streams["__nameid_version1.0/__substg1.0_00030102"];
        entries[4] = 0;
        entries[5] = 0;
        byte[] corrupt = OfficeCompoundFileWriter.Write(compound.Streams.Select(stream =>
            new OfficeCompoundStream(stream.Key, stream.Value)).ToArray());

        EmailReadResult result = new EmailDocumentReader().Read(corrupt);

        Assert.Equal("survives", result.Document.Subject);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_MSG_NAMEID_GUID_INVALID" &&
            diagnostic.Severity == EmailDiagnosticSeverity.Warning);
    }

    [Fact]
    public void OutOfRangeNamedPropertyIndexIsDiagnosedWithoutDroppingStandardMessageData() {
        var source = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            Subject = "standard data survives"
        };
        source.MapiProperties.Add(new MapiProperty(0x8000, MapiPropertyType.Unicode, "custom",
            name: new MapiNamedProperty(MsgProjection.PsetidCommon, 0x85FF)));
        byte[] valid = new EmailDocumentWriter().ToBytes(source, EmailFileFormat.OutlookMsg);
        Assert.True(OfficeCompoundFileReader.TryRead(valid, out OfficeCompoundFile? compound, out string? error), error);
        byte[] entries = compound!.Streams["__nameid_version1.0/__substg1.0_00030102"];
        entries[6] = 0xFF;
        entries[7] = 0xFF;
        byte[] corrupt = OfficeCompoundFileWriter.Write(compound.Streams.Select(stream =>
            new OfficeCompoundStream(stream.Key, stream.Value)).ToArray());

        EmailReadResult result = new EmailDocumentReader().Read(corrupt);

        Assert.Equal("standard data survives", result.Document.Subject);
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "EMAIL_MSG_NAMEID_PROPERTY_INDEX_INVALID" &&
            diagnostic.Severity == EmailDiagnosticSeverity.Warning);
    }

    [Fact]
    public void TruncatedCompoundFileReturnsStructuredErrors() {
        var source = new EmailDocument { Format = EmailFileFormat.OutlookMsg, Subject = "truncated" };
        byte[] valid = new EmailDocumentWriter().ToBytes(source, EmailFileFormat.OutlookMsg);

        EmailReadResult result = new EmailDocumentReader().Read(valid.Take(100).ToArray());

        Assert.Equal(EmailFileFormat.Unknown, result.Document.Format);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_MSG_COMPOUND_INVALID");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_FORMAT_UNKNOWN");
    }

    [Fact]
    public void ProtectedPayloadMetadataSurvivesWhenAttachmentBytesAreSkipped() {
        var source = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            MessageClass = "IPM.Note.SMIME",
            Subject = "protected"
        };
        source.Attachments.Add(new EmailAttachment {
            FileName = "smime.p7m",
            ContentType = "application/pkcs7-mime",
            Content = new byte[] { 1, 2, 3 },
            Length = 3
        });
        byte[] bytes = new EmailDocumentWriter().ToBytes(source, EmailFileFormat.OutlookMsg);

        EmailDocument result = new EmailDocumentReader(new EmailReaderOptions(includeAttachmentContent: false))
            .Read(bytes).Document;

        Assert.Equal(EmailProtectionKind.SmimeOpaque, result.Protection.Kind);
        Assert.Equal("smime.p7m", result.Protection.PayloadAttachment!.FileName);
        Assert.Null(result.Protection.PayloadAttachment.Content);
    }
}
