using MimeKit.Tnef;
using OfficeIMO.Email;
using System.Threading;
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class EmailTnefTests {
    [Fact]
    public void RoundTripsMessageMapiRecipientsAndAttachmentKinds() {
        DateTimeOffset start = new DateTimeOffset(2026, 10, 3, 9, 0, 0, TimeSpan.Zero);
        Guid classId = new Guid("6F9619FF-8B86-D011-B42D-00C04FC964FF");
        EmailDocument child = new EmailDocument { Format = EmailFileFormat.Tnef, Subject = "child" };
        child.Body.Text = "nested";
        EmailDocument source = new EmailDocument {
            Format = EmailFileFormat.Tnef,
            OutlookItemKind = OutlookItemKind.Appointment,
            Subject = "TNEF subject",
            MessageId = "tnef@example.com",
            Date = start,
            Appointment = new OutlookAppointment { Start = start, End = start.AddHours(1), Location = "Room" }
        };
        source.Body.Text = "TNEF body";
        source.Recipients.Add(new EmailRecipient(EmailRecipientKind.To, new EmailAddress("to@example.com", "To")));
        source.MapiProperties.Add(new MapiProperty(0x66AA, MapiPropertyType.MultipleUnicode, new object[] { "one", "two" }));
        source.MapiProperties.Add(new MapiProperty(0x66AB, MapiPropertyType.Guid, classId));
        source.MapiProperties.Add(new MapiProperty(0x66AC, MapiPropertyType.Integer32, 42));
        source.TnefAttributes.Add(new TnefAttribute(TnefAttributeLevel.Message, 0x0006F001, new byte[] { 7, 8 }));
        source.Attachments.Add(new EmailAttachment {
            FileName = "data.bin", ContentType = "application/octet-stream", Content = new byte[] { 1, 2, 3 }, Length = 3
        });
        source.Attachments.Add(new EmailAttachment { FileName = "child.dat", EmbeddedDocument = child });
        var storage = new EmailAttachment { FileName = "ole.dat", MapiAttachMethod = 6 };
        storage.StructuredStorageStreams["Contents"] = new byte[] { 9, 8, 7 };
        source.Attachments.Add(storage);

        byte[] first = new EmailDocumentWriter().WriteToBytes(source, EmailFileFormat.Tnef);
        byte[] second = new EmailDocumentWriter().WriteToBytes(source, EmailFileFormat.Tnef);
        EmailReadResult result = new EmailDocumentReader().Read(first);

        Assert.Equal(first, second);
        Assert.Equal(EmailFileFormat.Tnef, result.Document.Format);
        Assert.Equal("TNEF subject", result.Document.Subject);
        Assert.Equal("TNEF body", result.Document.Body.Text);
        Assert.True(result.Document.Appointment != null,
            string.Join(Environment.NewLine, result.Diagnostics.Select(diagnostic => string.Concat(diagnostic.Code, ": ", diagnostic.Message))));
        Assert.Equal("Room", result.Document.Appointment!.Location);
        Assert.Equal("to@example.com", Assert.Single(result.Document.Recipients).Address.Address);
        Assert.Equal(new byte[] { 1, 2, 3 }, result.Document.Attachments[0].Content);
        Assert.Equal("child", result.Document.Attachments[1].EmbeddedDocument!.Subject);
        Assert.True(result.Document.Attachments[1].Length > 0);
        Assert.Equal(new byte[] { 9, 8, 7 }, result.Document.Attachments[2].StructuredStorageStreams["Contents"]);
        Assert.Contains(result.Document.TnefAttributes, attribute => attribute.Tag == 0x0006F001);
        Assert.Equal(classId, result.Document.MapiProperties.Single(property => property.PropertyId == 0x66AB).Value);
        Assert.Equal(42, result.Document.MapiProperties.Single(property => property.PropertyId == 0x66AC).Value);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Severity == EmailDiagnosticSeverity.Error);
    }

    [Fact]
    public void OutputIsAcceptedByMimeKitTnefReaderOracle() {
        EmailDocument source = new EmailDocument { Format = EmailFileFormat.Tnef, Subject = "oracle" };
        source.Body.Text = "body";
        source.Attachments.Add(new EmailAttachment { FileName = "a.txt", Content = Encoding.UTF8.GetBytes("a"), Length = 1 });
        byte[] bytes = new EmailDocumentWriter().WriteToBytes(source, EmailFileFormat.Tnef);

        using MemoryStream stream = new MemoryStream(bytes);
        using var reader = new global::MimeKit.Tnef.TnefReader(stream);
        int count = 0;
        var compliance = new List<string>();
        while (reader.ReadNextAttribute()) {
            count++;
            if (reader.ComplianceStatus != TnefComplianceStatus.Compliant) {
                compliance.Add(string.Concat(reader.AttributeTag.ToString(), ": ", reader.ComplianceStatus.ToString()));
                reader.ResetComplianceStatus();
            }
        }

        Assert.True(count >= 6);
        Assert.Equal(0x00010000, reader.TnefVersion);
        Assert.True(compliance.Count == 0, string.Join(Environment.NewLine, compliance));
    }

    [Fact]
    public void RoundTripsCompressedRtfThroughTnefMapiProperties() {
        const string rtf = "{\\rtf1\\ansi TNEF RTF body\\par}";
        var source = new EmailDocument { Format = EmailFileFormat.Tnef, Subject = "rtf" };
        source.Body.Rtf = rtf;

        byte[] bytes = new EmailDocumentWriter().WriteToBytes(source, EmailFileFormat.Tnef);
        EmailReadResult result = new EmailDocumentReader().Read(bytes);

        Assert.Equal(rtf, result.Document.Body.Rtf);
        Assert.Contains(result.Document.MapiProperties, property => property.PropertyId == 0x1009);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Severity == EmailDiagnosticSeverity.Error);
    }

    [Fact]
    public void ReportsChecksumDamageAndEnforcesAttributeLimit() {
        EmailDocument source = new EmailDocument { Format = EmailFileFormat.Tnef, Subject = "checksum" };
        byte[] bytes = new EmailDocumentWriter().WriteToBytes(source, EmailFileFormat.Tnef);
        bytes[bytes.Length - 1] ^= 0x01;

        EmailReadResult damaged = new EmailDocumentReader().Read(bytes);

        Assert.Contains(damaged.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_TNEF_CHECKSUM_MISMATCH");
        Assert.Throws<EmailLimitExceededException>(() => new EmailDocumentReader(
            new EmailReaderOptions(maxTnefAttributeCount: 1)).Read(bytes));
    }

    [Fact]
    public void RejectsMalformedMapiAttributesBeforeRetainingRawPayload() {
        byte[] malformedProperties = new byte[4096];
        MsgBinary.WriteUInt32(malformedProperties, 0, 1);
        using var stream = new MemoryStream();
        using (var writer = new BinaryWriter(stream, Encoding.UTF8, leaveOpen: true)) {
            writer.Write(TnefConstants.Signature);
            writer.Write((ushort)1);
            WriteAttribute(writer, TnefAttributeLevel.Message, TnefConstants.MessageProperties, malformedProperties);
        }

        EmailReadResult result = new EmailDocumentReader(
            new EmailReaderOptions(maxDecodedPropertyBytes: 16)).Read(stream.ToArray());

        Assert.Equal(EmailFileFormat.Tnef, result.Document.Format);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_TNEF_MAPI_PREFLIGHT_INVALID" &&
            diagnostic.Severity == EmailDiagnosticSeverity.Error);
        Assert.DoesNotContain(result.Document.TnefAttributes,
            attribute => attribute.Tag == TnefConstants.MessageProperties);
    }

    [Fact]
    public void AppliesTransportHeaderRecipientsWhenRecipientTableIsAbsent() {
        var source = new EmailDocument { Format = EmailFileFormat.Tnef, Subject = "transport recipients" };
        source.Headers.Add(new EmailHeader("To", "Primary <primary@example.test>"));
        source.Headers.Add(new EmailHeader("Cc", "Copy <copy@example.test>"));

        byte[] bytes = new EmailDocumentWriter().WriteToBytes(source, EmailFileFormat.Tnef);
        EmailDocument document = new EmailDocumentReader().Read(bytes).Document;

        Assert.Contains(document.Recipients, recipient => recipient.Kind == EmailRecipientKind.To &&
            recipient.Address.Address == "primary@example.test");
        Assert.Contains(document.Recipients, recipient => recipient.Kind == EmailRecipientKind.Cc &&
            recipient.Address.Address == "copy@example.test");
    }

    [Fact]
    public void DecodesAndEncodesTnefString8UsingTheNumericCodePage() {
        var source = new[] { new MapiProperty(0x66AB, MapiPropertyType.String8, "日本") };
        var diagnostics = new List<EmailDiagnostic>();
        byte[] bytes = TnefMapiCodec.WriteProperties(source, 932, diagnostics, "tnef/mapi");
        var state = new MsgParserState(EmailReaderOptions.Default, diagnostics, CancellationToken.None);

        MapiProperty property = Assert.Single(TnefMapiCodec.ReadProperties(bytes, 932, state, "tnef/mapi"));

        Assert.Equal("日本", property.Value);
        Assert.DoesNotContain(diagnostics, diagnostic => diagnostic.Code == "EMAIL_MIME_CHARSET_UNSUPPORTED");
    }

    [Fact]
    public void PreservesRetainedString8RawBytesUntilTheyAreExplicitlyCleared() {
        byte[] raw = { 0x80, 0x00 };
        var source = new MapiProperty(0x66AB, MapiPropertyType.String8, "replacement") { RawData = raw };
        var diagnostics = new List<EmailDiagnostic>();
        byte[] bytes = TnefMapiCodec.WriteProperties(new[] { source }, 1252, diagnostics, "tnef/mapi");
        var state = new MsgParserState(EmailReaderOptions.Default, diagnostics, CancellationToken.None);

        MapiProperty property = Assert.Single(TnefMapiCodec.ReadProperties(bytes, 1252, state, "tnef/mapi"));

        Assert.Equal(raw, property.RawData);
        Assert.Equal("€", property.Value);
    }

    [Fact]
    public void RetainsDepthLimitedEmbeddedTnefForTnefAndMsgWriting() {
        var child = new EmailDocument { Format = EmailFileFormat.Tnef, Subject = "opaque TNEF child" };
        child.Body.Text = "inside";
        var parent = new EmailDocument { Format = EmailFileFormat.Tnef, Subject = "parent" };
        parent.Attachments.Add(new EmailAttachment { FileName = "child.dat", EmbeddedDocument = child });
        byte[] source = new EmailDocumentWriter().WriteToBytes(parent, EmailFileFormat.Tnef);

        EmailReadResult limited = new EmailDocumentReader(new EmailReaderOptions(maxNestedMessageDepth: 0)).Read(source);
        EmailAttachment opaque = Assert.Single(limited.Document.Attachments);
        Assert.Null(opaque.EmbeddedDocument);
        Assert.NotNull(opaque.Content);
        Assert.True(opaque.Length > 0);
        Assert.Contains(limited.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_TNEF_NESTED_MESSAGE_LIMIT");

        EmailAttachment metadataOnly = Assert.Single(new EmailDocumentReader(new EmailReaderOptions(
            maxNestedMessageDepth: 0, includeAttachmentContent: false)).Read(source).Document.Attachments);
        Assert.Null(metadataOnly.Content);
        Assert.True(metadataOnly.Length > 0);

        byte[] rewrittenTnef = new EmailDocumentWriter().WriteToBytes(limited.Document, EmailFileFormat.Tnef);
        EmailAttachment tnefRoundTrip = Assert.Single(new EmailDocumentReader().Read(rewrittenTnef).Document.Attachments);
        Assert.Equal("opaque TNEF child", tnefRoundTrip.EmbeddedDocument!.Subject);

        byte[] rewrittenMsg = new EmailDocumentWriter().WriteToBytes(
            limited.Document, EmailFileFormat.OutlookMsg, out EmailWriteResult msgWriteResult);
        EmailAttachment msgRoundTrip = Assert.Single(new EmailDocumentReader().Read(rewrittenMsg).Document.Attachments);
        Assert.DoesNotContain(msgWriteResult.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_ATTACHMENT_CONTENT_UNAVAILABLE");
        Assert.Equal("opaque TNEF child", msgRoundTrip.EmbeddedDocument!.Subject);
    }

    [Fact]
    public void ReplacesUnencodableLegacyStringsWithDiagnosticsInsteadOfThrowing() {
        var source = new EmailDocument {
            Format = EmailFileFormat.Tnef,
            OutlookCodePage = 1252,
            Subject = "emoji 😀"
        };
        source.Attachments.Add(new EmailAttachment { FileName = "資料.txt", Content = new byte[] { 1 }, Length = 1 });

        byte[] bytes = new EmailDocumentWriter().WriteToBytes(
            source, EmailFileFormat.Tnef, out EmailWriteResult writeResult);
        EmailDocument roundTrip = new EmailDocumentReader().Read(bytes).Document;

        Assert.Contains(writeResult.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_TNEF_STRING8_CHARACTER_UNENCODABLE");
        Assert.Contains("?", roundTrip.Subject, StringComparison.Ordinal);
        Assert.NotEmpty(bytes);
    }

    [Fact]
    public void ReplacesUnencodableMapiString8ValuesWithDiagnostics() {
        var source = new EmailDocument {
            Format = EmailFileFormat.Tnef,
            OutlookCodePage = 1252,
            Subject = "Mapi fallback"
        };
        source.MapiProperties.Add(new MapiProperty(0x66AB, MapiPropertyType.String8, "日本"));

        byte[] bytes = new EmailDocumentWriter().WriteToBytes(
            source, EmailFileFormat.Tnef, out EmailWriteResult writeResult);
        MapiProperty property = new EmailDocumentReader().Read(bytes).Document.MapiProperties
            .Single(item => item.PropertyId == 0x66AB);

        Assert.Contains(writeResult.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_TNEF_MAPI_STRING8_CHARACTER_UNENCODABLE");
        Assert.Contains("?", Assert.IsType<string>(property.Value), StringComparison.Ordinal);
    }

    [Fact]
    public void FallsBackFromUnsupportedPreservedCodePagesWhenWriting() {
        var source = new EmailDocument {
            Format = EmailFileFormat.Tnef,
            OutlookCodePage = 999999,
            Subject = "unsupported code page"
        };

        byte[] bytes = new EmailDocumentWriter().WriteToBytes(source, EmailFileFormat.Tnef);
        EmailDocument roundTrip = new EmailDocumentReader().Read(bytes).Document;

        Assert.Equal(1252, roundTrip.OutlookCodePage);
        Assert.Equal(source.Subject, roundTrip.Subject);
    }

    [Fact]
    public void WritesFixedWidthNullPropertiesWithoutDesynchronizingFollowingValues() {
        var source = new[] {
            new MapiProperty(0x66AA, MapiPropertyType.Null, null),
            new MapiProperty(0x66AB, MapiPropertyType.Integer32, 42)
        };
        var diagnostics = new List<EmailDiagnostic>();
        byte[] bytes = TnefMapiCodec.WriteProperties(source, 1252, diagnostics, "tnef/mapi");
        var state = new MsgParserState(EmailReaderOptions.Default, diagnostics, CancellationToken.None);

        List<MapiProperty> properties = TnefMapiCodec.ReadProperties(bytes, 1252, state, "tnef/mapi");

        Assert.Equal(2, properties.Count);
        Assert.Null(properties[0].Value);
        Assert.Equal(42, properties[1].Value);
        Assert.DoesNotContain(diagnostics, diagnostic => diagnostic.Code == "EMAIL_TNEF_MAPI_TRUNCATED");
    }

    [Fact]
    public void CountsMapiAttachmentBytesWhenPayloadRetentionIsDisabled() {
        byte[] payload = Enumerable.Range(0, 16).Select(value => (byte)value).ToArray();
        byte[] bytes = CreateTnefWithMapiAttachment(payload);

        EmailDocument document = new EmailDocumentReader(
            new EmailReaderOptions(includeAttachmentContent: false)).Read(bytes).Document;

        EmailAttachment attachment = Assert.Single(document.Attachments);
        Assert.Null(attachment.Content);
        Assert.Equal(payload.Length, attachment.Length);
        EmailLimitExceededException exception = Assert.Throws<EmailLimitExceededException>(() =>
            new EmailDocumentReader(new EmailReaderOptions(
                maxAttachmentBytes: payload.Length - 1,
                includeAttachmentContent: false)).Read(bytes));
        Assert.Equal(nameof(EmailReaderOptions.MaxAttachmentBytes), exception.LimitName);
        Assert.Equal(payload.Length, exception.ActualValue);
    }

    [Fact]
    public void DecodesTopLevelTnefAttributesUsingTheNumericCodePage() {
        byte[] codePage = new byte[8];
        MsgBinary.WriteUInt32(codePage, 0, 932);
        byte[] subject = MsgValueWriter.EncodeString8("日本\0", 932);
        using var stream = new MemoryStream();
        using (var writer = new BinaryWriter(stream, Encoding.UTF8, leaveOpen: true)) {
            writer.Write(TnefConstants.Signature);
            writer.Write((ushort)1);
            WriteAttribute(writer, TnefAttributeLevel.Message, TnefConstants.OemCodePage, codePage);
            WriteAttribute(writer, TnefAttributeLevel.Message, TnefConstants.Subject, subject);
        }

        EmailReadResult result = new EmailDocumentReader().Read(stream.ToArray());

        Assert.Equal("日本", result.Document.Subject);
        Assert.DoesNotContain(result.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_MIME_CHARSET_UNSUPPORTED");
    }

    [Fact]
    public void SkipsNamedRangePropertiesWithoutDescriptorsAndReportsDataLoss() {
        var source = new EmailDocument { Format = EmailFileFormat.Tnef, Subject = "named properties" };
        source.MapiProperties.Add(new MapiProperty(0x8001, MapiPropertyType.Unicode, "invalid"));
        source.MapiProperties.Add(new MapiProperty(0x66AA, MapiPropertyType.Integer32, 42));

        byte[] bytes = new EmailDocumentWriter().WriteToBytes(
            source, EmailFileFormat.Tnef, out EmailWriteResult writeResult);
        EmailReadResult readResult = new EmailDocumentReader().Read(bytes);

        Assert.True(writeResult.HasErrors);
        Assert.Contains(writeResult.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_TNEF_NAMED_PROPERTY_DESCRIPTOR_MISSING");
        Assert.DoesNotContain(readResult.Document.MapiProperties, property => property.PropertyId == 0x8001);
        Assert.Equal(42, readResult.Document.MapiProperties
            .Single(property => property.PropertyId == 0x66AA).Value);
    }

    [Fact]
    public void AppliesEmailCompoundLimitsToTnefOleAttachments() {
        var source = new EmailDocument { Format = EmailFileFormat.Tnef, Subject = "bounded OLE" };
        var attachment = new EmailAttachment { FileName = "object.ole", MapiAttachMethod = 6 };
        attachment.StructuredStorageStreams["Contents"] = new byte[] { 1, 2, 3 };
        source.Attachments.Add(attachment);
        byte[] bytes = new EmailDocumentWriter().WriteToBytes(source, EmailFileFormat.Tnef);

        EmailReadResult result = new EmailDocumentReader(
            new EmailReaderOptions(maxCompoundDirectoryEntries: 1)).Read(bytes);

        EmailAttachment parsed = Assert.Single(result.Document.Attachments);
        Assert.Empty(parsed.StructuredStorageStreams);
        Assert.NotNull(parsed.Content);
        Assert.Equal(parsed.Content!.LongLength, parsed.Length);
        Assert.Contains(result.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_TNEF_COMPOUND_ATTACHMENT_INVALID");

        EmailAttachment metadataOnly = Assert.Single(new EmailDocumentReader(new EmailReaderOptions(
            includeAttachmentContent: false, maxCompoundDirectoryEntries: 1)).Read(bytes).Document.Attachments);
        Assert.Null(metadataOnly.Content);
        Assert.True(metadataOnly.Length > 0);

        byte[] rewritten = new EmailDocumentWriter().WriteToBytes(result.Document, EmailFileFormat.Tnef);
        EmailAttachment reparsed = Assert.Single(new EmailDocumentReader(
            new EmailReaderOptions(maxCompoundDirectoryEntries: 1)).Read(rewritten).Document.Attachments);
        Assert.Equal(parsed.Content, reparsed.Content);
    }

    [Fact]
    public void RejectsOversizedTnefOlePayloadBeforeParsingItsStorage() {
        var source = new EmailDocument { Format = EmailFileFormat.Tnef, Subject = "bounded OLE bytes" };
        var attachment = new EmailAttachment { FileName = "object.ole", MapiAttachMethod = 6 };
        attachment.StructuredStorageStreams["Contents"] = new byte[] { 1, 2, 3 };
        source.Attachments.Add(attachment);
        byte[] bytes = new EmailDocumentWriter().WriteToBytes(source, EmailFileFormat.Tnef);

        EmailLimitExceededException exception = Assert.Throws<EmailLimitExceededException>(() =>
            new EmailDocumentReader(new EmailReaderOptions(maxAttachmentBytes: 512)).Read(bytes));

        Assert.Equal(nameof(EmailReaderOptions.MaxAttachmentBytes), exception.LimitName);
        Assert.True(exception.ActualValue > 512);
    }

    private static byte[] CreateTnefWithMapiAttachment(byte[] payload) {
        byte[] rendition = new byte[14];
        MsgBinary.WriteUInt16(rendition, 0, 1);
        byte[] properties = TnefMapiCodec.WriteProperties(new[] {
            new MapiProperty(0x3701, MapiPropertyType.Binary, payload),
            new MapiProperty(0x3705, MapiPropertyType.Integer32, 1)
        }, 1252, new List<EmailDiagnostic>(), "tnef/attachment/mapi");
        using var stream = new MemoryStream();
        using (var writer = new BinaryWriter(stream, Encoding.UTF8, leaveOpen: true)) {
            writer.Write(TnefConstants.Signature);
            writer.Write((ushort)1);
            WriteAttribute(writer, TnefAttributeLevel.Attachment, TnefConstants.AttachRendData, rendition);
            WriteAttribute(writer, TnefAttributeLevel.Attachment, TnefConstants.AttachmentProperties, properties);
        }
        return stream.ToArray();
    }

    private static void WriteAttribute(BinaryWriter writer, TnefAttributeLevel level, uint tag, byte[] data) {
        writer.Write((byte)level);
        writer.Write(tag);
        writer.Write(unchecked((uint)data.Length));
        writer.Write(data);
        writer.Write(unchecked((ushort)data.Sum(value => value)));
    }
}
