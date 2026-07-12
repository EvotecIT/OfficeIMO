using OfficeIMO.Email;
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class EmailMsgRoundTripTests {
    [Fact]
    public void RoundTripsMessageRecipientsAttachmentsEmbeddedMessagesAndUnknownProperties() {
        EmailDocument child = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            Subject = "Embedded",
            MessageClass = "IPM.Note"
        };
        child.Body.Text = "inside";
        EmailDocument source = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            Subject = "MSG subject",
            MessageClass = "IPM.Note",
            From = new EmailAddress("from@example.com", "From Person"),
            Sender = new EmailAddress("sender@example.com", "Sender Person"),
            MessageId = "id@example.com",
            Date = new DateTimeOffset(2026, 7, 10, 12, 30, 0, TimeSpan.Zero),
            ReceivedDate = new DateTimeOffset(2026, 7, 10, 12, 31, 0, TimeSpan.Zero)
        };
        source.Body.Text = "plain body";
        source.Body.Html = "<p>html body</p>";
        source.Headers.Add(new EmailHeader("X-Test", "value"));
        source.Recipients.Add(new EmailRecipient(EmailRecipientKind.To, new EmailAddress("to@example.com", "To")));
        source.Recipients.Add(new EmailRecipient(EmailRecipientKind.Cc, new EmailAddress("cc@example.com", "Cc")));
        source.MapiProperties.Add(new MapiProperty(0x66AA, MapiPropertyType.MultipleUnicode,
            new object[] { "one", "dwa" }));
        source.MapiProperties.Add(new MapiProperty(0x8000, MapiPropertyType.Unicode, "named value",
            name: new MapiNamedProperty(MsgProjection.PsetidCommon, 0x85FF)));
        source.Attachments.Add(new EmailAttachment {
            FileName = "data.bin",
            ContentType = "application/octet-stream",
            Content = new byte[] { 1, 2, 3, 4 },
            Length = 4
        });
        source.Attachments.Add(new EmailAttachment {
            FileName = "child.msg",
            ContentType = "application/vnd.ms-outlook",
            EmbeddedDocument = child
        });
        var structured = new EmailAttachment { FileName = "object.ole", MapiAttachMethod = 6 };
        structured.StructuredStorageStreams["Contents"] = new byte[] { 9, 8, 7 };
        structured.StructuredStorageStreams["Nested/Metadata"] = Encoding.UTF8.GetBytes("meta");
        source.Attachments.Add(structured);

        byte[] first = new EmailDocumentWriter().WriteToBytes(source, EmailFileFormat.OutlookMsg);
        byte[] second = new EmailDocumentWriter().WriteToBytes(source, EmailFileFormat.OutlookMsg);
        EmailReadResult result = new EmailDocumentReader().Read(first);

        Assert.Equal(first, second);
        Assert.Equal(EmailFileFormat.OutlookMsg, result.Document.Format);
        Assert.Equal(source.Subject, result.Document.Subject);
        Assert.Equal(source.Body.Text, result.Document.Body.Text);
        Assert.Equal(source.Body.Html, result.Document.Body.Html);
        Assert.Equal(3, result.Document.MapiProperties.Single(property => property.PropertyId == 0x1016).Value);
        Assert.Equal("from@example.com", result.Document.From!.Address);
        Assert.Equal("sender@example.com", result.Document.Sender!.Address);
        Assert.Equal(2, result.Document.Recipients.Count);
        Assert.Equal(3, result.Document.Attachments.Count);
        Assert.Equal(new byte[] { 1, 2, 3, 4 }, result.Document.Attachments[0].Content);
        Assert.Equal("Embedded", result.Document.Attachments[1].EmbeddedDocument!.Subject);
        Assert.Equal("meta", Encoding.UTF8.GetString(result.Document.Attachments[2].StructuredStorageStreams["Nested/Metadata"]));
        Assert.Equal(new object[] { "one", "dwa" }, result.Document.MapiProperties.Single(property => property.PropertyId == 0x66AA).Value);
        Assert.Equal("named value", result.Document.MapiProperties.Single(property => property.Name?.LocalId == 0x85FF).Value);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Severity == EmailDiagnosticSeverity.Error);
    }

    [Fact]
    public void OutputIsReadableByMsgReaderOracle() {
        EmailDocument source = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            Subject = "Oracle subject",
            From = new EmailAddress("sender@example.com", "Sender")
        };
        source.Body.Text = "Oracle body";
        source.Recipients.Add(new EmailRecipient(EmailRecipientKind.To, new EmailAddress("receiver@example.com", "Receiver")));
        source.Attachments.Add(new EmailAttachment {
            FileName = "a.txt",
            ContentType = "text/plain",
            Content = Encoding.UTF8.GetBytes("attachment"),
            Length = 10
        });
        byte[] bytes = new EmailDocumentWriter().WriteToBytes(source, EmailFileFormat.OutlookMsg);

        using MemoryStream stream = new MemoryStream(bytes);
        using var oracle = new global::MsgReader.Outlook.Storage.Message(stream, FileAccess.Read, true);

        Assert.Equal("Oracle subject", oracle.Subject);
        Assert.Equal("Oracle body", oracle.BodyText!.TrimEnd());
        Assert.Single(oracle.Recipients!);
        Assert.Single(oracle.Attachments!);
    }

    [Fact]
    public void OutputContainsRequiredMsgRootAndNamedPropertyMappingStorage() {
        EmailDocument source = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            Subject = "MSG structure"
        };
        source.Body.Text = "MSG body";
        byte[] bytes = new EmailDocumentWriter().WriteToBytes(source, EmailFileFormat.OutlookMsg);

        using MemoryStream stream = new MemoryStream(bytes);
        using var oracle = OpenMcdf.RootStorage.Open(stream, OpenMcdf.StorageModeFlags.LeaveOpen);
        Assert.Equal(new Guid("00020D0B-0000-0000-C000-000000000046"), oracle.EntryInfo.CLSID);
        OpenMcdf.Storage namedProperties = oracle.OpenStorage("__nameid_version1.0");
        using OpenMcdf.CfbStream guidStream = namedProperties.OpenStream("__substg1.0_00020102");
        using OpenMcdf.CfbStream entryStream = namedProperties.OpenStream("__substg1.0_00030102");
        using OpenMcdf.CfbStream stringStream = namedProperties.OpenStream("__substg1.0_00040102");
        Assert.Equal(32, guidStream.Length);
        Assert.Equal(16, entryStream.Length);
        Assert.Equal(32, stringStream.Length);
        string[] lookupNames = namedProperties.EnumerateEntries()
            .Select(entry => entry.Name)
            .Where(name => name != "__substg1.0_00020102" && name != "__substg1.0_00030102" &&
                name != "__substg1.0_00040102")
            .OrderBy(name => name, StringComparer.Ordinal)
            .ToArray();
        Assert.Equal(2, lookupNames.Length);
        Assert.Contains("__substg1.0_10010102", lookupNames);
        Assert.Contains("__substg1.0_101D0102", lookupNames);
        foreach (string lookupName in lookupNames) {
            using OpenMcdf.CfbStream lookup = namedProperties.OpenStream(lookupName);
            Assert.Equal(8, lookup.Length);
        }
        using (OpenMcdf.CfbStream acceptLanguageLookup = namedProperties.OpenStream("__substg1.0_101D0102")) {
            byte[] entry = new byte[8];
            ReadFully(acceptLanguageLookup, entry);
            Assert.Equal(new byte[] { 0x9E, 0x53, 0xE9, 0x83, 0x09, 0x00, 0x01, 0x00 }, entry);
        }

        EmailDocument roundTrip = new EmailDocumentReader().Read(bytes).Document;
        Assert.Equal(new DateTimeOffset(1970, 1, 1, 0, 0, 0, TimeSpan.Zero),
            roundTrip.MessageMetadata.CreatedDate);
        Assert.Equal(1, roundTrip.MapiProperties.Single(property => property.PropertyId == 0x1016).Value);
        MapiProperty sideEffects = Assert.Single(roundTrip.MapiProperties, property =>
            property.Name?.PropertySet == MsgProjection.PsetidCommon && property.Name.LocalId == 0x8510);
        Assert.Equal(0, sideEffects.Value);
        MapiProperty acceptLanguage = Assert.Single(roundTrip.MapiProperties, property =>
            property.Name?.PropertySet == MsgProjection.PsInternetHeaders &&
            string.Equals(property.Name.Name, "acceptlanguage", StringComparison.OrdinalIgnoreCase));
        Assert.Equal("en-US", acceptLanguage.Value);
    }

    [Fact]
    public void StringPropertyStreamsContainAdvertisedTerminators() {
        var source = new EmailDocument { Subject = "Terminated subject" };
        byte[] bytes = new EmailDocumentWriter().WriteToBytes(source, EmailFileFormat.OutlookMsg);

        using MemoryStream stream = new MemoryStream(bytes);
        using var oracle = OpenMcdf.RootStorage.Open(stream, OpenMcdf.StorageModeFlags.LeaveOpen);
        using OpenMcdf.CfbStream subjectStream = oracle.OpenStream("__substg1.0_0037001F");
        byte[] subject = new byte[subjectStream.Length];
        ReadFully(subjectStream, subject);
        using OpenMcdf.CfbStream propertiesStream = oracle.OpenStream("__properties_version1.0");
        byte[] properties = new byte[propertiesStream.Length];
        ReadFully(propertiesStream, properties);

        uint advertisedLength = 0;
        for (int offset = 32; offset + 16 <= properties.Length; offset += 16) {
            if (BitConverter.ToUInt32(properties, offset) == 0x0037001FU) {
                advertisedLength = BitConverter.ToUInt32(properties, offset + 8);
                break;
            }
        }

        Assert.Equal(unchecked((uint)subject.Length), advertisedLength);
        Assert.True(subject.Length >= 2);
        Assert.Equal(0, subject[subject.Length - 2]);
        Assert.Equal(0, subject[subject.Length - 1]);
    }

    [Fact]
    public void NamedPropertyLookupStreamsGroupHashCollisions() {
        EmailDocument source = new EmailDocument { Subject = "NameID collisions" };
        source.MapiProperties.Add(new MapiProperty(0x8000, MapiPropertyType.Integer32, 1,
            name: new MapiNamedProperty(MsgProjection.PsetidTask, 0x8006)));
        source.MapiProperties.Add(new MapiProperty(0x8001, MapiPropertyType.Integer32, 2,
            name: new MapiNamedProperty(MsgProjection.PsetidTask, 0x8019)));
        byte[] bytes = new EmailDocumentWriter().WriteToBytes(source, EmailFileFormat.OutlookMsg);

        using MemoryStream stream = new MemoryStream(bytes);
        using var oracle = OpenMcdf.RootStorage.Open(stream, OpenMcdf.StorageModeFlags.LeaveOpen);
        OpenMcdf.Storage namedProperties = oracle.OpenStorage("__nameid_version1.0");
        using OpenMcdf.CfbStream lookup = namedProperties.OpenStream("__substg1.0_10010102");
        byte[] entries = new byte[lookup.Length];
        ReadFully(lookup, entries);

        uint[] identifiers = Enumerable.Range(0, entries.Length / 8)
            .Select(index => BitConverter.ToUInt32(entries, index * 8))
            .ToArray();
        Assert.Contains(0x8006U, identifiers);
        Assert.Contains(0x8019U, identifiers);
    }

    [Fact]
    public void ReaderCanSkipMsgAttachmentBytes() {
        EmailDocument source = new EmailDocument { Format = EmailFileFormat.OutlookMsg, Subject = "skip" };
        source.Attachments.Add(new EmailAttachment { FileName = "a.bin", Content = new byte[] { 1, 2, 3 }, Length = 3 });
        byte[] bytes = new EmailDocumentWriter().WriteToBytes(source, EmailFileFormat.OutlookMsg);

        EmailDocument parsed = new EmailDocumentReader(new EmailReaderOptions(includeAttachmentContent: false)).Read(bytes).Document;

        EmailAttachment attachment = Assert.Single(parsed.Attachments);
        Assert.Equal(3, attachment.Length);
        Assert.Null(attachment.Content);
        Assert.Null(attachment.MapiProperties.Single(property => property.PropertyId == 0x3701).RawData);
    }

    [Fact]
    public void ReaderCanSkipOleStorageBytesWhileKeepingDeclaredLength() {
        var source = new EmailDocument { Format = EmailFileFormat.OutlookMsg, Subject = "skip OLE" };
        var attachment = new EmailAttachment { FileName = "object.bin", MapiAttachMethod = 6 };
        attachment.StructuredStorageStreams["Contents"] = new byte[] { 1, 2, 3, 4 };
        source.Attachments.Add(attachment);
        byte[] bytes = new EmailDocumentWriter().WriteToBytes(source, EmailFileFormat.OutlookMsg);

        EmailDocument parsed = new EmailDocumentReader(
            new EmailReaderOptions(includeAttachmentContent: false)).Read(bytes).Document;

        EmailAttachment result = Assert.Single(parsed.Attachments);
        Assert.Equal(4, result.Length);
        Assert.Empty(result.StructuredStorageStreams);
    }

    [Fact]
    public void WriterRetainsOpaqueMethodSixContentInsideObjectStorage() {
        byte[] opaque = Encoding.ASCII.GetBytes("not-a-readable-compound-payload");
        var source = new EmailDocument { Format = EmailFileFormat.OutlookMsg, Subject = "opaque OLE" };
        source.Attachments.Add(new EmailAttachment {
            FileName = "object.ole",
            MapiAttachMethod = 6,
            Content = opaque,
            Length = opaque.Length
        });

        byte[] bytes = new EmailDocumentWriter().WriteToBytes(
            source, EmailFileFormat.OutlookMsg, out EmailWriteResult writeResult);
        EmailAttachment roundTrip = Assert.Single(new EmailDocumentReader().Read(bytes).Document.Attachments);

        Assert.Equal(opaque, roundTrip.StructuredStorageStreams["CONTENTS"]);
        Assert.Contains(writeResult.Diagnostics, diagnostic =>
            diagnostic.Code == "EMAIL_MSG_OPAQUE_STRUCTURED_CONTENT_WRAPPED" &&
            diagnostic.Severity == EmailDiagnosticSeverity.Warning);
        Assert.DoesNotContain(writeResult.Diagnostics, diagnostic =>
            diagnostic.Code == "EMAIL_ATTACHMENT_CONTENT_UNAVAILABLE");
    }

    [Fact]
    public void WriterRetainsOpaqueMethodFiveContentInsideObjectStorage() {
        byte[] opaque = Encoding.ASCII.GetBytes("not-a-readable-embedded-message");
        var source = new EmailDocument { Format = EmailFileFormat.OutlookMsg, Subject = "opaque embedded" };
        source.Attachments.Add(new EmailAttachment {
            FileName = "embedded.msg",
            MapiAttachMethod = 5,
            Content = opaque,
            Length = opaque.Length
        });

        byte[] bytes = new EmailDocumentWriter().WriteToBytes(
            source, EmailFileFormat.OutlookMsg, out EmailWriteResult writeResult);
        EmailReadResult readResult = new EmailDocumentReader().Read(bytes);
        EmailAttachment roundTrip = Assert.Single(readResult.Document.Attachments);

        Assert.Null(roundTrip.EmbeddedDocument);
        Assert.Equal(opaque, roundTrip.StructuredStorageStreams["CONTENTS"]);
        Assert.Contains(writeResult.Diagnostics, diagnostic =>
            diagnostic.Code == "EMAIL_MSG_OPAQUE_EMBEDDED_CONTENT_WRAPPED" &&
            diagnostic.Severity == EmailDiagnosticSeverity.Warning);
        Assert.Contains(readResult.Diagnostics, diagnostic =>
            diagnostic.Code == "EMAIL_MSG_EMBEDDED_STORAGE_INVALID" &&
            diagnostic.Severity == EmailDiagnosticSeverity.Warning);
        Assert.DoesNotContain(writeResult.Diagnostics, diagnostic =>
            diagnostic.Code == "EMAIL_ATTACHMENT_CONTENT_UNAVAILABLE");
    }

    [Fact]
    public void LogicalAttachmentNamesDoNotUseFilesystemPathValidation() {
        var diagnostics = new List<EmailDiagnostic>();
        var attachment = new EmailAttachment { FileName = "report|2026.txt", Content = new byte[] { 1 }, Length = 1 };

        MsgPropertyBuilder properties = MsgWriter.CreateAttachmentProperties(
            attachment, 0, 1, diagnostics, "attachment[0]");

        Assert.Equal(".txt", properties.Properties.Single(property => property.PropertyId == 0x3703).Value);
        Assert.DoesNotContain(diagnostics, diagnostic => diagnostic.Severity == EmailDiagnosticSeverity.Error);
    }

    [Fact]
    public void MissingStructuredAttachmentStreamsProduceDataLossDiagnostics() {
        var source = new EmailDocument { Subject = "missing structured payload" };
        source.Attachments.Add(new EmailAttachment {
            FileName = "object.ole",
            MapiAttachMethod = 6,
            Length = 10
        });

        EmailDocumentWriter writer = new EmailDocumentWriter();
        writer.WriteToBytes(source, EmailFileFormat.OutlookMsg, out EmailWriteResult result);

        Assert.True(result.HasErrors);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_ATTACHMENT_CONTENT_UNAVAILABLE");
    }

    [Fact]
    public void MissingEmbeddedMsgDocumentProducesDataLossDiagnostics() {
        var source = new EmailDocument { Subject = "missing embedded payload" };
        source.Attachments.Add(new EmailAttachment {
            FileName = "embedded.msg",
            MapiAttachMethod = 5
        });

        new EmailDocumentWriter().WriteToBytes(
            source, EmailFileFormat.OutlookMsg, out EmailWriteResult result);

        Assert.True(result.HasErrors);
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "EMAIL_ATTACHMENT_CONTENT_UNAVAILABLE" &&
            diagnostic.Message.Contains("embedded MSG", StringComparison.Ordinal));
    }

    [Fact]
    public void NestedMsgChildrenAreNotProjectedAsParentFallbackAttachments() {
        var embedded = new EmailDocument { Subject = "embedded" };
        embedded.Attachments.Add(new EmailAttachment { FileName = "one.bin", Content = new byte[] { 1 }, Length = 1 });
        embedded.Attachments.Add(new EmailAttachment { FileName = "two.bin", Content = new byte[] { 2 }, Length = 1 });
        var source = new EmailDocument { Subject = "parent" };
        source.Attachments.Add(new EmailAttachment { FileName = "embedded.msg", EmbeddedDocument = embedded });

        EmailReadResult result = new EmailDocumentReader().Read(
            new EmailDocumentWriter().WriteToBytes(source, EmailFileFormat.OutlookMsg));

        EmailAttachment parentAttachment = Assert.Single(result.Document.Attachments);
        Assert.Equal(2, parentAttachment.EmbeddedDocument!.Attachments.Count);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_MSG_PROPERTIES_MISSING");
    }

    [Fact]
    public void RetainsDepthLimitedEmbeddedMsgAsOpaqueStorageForLaterWriting() {
        var child = new EmailDocument { Subject = "depth-limited child" };
        child.Body.Text = "inside";
        var parent = new EmailDocument { Subject = "parent" };
        parent.Attachments.Add(new EmailAttachment { FileName = "child.msg", EmbeddedDocument = child });
        byte[] source = new EmailDocumentWriter().WriteToBytes(parent, EmailFileFormat.OutlookMsg);

        EmailReadResult limited = new EmailDocumentReader(new EmailReaderOptions(maxNestedMessageDepth: 0)).Read(source);
        EmailAttachment opaque = Assert.Single(limited.Document.Attachments);
        Assert.Null(opaque.EmbeddedDocument);
        Assert.NotEmpty(opaque.StructuredStorageStreams);
        Assert.Contains(limited.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_MSG_NESTED_MESSAGE_LIMIT");

        byte[] rewritten = new EmailDocumentWriter().WriteToBytes(
            limited.Document, EmailFileFormat.OutlookMsg, out EmailWriteResult writeResult);
        EmailAttachment roundTrip = Assert.Single(new EmailDocumentReader().Read(rewritten).Document.Attachments);

        Assert.DoesNotContain(writeResult.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_ATTACHMENT_CONTENT_UNAVAILABLE");
        Assert.Equal("depth-limited child", roundTrip.EmbeddedDocument!.Subject);
        Assert.Equal("inside", roundTrip.EmbeddedDocument.Body.Text);
    }

    [Fact]
    public void ReadsMsgKitGeneratedMessageAsCompatibilityOracle() {
        string directory = Path.Combine(Path.GetTempPath(), string.Concat("OfficeIMO-Email-", Guid.NewGuid().ToString("N")));
        Directory.CreateDirectory(directory);
        string emlPath = Path.Combine(directory, "source.eml");
        string msgPath = Path.Combine(directory, "source.msg");
        const string eml = "From: Sender <sender@example.com>\r\nTo: Receiver <receiver@example.com>\r\n" +
            "Subject: MsgKit source\r\nMIME-Version: 1.0\r\nContent-Type: multipart/mixed; boundary=x\r\n\r\n" +
            "--x\r\nContent-Type: text/plain; charset=utf-8\r\n\r\nbody from MsgKit\r\n" +
            "--x\r\nContent-Type: application/octet-stream; name=a.bin\r\n" +
            "Content-Disposition: attachment; filename=a.bin\r\nContent-Transfer-Encoding: base64\r\n\r\nAQID\r\n--x--\r\n";
        try {
            File.WriteAllText(emlPath, eml, new UTF8Encoding(false));
            MsgKit.Converter.ConvertEmlToMsg(emlPath, msgPath);

            EmailReadResult result = new EmailDocumentReader().Read(msgPath);

            Assert.Equal("MsgKit source", result.Document.Subject);
            Assert.Contains("body from MsgKit", result.Document.Body.Text);
            Assert.Equal("receiver@example.com", Assert.Single(result.Document.Recipients).Address.Address);
            Assert.Equal(new byte[] { 1, 2, 3 }, Assert.Single(result.Document.Attachments).Content);
            Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Severity == EmailDiagnosticSeverity.Error);
        } finally {
            if (Directory.Exists(directory)) Directory.Delete(directory, true);
        }
    }

    private static void ReadFully(Stream stream, byte[] buffer) {
        int offset = 0;
        while (offset < buffer.Length) {
            int read = stream.Read(buffer, offset, buffer.Length - offset);
            if (read == 0) throw new EndOfStreamException();
            offset += read;
        }
    }
}
