namespace OfficeIMO.Email.Tests;

public sealed class EmailTransportIntegrityTests {
    [Theory]
    [InlineData("Content-MD5")]
    [InlineData("Digest")]
    [InlineData("Content-Digest")]
    [InlineData("Repr-Digest")]
    [InlineData("Content-Length")]
    public void RegeneratedHtmlAndAttachmentPayloadHeadersHaveSeparateLossEvidence(string header) {
        string mime = "MIME-Version: 1.0\r\nContent-Type: multipart/mixed; boundary=parts\r\n\r\n" +
            "--parts\r\nContent-Type: text/html\r\n" + header + ": original\r\n\r\n<p>body</p>\r\n" +
            "--parts\r\nContent-Type: application/octet-stream\r\nContent-Disposition: attachment; filename=file.bin\r\n" +
            header + ": original\r\n\r\nbytes\r\n--parts--\r\n";
        using var parsed = new EmailDocumentReader().Read(Encoding.ASCII.GetBytes(mime));
        var writer = new EmailDocumentWriter();
        byte[] output = writer.ToBytes(parsed.Document, EmailFileFormat.Eml, out var result);
        Assert.False(result.HasErrors);
        Assert.DoesNotContain(header + ": original", Encoding.ASCII.GetString(output));
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_PAYLOAD_METADATA_REMOVED" && diagnostic.Location == "headers/html");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_PAYLOAD_METADATA_REMOVED" && diagnostic.Location == "headers/attachment/0");
    }

    [Theory]
    [InlineData("DKIM-Signature")]
    [InlineData("DomainKey-Signature")]
    [InlineData("ARC-Message-Signature")]
    [InlineData("ARC-Seal")]
    public void RegenerationBlocksTransportSignaturesBeforeTouchingOutput(string header) {
        EmailDocument document = Message(header);
        using var output = new MemoryStream();
        output.WriteByte(42);
        EmailWriteResult result = new EmailDocumentWriter().Write(document, output);
        Assert.True(result.HasErrors);
        Assert.Equal(new byte[] { 42 }, output.ToArray());
        Assert.Equal(EmailConversionLossDisposition.Blocked, result.LossDisposition);
        Assert.Contains(result.Diagnostics, item => item.Code == "EMAIL_TRANSPORT_SIGNATURE_INVALIDATED");
        Assert.Throws<InvalidDataException>(() => document.ToBytes());
    }

    [Theory]
    [InlineData(OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures, false)]
    [InlineData(OfficeSignatureMutationPolicy.PreserveSignatureMarkup, true)]
    public void ExplicitPoliciesReportInvalidationAndNeverCopyOldPayloadDigests(OfficeSignatureMutationPolicy policy, bool keep) {
        EmailDocument document = Message("DKIM-Signature");
        document.Headers.Add(new EmailHeader("ARC-Authentication-Results", "i=1; example.test; dkim=pass"));
        string[] payloadHeaders = { "Content-Length", "Content-MD5", "Content-Digest", "Repr-Digest", "Digest" };
        foreach (string name in payloadHeaders) document.Headers.Add(new EmailHeader(name, "original"));
        int originalCount = document.Headers.Count;
        var writer = new EmailDocumentWriter(new EmailWriterOptions(policy));
        byte[] bytes = writer.ToBytes(document, EmailFileFormat.Eml, out EmailWriteResult result);
        using EmailReadResult parsed = new EmailDocumentReader().Read(bytes);
        Assert.False(result.HasErrors);
        Assert.Equal(EmailConversionLossDisposition.Accepted, result.LossDisposition);
        Assert.Equal(keep, parsed.Document.Headers.Any(item => item.Name == "DKIM-Signature"));
        Assert.Equal(keep, parsed.Document.Headers.Any(item => item.Name == "ARC-Authentication-Results"));
        Assert.DoesNotContain(parsed.Document.Headers, item => payloadHeaders.Contains(item.Name, StringComparer.OrdinalIgnoreCase));
        Assert.Contains(result.Diagnostics, item => item.Code == "EMAIL_PAYLOAD_METADATA_REMOVED");
        Assert.Contains(result.Diagnostics, item => item.Code == "EMAIL_TRANSPORT_SIGNATURE_INVALIDATED");
        Assert.Equal(originalCount, document.Headers.Count);
    }

    [Fact]
    public void UnchangedRawSourceKeepsTransportEvidenceExactlyAndMutationBlocksReuse() {
        byte[] source = Encoding.ASCII.GetBytes("Subject: Original\r\nDKIM-Signature: v=1; bh=AAAA; b=BBBB\r\nContent-Length: 4\r\nContent-Type: text/plain\r\n\r\nbody");
        using EmailReadResult read = new EmailDocumentReader(new EmailReaderOptions(preserveRawSource: true)).Read(source);
        var writer = new EmailDocumentWriter(new EmailWriterOptions(usePreservedRawSource: true));
        Assert.True(writer.AnalyzeConversion(read.Document).CanWrite);
        Assert.Equal(source, writer.ToBytes(read.Document, EmailFileFormat.Eml, out EmailWriteResult result));
        Assert.Equal(EmailArtifactSourceSelection.PreservedSource, result.SourceSelection);
        Assert.Empty(result.Diagnostics);
        read.Document.Subject = "Changed";
        Assert.Empty(writer.ToBytes(read.Document, EmailFileFormat.Eml, out EmailWriteResult blocked));
        Assert.True(blocked.HasErrors);
        Assert.Contains(blocked.Diagnostics, item => item.Code == "EMAIL_TRANSPORT_SIGNATURE_INVALIDATED");
    }

    [Fact]
    public async Task NestedSignedMessageIsCoveredByPreflightAndAsyncRemoval() {
        var document = new EmailDocument();
        document.Attachments.Add(new EmailAttachment { FileName = "nested.eml", EmbeddedDocument = Message("DKIM-Signature") });
        using var output = new MemoryStream();
        Assert.True((await new EmailDocumentWriter().WriteAsync(document, output)).HasErrors);
        Assert.Empty(output.ToArray());
        var writer = new EmailDocumentWriter(new EmailWriterOptions(OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures));
        EmailWriteResult result = await writer.WriteAsync(document, output);
        Assert.False(result.HasErrors);
        Assert.Contains(result.Diagnostics, item => item.Location == "headers/attachment/0");
        Assert.DoesNotContain("DKIM-Signature:", Encoding.ASCII.GetString(output.ToArray()), StringComparison.OrdinalIgnoreCase);
        Assert.Single(document.Attachments[0].EmbeddedDocument!.Headers);
    }

    private static EmailDocument Message(string header) {
        var document = new EmailDocument { Subject = "Changed" };
        document.Body.Text = "Changed body";
        document.Headers.Add(new EmailHeader(header, "v=1; bh=AAAA; b=BBBB"));
        return document;
    }

    [Theory]
    [InlineData(EmailFileFormat.OutlookMsg)]
    [InlineData(EmailFileFormat.OutlookTemplate)]
    [InlineData(EmailFileFormat.Tnef)]
    public void RemovalPolicyAlsoFiltersOutlookTransportMessageHeaders(EmailFileFormat format) {
        var document = Message("DKIM-Signature");
        document.Headers.Add(new EmailHeader("Content-MD5", "old-digest"));
        document.Headers.Add(new EmailHeader("Received", "from original.example.test"));
        var writer = new EmailDocumentWriter(new EmailWriterOptions(OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures));
        byte[] bytes = writer.ToBytes(document, format, out EmailWriteResult result);
        using EmailReadResult read = new EmailDocumentReader().Read(bytes);
        Assert.False(result.HasErrors);
        Assert.DoesNotContain(read.Document.Headers, header => header.Name == "DKIM-Signature" || header.Name == "Content-MD5");
        Assert.Contains(read.Document.Headers, header => header.Name == "Received");
    }
}
