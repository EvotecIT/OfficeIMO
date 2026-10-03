using OfficeIMO;

namespace OfficeIMO.Email.Tests;

public sealed class SynthesizedTaskPayloadTests {
    [Theory]
    [InlineData(EmailConversionLossPolicy.Block)]
    [InlineData(EmailConversionLossPolicy.Warn)]
    [InlineData(EmailConversionLossPolicy.Allow)]
    public async Task SynthesizedTaskPayloadUndergoesNestedLossAnalysis(EmailConversionLossPolicy policy) {
        foreach (string kind in new[] { "Protected", "ArchiveMetadata", "IncompleteProjection" }) {
            EmailDocument leaf;
            string code;
            if (kind == "Protected") {
                leaf = new EmailDocumentReader(new EmailReaderOptions(preserveRawSource: true)).Read(Encoding.ASCII.GetBytes(
                    "Subject: Protected\r\nContent-Type: application/pkcs7-mime; smime-type=enveloped-data\r\n" +
                    "Content-Transfer-Encoding: base64\r\n\r\nAQID\r\n")).Document;
                Assert.True(leaf.Protection.IsProtected);
                code = "EMAIL_PROTECTED_CONTENT_REWRITE";
            } else if (kind == "ArchiveMetadata") {
                leaf = new EmailDocument { Subject = "Source metadata" };
                leaf.Properties["Emlx:Metadata:remote-id"] = "42";
                code = "EMAIL_EMLX_METADATA_NOT_REPRESENTED";
            } else {
                leaf = new EmailDocumentReader().Read(Encoding.ASCII.GetBytes(
                    "Content-Type: text/vcard\r\n\r\nBEGIN:VCARD\r\nVERSION:4.0\r\nFN:Contact\r\nNICKNAME:First,Second\r\nEND:VCARD\r\n")).Document;
                Assert.True(leaf.MimeSemanticProjectionIsIncomplete);
                code = "EMAIL_STORE_SEMANTIC_PROJECTION_INCOMPLETE";
            }
            EmailDocument root = Request();
            root.TaskCommunication!.EmbeddedTask!.Attachments.Add(new EmailAttachment { EmbeddedDocument = leaf, FileName = "leaf.eml" });
            var writer = new EmailDocumentWriter(new EmailWriterOptions(conversionLossPolicy: policy));
            bool blocked = policy == EmailConversionLossPolicy.Block;
            foreach (EmailFileFormat format in new[] { EmailFileFormat.OutlookMsg, EmailFileFormat.OutlookTemplate, EmailFileFormat.Tnef }) {
                Assert.Equal(!blocked, writer.AnalyzeConversion(root, format).CanWrite);
                foreach (bool asynchronous in new[] { false, true }) {
                    using var output = new MemoryStream();
                    output.WriteByte(123);
                    EmailWriteResult result = asynchronous ? await writer.WriteAsync(root, output, format) : writer.Write(root, output, format);
                    Assert.Equal(blocked ? EmailConversionLossDisposition.Blocked : EmailConversionLossDisposition.Accepted, result.LossDisposition);
                    Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == code && diagnostic.Location?.StartsWith("task/embedded/attachment/0/", StringComparison.Ordinal) == true);
                    if (blocked) Assert.Equal(new byte[] { 123 }, output.ToArray());
                }
            }
        }
    }

    [Theory]
    [InlineData(EmailFileFormat.OutlookMsg)]
    [InlineData(EmailFileFormat.Tnef)]
    public void SynthesizedTaskTransportSignaturesBlockBeforeWriting(EmailFileFormat format) {
        EmailDocument root = Request();
        root.TaskCommunication!.EmbeddedTask!.Headers.Add(new EmailHeader("DKIM-Signature", "v=1; a=rsa-sha256; bh=hash; b=signature"));
        var writer = new EmailDocumentWriter(new EmailWriterOptions(conversionLossPolicy: EmailConversionLossPolicy.Allow));
        Assert.False(writer.AnalyzeConversion(root, format).CanWrite);
        using var output = new MemoryStream();
        output.WriteByte(123);
        EmailWriteResult result = writer.Write(root, output, format);
        Assert.True(result.HasErrors);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_TRANSPORT_SIGNATURE_INVALIDATED");
        Assert.Equal(new byte[] { 123 }, output.ToArray());
    }

    [Theory]
    [InlineData(EmailFileFormat.OutlookMsg)]
    [InlineData(EmailFileFormat.Tnef)]
    public async Task SynthesizedTaskAttachmentUsesAsyncStagingAndHonorsOutputBounds(EmailFileFormat format) {
        EmailDocument root = Request();
        var source = new AsyncSource(new byte[] { 7, 8, 9 });
        root.TaskCommunication!.EmbeddedTask!.Attachments.Add(new EmailAttachment { FileName = "payload.bin", ContentSource = source, Length = 3 });
        using var output = new MemoryStream();
        EmailWriteResult result = await new EmailDocumentWriter().WriteAsync(root, output, format);
        Assert.False(result.HasErrors);
        Assert.Equal(1, source.AsyncOpenCount);
        EmailDocument read = new EmailDocumentReader().Read(output.ToArray()).Document;
        Assert.Equal(new byte[] { 7, 8, 9 }, Assert.Single(read.TaskCommunication!.EmbeddedTask!.Attachments).Content);

        var oversized = new AsyncSource(new byte[4096]);
        root.TaskCommunication.EmbeddedTask.Attachments[0].ContentSource = oversized;
        using var destination = new MemoryStream();
        destination.WriteByte(123);
        await Assert.ThrowsAsync<EmailLimitExceededException>(() => new EmailDocumentWriter(new EmailWriterOptions(maxOutputBytes: 1024))
            .WriteAsync(root, destination, format));
        Assert.Equal(0, oversized.AsyncOpenCount);
        Assert.Equal(new byte[] { 123 }, destination.ToArray());
    }

    [Fact]
    public void SynthesizedTaskAnalysisDoesNotAssignIdentityAndCountsItsNestingDepth() {
        EmailDocument root = Request();
        EmailDocument task = root.TaskCommunication!.EmbeddedTask!;
        task.Task!.GlobalId = null;
        var writer = new EmailDocumentWriter();
        Assert.True(writer.AnalyzeConversion(root, EmailFileFormat.OutlookMsg).CanWrite);
        Assert.Null(task.Task.GlobalId);
        task.Attachments.Add(new EmailAttachment { EmbeddedDocument = new EmailDocument { Subject = "Depth two" } });
        var bounded = new EmailDocumentWriter(new EmailWriterOptions(maxNestedMessageDepth: 1));
        foreach (EmailFileFormat format in new[] { EmailFileFormat.OutlookMsg, EmailFileFormat.OutlookTemplate, EmailFileFormat.Tnef }) {
            EmailLimitExceededException error = Assert.Throws<EmailLimitExceededException>(() => bounded.AnalyzeConversion(root, format));
            Assert.Equal(nameof(EmailWriterOptions.MaxNestedMessageDepth), error.LimitName);
        }
    }

    private static EmailDocument Request() => new EmailDocument {
        OutlookItemKind = OutlookItemKind.Task, Subject = "Task request",
        TaskCommunication = OutlookTaskCommunication.Create(OutlookTaskCommunicationKind.Request, new EmailDocument {
            OutlookItemKind = OutlookItemKind.Task, Subject = "Assigned work",
            Task = new OutlookTask { GlobalId = new Guid("759228F7-84D3-49D6-865E-F8655ADFC1DC") }
        })
    };

    private sealed class AsyncSource : IEmailContentSource {
        private readonly byte[] _bytes;
        internal AsyncSource(byte[] bytes) { _bytes = bytes; }
        public long? Length => _bytes.Length;
        internal int AsyncOpenCount { get; private set; }
        public Stream OpenRead() => throw new InvalidOperationException("The async writer must stage this source through OpenReadAsync.");
        public Task<Stream> OpenReadAsync(CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            AsyncOpenCount++;
            return Task.FromResult<Stream>(new MemoryStream(_bytes, writable: false));
        }
    }
}
