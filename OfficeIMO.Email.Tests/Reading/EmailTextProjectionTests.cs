using OfficeIMO.Reader;
using OfficeIMO.Reader.Email;

namespace OfficeIMO.Email.Tests;

public sealed class EmailTextProjectionTests {
    [Theory]
    [InlineData("> alpha \r\n> beta", false, ">alpha beta")]
    [InlineData(" >literal \r\n>quoted", false, ">literal \r\n>quoted")]
    [InlineData("> -- \r\n>signature", false, ">-- \r\n>signature")]
    [InlineData("> alpha \r\n> -- \r\n>signature", false, ">alpha \r\n>-- \r\n>signature")]
    [InlineData("> alpha \r\n> beta", true, ">alphabeta")]
    [InlineData(">> nested \r\n>> continued\r\n> other", false, ">>nested continued\r\n>other")]
    public async Task FlowedQuoteDepthAndSignatureBoundariesAgreeAcrossReaders(string wire, bool delsp, string expected) {
        byte[] bytes = Encoding.UTF8.GetBytes("Content-Type: text/plain; charset=utf-8; format=flowed" +
            (delsp ? "; delsp=yes" : "") + "\r\n\r\n" + wire);
        using EmailReadResult buffered = new EmailDocumentReader().Read(bytes);
        Assert.Equal(expected, buffered.Document.Body.Text);
        using var stream = new MemoryStream(bytes);
        using EmailReadResult streaming = await new EmailDocumentReader().ReadStreamingAsync(stream, "flowed.eml");
        Assert.Equal(expected, streaming.Document.Body.Text);
    }

    [Fact]
    public void ReaderBodyAndAttachmentChunksRecombineWithoutBreakingUnicodeScalars() {
        string text = new string('x', 255) + "😀" + "suffix";
        var document = new EmailDocument { Subject = new string('y', 255) + "😀" };
        document.Body.Text = text;
        document.Attachments.Add(new EmailAttachment { FileName = "text.txt", ContentType = "text/plain", Content = Encoding.UTF8.GetBytes(text) });
        var reader = new OfficeDocumentReaderBuilder().AddEmailHandlers().Build();
        using var stream = new MemoryStream(document.ToBytes());
        OfficeDocumentReadResult result = reader.ReadDocument(stream, "unicode.eml", new ReaderOptions { MaxChars = 256 });
        Assert.All(result.Chunks, chunk => new UTF8Encoding(false, true).GetBytes(chunk.Text));
        Assert.Equal(text, string.Concat(result.Chunks.Where(chunk => chunk.Location.SourceBlockKind == "email-body").Select(chunk => chunk.Text)));
        Assert.Equal(text, string.Concat(result.Chunks.Where(chunk => chunk.Id.StartsWith("email:attachment-content:")).Select(chunk => chunk.Text)));
        Assert.All(result.Chunks, chunk => Assert.InRange(chunk.Text.Length, 0, 256));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ReaderUsesDeclaredAttachmentCharsetEvenWithNestedTextHandler(bool registerTextHandler) {
        const string eml = "Subject: Text attachment\r\nContent-Type: multipart/mixed; boundary=b\r\n\r\n--b\r\nContent-Type: text/plain\r\n\r\nBody\r\n--b\r\nContent-Type: text/plain; charset=iso-8859-1\r\nContent-Disposition: attachment; filename=cafe.txt\r\nContent-Transfer-Encoding: base64\r\n\r\nY2Fm6Q==\r\n--b--\r\n";
        var builder = new OfficeDocumentReaderBuilder().AddEmailHandlers();
        if (registerTextHandler) builder.AddHandler(new ReaderHandlerRegistration {
            Id = "text-test", Kind = ReaderInputKind.Text, Extensions = new[] { ".txt" },
            ReadStream = (input, name, options, token) => {
                using var textReader = new StreamReader(input, Encoding.UTF8, true, 4096, leaveOpen: true);
                string text = textReader.ReadToEnd();
                return new[] { new ReaderChunk { Text = text, Markdown = text, Location = new ReaderLocation { SourceBlockKind = "text" } } };
            }
        });
        var reader = builder.Build();
        using var stream = new MemoryStream(Encoding.ASCII.GetBytes(eml));
        var result = reader.ReadDocument(stream, "attachment.eml");
        Assert.Equal("café", Assert.Single(result.Chunks, chunk => chunk.Id.StartsWith("email:attachment-content:")).Text);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code.Contains("CHARSET"));
    }

    [Theory]
    [InlineData("calendar.ics", "BEGIN:VCALENDAR\r\nVERSION:2.0\r\nBEGIN:VEVENT\r\nUID:test\r\nSUMMARY:", "\r\nEND:VEVENT\r\nEND:VCALENDAR\r\n")]
    [InlineData("contact.vcf", "BEGIN:VCARD\r\nVERSION:4.0\r\nFN:", "\r\nEND:VCARD\r\n")]
    public void CalendarAndContactChunksPreserveUnicodeAcrossBoundaries(string name, string prefix, string suffix) {
        string value = new string('x', 400) + "😀" + new string('y', 400);
        byte[] bytes = Encoding.UTF8.GetBytes(prefix + value + suffix);
        var reader = new OfficeDocumentReaderBuilder().AddEmailHandlers().Build();
        using var stream = new MemoryStream(bytes);
        var result = reader.ReadDocument(stream, name, new ReaderOptions { MaxChars = 256 });
        Assert.All(result.Chunks, chunk => new UTF8Encoding(false, true).GetBytes(chunk.Text));
        string text = string.Concat(result.Chunks.Select(chunk => chunk.Text)).Replace("\n ", "");
        Assert.Contains(value, text);
    }

    [Fact]
    public void AttachmentBomOverridesDeclarationWithEvidenceAndExactBudget() {
        var attachment = new EmailAttachment { Content = Encoding.Unicode.GetPreamble().Concat(Encoding.Unicode.GetBytes("café")).ToArray() };
        attachment.ContentTypeParameters["charset"] = "iso-8859-1";
        var result = EmailAttachmentTextReader.Read(attachment, attachment.Content.Length);
        Assert.Equal("café", result.Text);
        Assert.Equal("utf-16", result.Charset);
        Assert.Contains(result.Diagnostics, item => item.Code == "EMAIL_ATTACHMENT_BOM_OVERRIDES_CHARSET");
        Assert.Throws<EmailLimitExceededException>(() => EmailAttachmentTextReader.Read(attachment, attachment.Content.Length - 1));
        Assert.Throws<OperationCanceledException>(() => EmailAttachmentTextReader.Read(attachment, cancellationToken: new CancellationToken(true)));
    }

    [Theory]
    [InlineData("utf-8", "EMAIL_ATTACHMENT_TEXT_INVALID_ENCODING")]
    [InlineData("unknown-charset", "EMAIL_MIME_CHARSET_UNSUPPORTED")]
    [InlineData(null, "EMAIL_MIME_CHARSET_GUESSED")]
    public void AttachmentRecoveryAlwaysReportsEncodingAmbiguity(string? charset, string diagnosticCode) {
        var attachment = new EmailAttachment { Content = new byte[] { 0xe9 } };
        if (charset != null) attachment.ContentTypeParameters["charset"] = charset;
        Assert.Contains(EmailAttachmentTextReader.Read(attachment).Diagnostics, item => item.Code == diagnosticCode);
    }
}
