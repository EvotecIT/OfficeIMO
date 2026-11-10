using OfficeIMO.Email;
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class MimeTransferCodecTests {
    [Theory]
    [InlineData("a b=\r\nc\n", "a=20b=3D\r\nc=0A\r\n")]
    [InlineData("line\r\n", "line\r\n")]
    [InlineData("", "\r\n")]
    public void PreservedQuotedPrintablePartsKeepExactHexAndLineEndingBytes(string text, string expected) {
        const string source = "Content-Type: application/octet-stream\r\n" +
            "Content-Disposition: attachment; filename=qp.bin\r\n" +
            "Content-Transfer-Encoding: quoted-printable\r\n\r\noriginal";
        using var parsed = new EmailDocumentReader().ReadDocument(Encoding.ASCII.GetBytes(source));
        EmailAttachment attachment = Assert.Single(parsed.Document.Attachments);
        // MHTML rewrites use the same retained-part contract in the Email writer.
        attachment.PreserveMimeHeadersOnWrite = true;
        attachment.Content = Encoding.ASCII.GetBytes(text);
        attachment.Length = attachment.Content.Length;
        string output = Encoding.ASCII.GetString(new EmailDocumentWriter().ToBytes(parsed.Document));
        int header = output.IndexOf("Content-Transfer-Encoding: quoted-printable", StringComparison.Ordinal);
        Assert.True(header >= 0);
        int start = output.IndexOf("\r\n\r\n", header, StringComparison.Ordinal) + 4;
        int boundary = output.IndexOf("\r\n--", start, StringComparison.Ordinal);
        Assert.True(boundary >= start);
        Assert.Equal(expected, output.Substring(start, boundary - start + 2));
    }

    [Fact]
    public void StreamedQuotedPrintableKeepsEscapesAcrossBuffersAndExactAttachmentLimit() {
        string payload = new string('a', 65534) + "=3D=\r\n" + new string('b', 65535) + "=X1=";
        byte[] expected = Encoding.ASCII.GetBytes(new string('a', 65534) + "=" + new string('b', 65535) + "=X1=");
        byte[] source = Encoding.ASCII.GetBytes("Content-Type: application/octet-stream\r\n" +
            "Content-Disposition: attachment; filename=qp.bin\r\nContent-Transfer-Encoding: quoted-printable\r\n\r\n" + payload);
        var reader = new EmailDocumentReader(new EmailReaderOptions(maxAttachmentBytes: expected.Length));
        using var result = reader.ReadDocument(new MemoryStream(source), EmailReadContentStorage.TemporaryFiles, "qp.eml");
        using var content = Assert.Single(result.Document.Attachments).OpenContentStream();
        using var decoded = new MemoryStream();
        content.CopyTo(decoded);
        Assert.Equal(expected, decoded.ToArray());
        Assert.Equal(expected.Length, result.ProcessingBudget.AttachmentBytes);
        Assert.Single(result.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_MIME_QUOTED_PRINTABLE_INVALID");
        var limited = new EmailDocumentReader(new EmailReaderOptions(maxAttachmentBytes: expected.Length - 1));
        var exceeded = Assert.Throws<EmailLimitExceededException>(() =>
            limited.ReadDocument(new MemoryStream(source), EmailReadContentStorage.TemporaryFiles, "qp.eml"));
        Assert.Equal(nameof(EmailReaderOptions.MaxAttachmentBytes), exceeded.LimitName);
        Assert.Equal(expected.Length, exceeded.ActualValue);
    }

    [Theory]
    [InlineData("literal=20text=3Dtail", "literal text=tail", 0)]
    [InlineData("first=\r\nsecond=\nthird", "firstsecondthird", 0)]
    [InlineData("=00=ff=7F", "\0\u00ff\u007f", 0)]
    [InlineData("=X1=\rX=\r=", "=X1=\rX=\r=", 4)]
    [InlineData("truncated=0", "truncated=0", 1)]
    public void QuotedPrintablePreservesLiteralRunsEscapesAndEachMalformedDiagnostic(
        string encoded, string expected, int invalidEscapes) {
        // Shift escapes around vector-search boundaries and decode only the supplied MIME range.
        foreach (int prefix in new[] { 0, 15, 16, 31, 32, 65535 }) {
            byte[] payload = (new string('a', prefix) + encoded).Select(character => (byte)character).ToArray();
            byte[] input = new byte[payload.Length + 2];
            input[0] = input[input.Length - 1] = (byte)'=';
            payload.CopyTo(input, 1);
            long length = MimeTextCodec.GetDecodedLength(input, 1, payload.Length, "quoted-printable");
            var diagnostics = new List<EmailDiagnostic>();
            byte[] decoded = MimeTextCodec.DecodeTransfer(input, 1, payload.Length, length,
                "quoted-printable", diagnostics, "part");
            Assert.Equal((new string('a', prefix) + expected).Select(character => (byte)character), decoded);
            Assert.Equal(decoded.Length, length);
            Assert.Equal(invalidEscapes, diagnostics.Count);
            Assert.All(diagnostics, diagnostic => {
                Assert.Equal("EMAIL_MIME_QUOTED_PRINTABLE_INVALID", diagnostic.Code);
                Assert.Equal("part", diagnostic.Location);
            });
        }
    }
}
