using OfficeIMO.Email.Store;
using System.IO.Compression;

namespace OfficeIMO.Email.Tests;

public sealed class EmailDecodedBodyBudgetTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task MimeAlternativesShareDecodedPropertyBudgetBeforePayloadMaterialization(bool streaming) {
        const string mime = "Content-Type: multipart/alternative; boundary=b\r\n\r\n--b\r\nContent-Type: text/plain\r\n" +
            "Content-Transfer-Encoding: base64\r\n\r\nYWJjZA==\r\n--b\r\nContent-Type: text/html\r\n" +
            "Content-Transfer-Encoding: base64\r\n\r\nZWZnaA==\r\n--b--\r\n";
        var reader = new EmailDocumentReader(new EmailReaderOptions(maxDecodedPropertyBytes: 7));
        using var input = new MemoryStream(Encoding.ASCII.GetBytes(mime));
        EmailLimitExceededException exceeded = streaming
            ? await Assert.ThrowsAsync<EmailLimitExceededException>(() => reader.ReadStreamingAsync(input, "budget.eml"))
            : Assert.Throws<EmailLimitExceededException>(() => reader.Read(input));
        Assert.Equal(nameof(EmailReaderOptions.MaxDecodedPropertyBytes), exceeded.LimitName);
        Assert.Equal(8, exceeded.ActualValue);
        input.Position = 0;
        var exact = new EmailDocumentReader(new EmailReaderOptions(maxDecodedPropertyBytes: 8));
        using var result = streaming ? await exact.ReadStreamingAsync(input, "budget.eml") : exact.Read(input);
        Assert.Equal("abcd", result.Document.Body.Text);
        Assert.Equal("efgh", result.Document.Body.Html);
        Assert.Equal(8, result.ProcessingBudget.DecodedPropertyBytes);
    }

    [Theory]
    [InlineData("eml")]
    [InlineData("directory-emlx")]
    [InlineData("mbox")]
    [InlineData("emlx")]
    [InlineData("olm")]
    public void NarrowerSelectiveReadBudgetCannotBeDroppedByNonPstBackends(string kind) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-selective-budget-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            byte[] message = Encoding.UTF8.GetBytes("Subject: limit\r\n\r\nbody beyond five bytes");
            string path = root;
            if (kind == "eml") File.WriteAllBytes(Path.Combine(root, "message.eml"), message);
            else if (kind == "directory-emlx" || kind == "emlx") {
                string file = Path.Combine(root, "message.emlx");
                File.WriteAllBytes(file, Encoding.ASCII.GetBytes(message.Length + "\n").Concat(message).ToArray());
                if (kind == "emlx") path = file;
            } else if (kind == "mbox") {
                path = Path.Combine(root, "message.mbox");
                File.WriteAllText(path, "From a@example.test Sat Jan 01 00:00:00 2022\n" + Encoding.UTF8.GetString(message));
            } else {
                path = Path.Combine(root, "message.olm");
                using (var archive = new ZipArchive(File.Create(path), ZipArchiveMode.Create)) {
                    using var writer = new StreamWriter(archive.CreateEntry("Local/com.microsoft.__Messages/Inbox/Messages.xml").Open());
                    writer.Write("<emails><email><OPFMessageCopyBody>body beyond five bytes</OPFMessageCopyBody></email></emails>");
                }
            }
            using var session = EmailStoreSession.Open(path);
            var reference = Assert.Single(session.EnumerateItems());
            var exceeded = Assert.Throws<EmailStoreLimitExceededException>(() => session.ReadItem(reference,
                new EmailStoreItemReadOptions(EmailStoreItemReadParts.Bodies, maxDecodedPropertyBytes: 5)));
            Assert.Equal(nameof(EmailStoreReaderOptions.MaxDecodedPropertyBytesPerItem), exceeded.LimitName);
            Assert.True(exceeded.Actual > 5);
            Assert.Contains("body beyond", session.ReadItem(reference).Document.Body.Text);
        } finally { Directory.Delete(root, recursive: true); }
    }
}
