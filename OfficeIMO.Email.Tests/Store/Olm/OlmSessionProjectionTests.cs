using System.Xml.Linq;

namespace OfficeIMO.Email.Store.Tests.Olm;

public sealed class OlmSessionProjectionTests {
    [Fact]
    public void SelectedReadStreamsContentAndExpiresOutstandingReadersWithSession() {
        byte[] bytes = Enumerable.Range(0, 65536).Select(value => (byte)value).ToArray();
        using var builder = new OlmTestArchiveBuilder();
        byte[] archive = builder.AddText("Local/com.microsoft.__Messages/Inbox/items.xml", Message("First") + "</emails>")
            .Add("Local/attachment", bytes).Build();
        using var source = new MemoryStream(archive, writable: false);
        using var session = EmailStoreSession.Open(source, "archive.olm", leaveOpen: true);
        EmailStoreItemReference reference = Assert.Single(session.EnumerateItems());
        Assert.Equal("First", session.ReadSummary(reference).Subject);
        EmailStoreItem metadata = session.ReadItem(reference, new EmailStoreItemReadOptions(EmailStoreItemReadParts.Metadata));
        Assert.Null(metadata.Document.Body.Text);
        Assert.Empty(metadata.Document.Attachments);
        EmailStoreItem item = session.ReadItem(reference, new EmailStoreItemReadOptions(preferStreamingAttachmentContent: true));
        EmailAttachment attachment = Assert.Single(item.Document.Attachments);
        Assert.Null(attachment.Content);
        Assert.NotNull(attachment.ContentSource);
        using Stream content = attachment.ContentSource!.OpenRead();
        using var output = new MemoryStream();
        content.CopyTo(output);
        Assert.Equal(bytes, output.ToArray());
        session.Dispose();
        Assert.Throws<ObjectDisposedException>(() => content.ReadByte());
        Assert.Throws<ObjectDisposedException>(() => attachment.ContentSource.OpenRead());
        Assert.True(source.CanRead);
    }

    [Fact]
    public void PreservesUnknownStructureAttributesAndCaseCollisionsAndReportsPortableLoss() {
        const string xml = "<emails><email vendor='root'><OPFMessageCopySubject>Żółć 日本語</OPFMessageCopySubject>" +
            "<VendorCopyBody>retain me</VendorCopyBody><vendor flag='a'><value>one</value></vendor>" +
            "<Vendor flag='b'>two</Vendor><repeat>first</repeat><repeat>second</repeat></email></emails>";
        using var builder = new OlmTestArchiveBuilder();
        using var source = new MemoryStream(builder.AddText("Local/Inbox/items.xml", xml).Build());
        using var session = EmailStoreSession.Open(source, "archive.olm", leaveOpen: true);
        EmailDocument document = session.ReadItem(Assert.Single(session.EnumerateItems())).Document;
        Assert.Equal("retain me", document.Properties["Olm:VendorCopyBody"]);
        XElement attributes = XElement.Parse(Assert.IsType<string>(document.Properties["Olm:ItemAttributes"]));
        Assert.Equal("root", attributes.Attribute("vendor")?.Value);
        string[] fragments = Assert.IsType<string[]>(document.Properties["Olm:StructuredProperties"]);
        XElement[] preserved = fragments.Select(XElement.Parse).ToArray();
        Assert.Equal(new[] { "vendor", "Vendor", "repeat", "repeat" }, preserved.Select(element => element.Name.LocalName));
        Assert.Equal(new[] { "first", "second" }, preserved.Where(element => element.Name.LocalName == "repeat").Select(element => element.Value));
        using var destination = new MemoryStream();
        EmailWriteResult conversion = new EmailDocumentWriter(new EmailWriterOptions(conversionLossPolicy: EmailConversionLossPolicy.Warn)).Write(document, destination, EmailFileFormat.Eml);
        Assert.Contains(conversion.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_OLM_METADATA_NOT_REPRESENTED");
        Assert.Equal("Żółć 日本語", new EmailDocumentReader().Read(new MemoryStream(destination.ToArray())).Document.Subject);
    }

    [Fact]
    public void EnforcesPerItemPropertyAndNarrowedDecodedBytesAndRetainedStreamBudget() {
        using var builder = new OlmTestArchiveBuilder();
        byte[] archive = builder.AddText("Local/Inbox/items.xml", Message("First") + "</emails>")
            .Add("Local/attachment", new byte[8]).Build();
        using var limitedSource = new MemoryStream(archive);
        Assert.Equal(nameof(EmailStoreReaderOptions.MaxPropertiesPerItem),
            Assert.Throws<EmailStoreLimitExceededException>(() => EmailStoreSession.Open(limitedSource, "archive.olm",
                new EmailStoreReaderOptions(maxPropertiesPerItem: 2), leaveOpen: true)).LimitName);
        using var source = new MemoryStream(archive);
        using var session = EmailStoreSession.Open(source, "archive.olm", new EmailStoreReaderOptions(maxTotalAttachmentBytes: 8), leaveOpen: true);
        EmailStoreItemReference reference = Assert.Single(session.EnumerateItems());
        Assert.Equal(nameof(EmailStoreReaderOptions.MaxDecodedPropertyBytesPerItem),
            Assert.Throws<EmailStoreLimitExceededException>(() => session.ReadItem(reference,
                new EmailStoreItemReadOptions(maxDecodedPropertyBytes: 1))).LimitName);
        var streaming = new EmailStoreItemReadOptions(preferStreamingAttachmentContent: true);
        EmailAttachment attachment = Assert.Single(session.ReadItem(reference, streaming).Document.Attachments);
        Assert.Equal(nameof(EmailStoreReaderOptions.MaxTotalAttachmentBytes),
            Assert.Throws<EmailStoreLimitExceededException>(() => session.ReadItem(reference, streaming)).LimitName);
        using Stream retained = attachment.ContentSource!.OpenRead();
        Assert.Equal(0, retained.ReadByte());
    }

    [Fact]
    public void RepeatedMalformedAttachmentDiagnosticsStayDeduplicated() {
        using var builder = new OlmTestArchiveBuilder();
        using var source = new MemoryStream(builder.AddText("Local/Inbox/items.xml", Message("First") + "</emails>").Build());
        using var session = EmailStoreSession.Open(source, "archive.olm", leaveOpen: true);
        EmailStoreItemReference reference = Assert.Single(session.EnumerateItems());
        for (int index = 0; index < 100; index++) session.ReadItem(reference);
        Assert.Single(session.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_STORE_OLM_ATTACHMENT_MISSING");
    }

    private static string Message(string subject) => "<emails><email><OPFMessageCopySubject>" + subject +
        "</OPFMessageCopySubject><OPFMessageCopyBody>Body</OPFMessageCopyBody><OPFMessageCopyAttachmentList>" +
        "<messageAttachment OPFAttachmentName='payload.bin' OPFAttachmentURL='Local/attachment' />" +
        "</OPFMessageCopyAttachmentList></email>";
}
