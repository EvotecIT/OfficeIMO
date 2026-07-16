using OfficeIMO.Email;

namespace OfficeIMO.Email.Store.Tests.Olm;

public sealed class OlmStoreReaderTests {
    [Fact]
    public void ReadsMessageHierarchyRecipientsBodiesCategoriesAndAttachment() {
        const string attachmentPath = "Local/com.microsoft.__Messages/Account/Inbox/com.microsoft.__Attachments/item_0000";
        string xml = "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
                     "<emails><email>" +
                     "<OPFMessageCopySubject>Quarterly update</OPFMessageCopySubject>" +
                     "<OPFMessageCopyMessageID>olm-message@example.test</OPFMessageCopyMessageID>" +
                     "<OPFMessageCopySentTime>2026-07-14T08:30:00Z</OPFMessageCopySentTime>" +
                     "<OPFMessageCopyReceivedTime>2026-07-14T08:31:00Z</OPFMessageCopyReceivedTime>" +
                     "<OPFMessageCopyBody>Plain body</OPFMessageCopyBody>" +
                     "<OPFMessageCopyHTMLBody>&lt;p&gt;HTML body&lt;/p&gt;</OPFMessageCopyHTMLBody>" +
                     "<OPFMessageCopyFromAddresses><emailAddress OPFContactEmailAddressAddress=\"sender@example.test\" OPFContactEmailAddressName=\"Sender\" /></OPFMessageCopyFromAddresses>" +
                     "<OPFMessageCopyToAddresses><emailAddress OPFContactEmailAddressAddress=\"to@example.test\" OPFContactEmailAddressName=\"To\" /></OPFMessageCopyToAddresses>" +
                     "<OPFMessageCopyCCAddresses><emailAddress OPFContactEmailAddressAddress=\"cc@example.test\" /></OPFMessageCopyCCAddresses>" +
                     "<OPFMessageCopyCategoryList><category OPFCategoryCopyName=\"Project Blue\" /></OPFMessageCopyCategoryList>" +
                     "<OPFMessageCopyAttachmentList><messageAttachment OPFAttachmentName=\"inline.png\" OPFAttachmentURL=\"" + attachmentPath + "\" OPFAttachmentContentType=\"image/png\" OPFAttachmentContentID=\"&lt;logo@example.test&gt;\" /></OPFMessageCopyAttachmentList>" +
                     "<OPFMessageGetIsRead>1E0</OPFMessageGetIsRead>" +
                     "<OPFMessageGetPriority>2</OPFMessageGetPriority>" +
                     "</email></emails>";
        byte[] attachmentBytes = { 1, 2, 3, 4 };
        byte[] archive;
        using (var builder = new OlmTestArchiveBuilder()) {
            archive = builder
                .AddText("Local/com.microsoft.__Messages/Account/Inbox/message_00000.xml", xml)
                .Add(attachmentPath, attachmentBytes)
                .Build();
        }

        EmailStoreReadResult result = Read(archive);

        Assert.Equal(EmailStoreFormat.Olm, result.Store.Format);
        Assert.Equal("archive", result.Store.DisplayName);
        Assert.Equal(3, result.Store.Folders.Count);
        Assert.DoesNotContain(result.Store.Folders, folder => folder.Name == "com.microsoft.__Messages");
        EmailStoreFolder inbox = Assert.Single(result.Store.Folders, folder => folder.Name == "Inbox");
        EmailDocument document = Assert.Single(inbox.Items).Document;
        Assert.Equal("Quarterly update", document.Subject);
        Assert.Equal("sender@example.test", document.From?.Address);
        Assert.Equal("Plain body", document.Body.Text);
        Assert.Equal("<p>HTML body</p>", document.Body.Html);
        Assert.Equal(2, document.Recipients.Count);
        Assert.Contains(document.Recipients, recipient => recipient.Kind == EmailRecipientKind.To && recipient.Address.Address == "to@example.test");
        Assert.Contains(document.Recipients, recipient => recipient.Kind == EmailRecipientKind.Cc && recipient.Address.Address == "cc@example.test");
        Assert.True(document.MessageMetadata.IsRead);
        Assert.Equal(EmailMessagePriority.Urgent, document.MessageMetadata.Priority);
        Assert.Contains("Project Blue", document.MessageMetadata.Categories);
        EmailAttachment attachment = Assert.Single(document.Attachments);
        Assert.Equal("inline.png", attachment.FileName);
        Assert.Equal("logo@example.test", attachment.ContentId);
        Assert.True(attachment.IsInline);
        Assert.Equal(attachmentBytes, attachment.Content);
        Assert.Empty(result.Diagnostics);
    }

    [Fact]
    public void ProjectsPortableOutlookItems() {
        const string contacts = "<contacts><contact><OPFContactCopyDisplayName>Ada Lovelace</OPFContactCopyDisplayName><OPFContactCopyFirstName>Ada</OPFContactCopyFirstName><OPFContactCopyLastName>Lovelace</OPFContactCopyLastName><OPFContactCopyBusinessCompany>Analytical Engines</OPFContactCopyBusinessCompany><OPFContactCopyEmailAddressList><contactEmailAddress OPFContactEmailAddressAddress=\"ada@example.test\" OPFContactEmailAddressType=\"0\" /></OPFContactCopyEmailAddressList></contact></contacts>";
        const string appointments = "<appointments><appointment><OPFCalendarEventCopySummary>Design review</OPFCalendarEventCopySummary><OPFCalendarEventCopyUUID>event-id</OPFCalendarEventCopyUUID><OPFCalendarEventCopyStartTime>2026-07-20T10:00:00Z</OPFCalendarEventCopyStartTime><OPFCalendarEventCopyEndTime>2026-07-20T10:45:00Z</OPFCalendarEventCopyEndTime><OPFCalendarEventCopyLocation>Room 4</OPFCalendarEventCopyLocation><OPFCalendarEventGetIsAllDayEvent>0</OPFCalendarEventGetIsAllDayEvent><OPFCalendarEventCopyOrganizer>organizer@example.test</OPFCalendarEventCopyOrganizer><OPFCalendarEventCopyAttendeeList><appointmentAttendee OPFCalendarAttendeeAddress=\"required@example.test\" OPFCalendarAttendeeName=\"Required\" OPFCalendarAttendeeType=\"1\" /><appointmentAttendee OPFCalendarAttendeeAddress=\"optional@example.test\" OPFCalendarAttendeeName=\"Optional\" OPFCalendarAttendeeType=\"2\" /></OPFCalendarEventCopyAttendeeList></appointment></appointments>";
        const string tasks = "<tasks><task><OPFTaskCopyName>Ship release</OPFTaskCopyName><OPFTaskCopyNotePlain>Verify packages</OPFTaskCopyNotePlain><OPFTaskCopyStartDateTime>2026-07-20T08:00:00Z</OPFTaskCopyStartDateTime><OPFTaskCopyDueDateTime>2026-07-21T17:00:00Z</OPFTaskCopyDueDateTime></task></tasks>";
        const string notes = "<notes><note><OPFNoteCopyTitle>Remember</OPFNoteCopyTitle><OPFNoteCopyText>Use the shared owner.</OPFNoteCopyText><OPFNoteCopyCreationDate>2026-07-19T07:00:00Z</OPFNoteCopyCreationDate></note></notes>";
        byte[] archive;
        using (var builder = new OlmTestArchiveBuilder()) {
            archive = builder
                .AddText("Local/Address Book/Contacts.xml", contacts)
                .AddText("Local/Calendar/Calendar.xml", appointments)
                .AddText("Local/Tasks/Tasks.xml", tasks)
                .AddText("Local/Notes/Notes.xml", notes)
                .Build();
        }

        EmailStoreReadResult result = Read(archive);
        EmailDocument[] items = result.Store.Folders.SelectMany(folder => folder.Items)
            .Select(message => message.Document).ToArray();

        Assert.Equal(4, items.Length);
        EmailDocument contact = Assert.Single(items, item => item.OutlookItemKind == OutlookItemKind.Contact);
        Assert.Equal("Ada", contact.Contact?.GivenName);
        Assert.Equal("ada@example.test", contact.Contact?.Email1.Address);
        EmailDocument appointment = Assert.Single(items, item => item.OutlookItemKind == OutlookItemKind.Appointment);
        Assert.Equal(45, appointment.Appointment?.DurationMinutes);
        Assert.Equal("Room 4", appointment.Appointment?.Location);
        Assert.Equal("organizer@example.test", appointment.From?.Address);
        Assert.Equal(2, appointment.Recipients.Count);
        Assert.Equal("Required", appointment.Appointment?.RequiredAttendees);
        Assert.Equal("Optional", appointment.Appointment?.OptionalAttendees);
        EmailDocument task = Assert.Single(items, item => item.OutlookItemKind == OutlookItemKind.Task);
        Assert.Equal("Verify packages", task.Body.Text);
        Assert.NotNull(task.Task?.Due);
        EmailDocument note = Assert.Single(items, item => item.OutlookItemKind == OutlookItemKind.Note);
        Assert.Equal("Use the shared owner.", note.Body.Text);
        Assert.NotNull(note.Note);
    }

    [Fact]
    public void DoesNotRetainAttachmentContentWhenDisabled() {
        const string attachmentPath = "Local/com.microsoft.__Messages/Inbox/com.microsoft.__Attachments/a";
        const string xml = "<emails><email><OPFMessageCopyAttachmentList><messageAttachment OPFAttachmentName=\"a.bin\" OPFAttachmentURL=\"" + attachmentPath + "\" /></OPFMessageCopyAttachmentList></email></emails>";
        byte[] archive;
        using (var builder = new OlmTestArchiveBuilder()) {
            archive = builder.AddText("Local/com.microsoft.__Messages/Inbox/message_00000.xml", xml)
                .Add(attachmentPath, new byte[] { 8, 9, 10 }).Build();
        }
        var options = new EmailStoreReaderOptions(retainAttachmentContent: false);

        EmailStoreReadResult result = Read(archive, options);
        EmailAttachment attachment = Assert.Single(result.Store.Folders.SelectMany(folder => folder.Items)
            .SelectMany(message => message.Document.Attachments));

        Assert.Equal(3, attachment.Length);
        Assert.Null(attachment.Content);
    }

    [Fact]
    public void RejectsUnsafeAttachmentPathWithoutOpeningIt() {
        const string xml = "<emails><email><OPFMessageCopyAttachmentList><messageAttachment OPFAttachmentName=\"secret.txt\" OPFAttachmentURL=\"../secret.txt\" /></OPFMessageCopyAttachmentList></email></emails>";
        byte[] archive;
        using (var builder = new OlmTestArchiveBuilder()) {
            archive = builder.AddText("Local/com.microsoft.__Messages/Inbox/message_00000.xml", xml).Build();
        }

        EmailStoreReadResult result = Read(archive);

        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_STORE_OLM_ATTACHMENT_PATH_UNSAFE");
        Assert.Null(Assert.Single(result.Store.Folders.SelectMany(folder => folder.Items)
            .SelectMany(message => message.Document.Attachments)).Content);
    }

    [Fact]
    public void ReportsMalformedXmlAndContinuesWithValidItems() {
        byte[] archive;
        using (var builder = new OlmTestArchiveBuilder()) {
            archive = builder
                .AddText("Local/com.microsoft.__Messages/Bad/message_00000.xml", "<emails><email>")
                .AddText("Local/com.microsoft.__Messages/Good/message_00000.xml", "<emails><email><OPFMessageCopySubject>Good</OPFMessageCopySubject></email></emails>")
                .Build();
        }

        EmailStoreReadResult result = Read(archive);

        Assert.Equal(1, result.Store.ItemCount);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_STORE_OLM_XML_INVALID" && diagnostic.Severity == EmailStoreDiagnosticSeverity.Error);
    }

    [Fact]
    public void EnforcesArchiveEntryLimitBeforeProjection() {
        byte[] archive;
        using (var builder = new OlmTestArchiveBuilder()) {
            archive = builder.AddText("one.xml", "<notes />").AddText("two.xml", "<notes />").Build();
        }
        var options = new EmailStoreReaderOptions(maxArchiveEntries: 1);

        EmailStoreLimitExceededException exception = Assert.Throws<EmailStoreLimitExceededException>(() => Read(archive, options));

        Assert.Equal(nameof(EmailStoreReaderOptions.MaxArchiveEntries), exception.LimitName);
    }

    [Fact]
    public void PreservesEmptyMailFoldersButHidesAttachmentDirectories() {
        byte[] archive;
        using (var builder = new OlmTestArchiveBuilder()) {
            archive = builder
                .AddDirectory("Local/com.microsoft.__Messages/Account/Empty: Folder")
                .AddDirectory("Local/com.microsoft.__Messages/Account/Empty: Folder/com.microsoft.__Attachments")
                .Build();
        }

        EmailStoreReadResult result = Read(archive);

        EmailStoreFolder empty = Assert.Single(result.Store.Folders, folder => folder.Name == "Empty: Folder");
        Assert.Empty(empty.Items);
        Assert.DoesNotContain(result.Store.Folders, folder => folder.Name == "com.microsoft.__Attachments");
        Assert.Empty(result.Diagnostics);
    }

    [Fact]
    public void EnforcesDecodedArchiveAndXmlLimits() {
        byte[] archive;
        using (var builder = new OlmTestArchiveBuilder()) {
            archive = builder.AddText("Local/Notes/Notes.xml",
                "<notes><note><OPFNoteCopyText>Longer than the configured bound</OPFNoteCopyText></note></notes>")
                .Build();
        }

        var decodedOptions = new EmailStoreReaderOptions(maxArchiveDecodedBytes: 16);
        Assert.Equal(nameof(EmailStoreReaderOptions.MaxArchiveDecodedBytes),
            Assert.Throws<EmailStoreLimitExceededException>(() => Read(archive, decodedOptions)).LimitName);

        var xmlOptions = new EmailStoreReaderOptions(maxXmlCharactersPerItem: 24);
        EmailStoreReadResult result = Read(archive, xmlOptions);
        Assert.Equal(0, result.Store.ItemCount);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_STORE_OLM_XML_INVALID");
    }

    private static EmailStoreReadResult Read(byte[] archive, EmailStoreReaderOptions? options = null) {
        using (var stream = new MemoryStream(archive, writable: false)) {
            return new EmailStoreReader(options).Read(stream, "archive.olm");
        }
    }
}
