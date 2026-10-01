using OfficeIMO.Email;
using OfficeIMO.Email.Data;
using OfficeIMO.Email.Store;
using System.Text;

internal static class EmailToolkitSmoke {
    internal static void Run() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-email-toolkit-smoke-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            using EmailReadResult read = new EmailDocumentReader().Read(Encoding.UTF8.GetBytes(
                "From: sender@example.test\r\nTo: me@example.test\r\n" +
                "Message-ID: <source@example.test>\r\nSubject: Toolkit smoke\r\n" +
                "Content-Type: text/html; charset=utf-8\r\n\r\n" +
                "<p>Searchable toolkit body</p><script>blocked()</script>"));
            EmailDocument message = read.Document;
            EmailIndexTextResult index = EmailIndexText.Create(message);
            Require(index.SelectedText.Contains("Searchable toolkit body", StringComparison.Ordinal), "HTML index text");
            Require(!index.SelectedText.Contains("blocked()", StringComparison.Ordinal), "active-content exclusion");

            EmailCompositionResult reply = EmailComposer.ReplyAll(message,
                new EmailAddress("me@example.test"), "Thank you.");
            EmailRecipient[] recipients = reply.Document.Recipients.Where(value => value.Kind == EmailRecipientKind.To).ToArray();
            Require(recipients.Length == 1 && recipients[0].Address.Address == "sender@example.test", "reply-all recipients");
            EmailHtmlCompositionResult htmlReply = EmailHtmlComposer.Reply(message,
                new EmailAddress("me@example.test"), "<authored text>");
            Require(htmlReply.Document.Body.Html?.Contains("&lt;authored text&gt;", StringComparison.Ordinal) == true,
                "encoded HTML draft text");
            Require(htmlReply.Document.Body.Html?.Contains("<script", StringComparison.OrdinalIgnoreCase) != true,
                "resource-free HTML draft");

            EmailShareCopyResult share = EmailShareCopy.Create(message, new EmailShareCopyOptions {
                ReplacementBodyText = index.SelectedText, ReplacementSubject = "Shared smoke"
            });
            Require(share.Document.Recipients.Count == 0 && share.Document.Subject == "Shared smoke", "field-selected share copy");
            Require(message.Subject == "Toolkit smoke", "source preservation");

            string messagePath = Path.Combine(root, "message.eml");
            message.Save(messagePath);
            EmailHtmlDataInspectionReport inspection = EmailHtmlDataInspector.Inspect(messagePath);
            Require(inspection.Metadata.MessageInspected && inspection.HtmlInspectionStatus == "Completed", "data inspection");
            Require(inspection.BlockedElementCount == 1 && !inspection.Metadata.CryptographicallyVerified, "inspection evidence");

            message.Attachments.Add(new EmailAttachment { FileName = "sample.bin", Content = new byte[] { 1, 2, 3 } });
            EmailAttachmentExtractionResult extraction = EmailAttachmentExtractor.Extract(message, Path.Combine(root, "attachments"));
            EmailAttachmentExtractionEntry extracted = extraction.Entries.Single();
            Require(extracted.Sha256?.Length == 64 && extracted.OutputPath != null &&
                File.ReadAllBytes(extracted.OutputPath).SequenceEqual(new byte[] { 1, 2, 3 }), "attachment extraction manifest");

            string mailbox = Path.Combine(root, "mailbox");
            Directory.CreateDirectory(mailbox);
            File.Copy(messagePath, Path.Combine(mailbox, "one.eml"));
            File.Copy(messagePath, Path.Combine(mailbox, "two.eml"));
            using EmailStoreSession store = EmailStoreSession.Open(mailbox);
            EmailArchiveAnalysisReport archive = store.AnalyzeArchive();
            Require(archive.ItemsProjected == 2 && archive.DuplicateCandidateGroupCount == 1, "archive candidates");
            EmailStoreContentSearchReport search = store.SearchContent(new EmailStoreContentQuery(new[] { "Searchable toolkit" }));
            Require(search.Results.Count == 2 && search.Results[0].ResumeAfter.Value.Length > 0, "content search and checkpoint");

            var appointment = new EmailDocument {
                OutlookItemKind = OutlookItemKind.Appointment, Subject = "Package meeting",
                Appointment = new OutlookAppointment {
                    Start = new DateTimeOffset(2026, 10, 1, 9, 0, 0, TimeSpan.Zero),
                    End = new DateTimeOffset(2026, 10, 1, 10, 0, 0, TimeSpan.Zero)
                }
            };
            Require(IcsDocument.Load(EmailPortableContentExport.ToCalendar(appointment).ToBytes())
                .GetComponents("VEVENT").Count() == 1, "calendar export");
            var contact = new OutlookContact { DisplayName = "Package contact" };
            contact.Email1.Address = "contact@example.test";
            var contactDocument = new EmailDocument { OutlookItemKind = OutlookItemKind.Contact, Contact = contact };
            Require(VCardDocument.Load(EmailPortableContentExport.ToVCard(contactDocument).ToBytes()).Cards.Count == 1,
                "contact export");
            EmailContactReviewResult contacts = EmailContactConsolidation.Review(new[] { new EmailContactSource("package:one", contactDocument) });
            Require(EmailContactConsolidation.Consolidate(contacts, 0).Cards[0].GetVCardText("FN") == "Package contact",
                "contact consolidation");
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    private static void Require(bool condition, string operation) {
        if (!condition) throw new InvalidOperationException("Packed email toolkit failed: " + operation);
    }
}
