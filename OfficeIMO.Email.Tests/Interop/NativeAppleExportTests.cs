using System.Security.Cryptography;
using System.Text.Json;
using OfficeIMO.Email.Store;

namespace OfficeIMO.Email.Tests;

public sealed partial class NativeAppleExportTests {
    [Fact]
    public void MailExportPreservesChildIdentityHtmlAndAttachmentBytesThroughEml() {
        VerifyFixtures("mail/");
        using EmailStoreSession session = EmailStoreSession.Open(Path.Combine(FixtureRoot, "mail"));
        EmailStoreItemReference[] references = session.EnumerateItems().ToArray();
        Assert.Equal(2, references.Length);
        var documents = references.Select(reference => (reference, document: session.ReadItem(reference,
            new EmailStoreItemReadOptions(preferStreamingAttachmentContent: true)).Document)).ToArray();
        var message = Assert.Single(documents, item => item.document.Subject == "OfficeIMO validation — Zażółć 日本語");
        var child = Assert.Single(documents, item => item.document.Subject == "OfficeIMO child mailbox");
        EmailStoreFolderInfo childFolder = Assert.Single(session.Folders, folder => folder.Id == child.reference.FolderId);
        Assert.Equal("mbox", childFolder.Name);
        Assert.Equal("Child", Assert.Single(session.Folders, folder => folder.Id == childFolder.ParentId).Name);
        Assert.Equal("synthetic@example.test", message.document.From!.Address);
        Assert.Contains("Zażółć 日本語", message.document.Body.Html);
        Assert.Equal(new[] { "synthetic-pixel", null }, message.document.Attachments.Select(attachment => attachment.ContentId));
        Assert.Equal(AttachmentHashes, message.document.Attachments.Select(HashAttachment));

        using EmailReadResult reopened = new EmailDocumentReader().Read(new EmailDocumentWriter().ToBytes(message.document));
        Assert.Equal(message.document.Subject, reopened.Document.Subject);
        Assert.Equal(message.document.Body.Html, reopened.Document.Body.Html);
        Assert.Equal(AttachmentHashes, reopened.Document.Attachments.Select(HashAttachment));
    }

    [Fact]
    public void CalendarExportRetainsZonedRecurrenceAlarmAndVendorFieldsWhenEdited() {
        VerifyFixtures("calendar/");
        IcsDocument calendar = IcsDocument.Load(Path.Combine(FixtureRoot, "calendar", "native-calendar.ics"));
        Assert.DoesNotContain(calendar.Validate(), issue => issue.Severity == ContentLineValidationSeverity.Error);
        ContentLineComponent meeting = Assert.Single(calendar.GetComponents("VEVENT"));
        Assert.Equal("OfficeIMO edited — Zażółć 日本語", meeting.GetFirstProperty("SUMMARY")!.Value);
        Assert.Equal("Europe/Warsaw", meeting.GetTemporalValue("DTSTART")!.Value.TimeZoneId);
        Assert.Equal("20261018T090000", meeting.GetFirstProperty("DTSTART")!.Value);
        Assert.Equal("20261018T100000", meeting.GetFirstProperty("DTEND")!.Value);
        Assert.Equal("FREQ=WEEKLY;COUNT=4", meeting.GetFirstProperty("RRULE")!.Value);
        Assert.Equal("custom-calendar-property", meeting.GetFirstProperty("X-OFFICEIMO-VALIDATION")!.Value);
        ContentLineComponent alarm = Assert.Single(meeting.GetComponents("VALARM"));
        Assert.Equal("DISPLAY", alarm.GetFirstProperty("ACTION")!.Value);
        Assert.Equal("-PT10M", alarm.GetFirstProperty("TRIGGER")!.Value);
        string alarmIdentity = alarm.GetFirstProperty("X-WR-ALARMUID")!.Value;
        string zone = SerializeZone(Assert.Single(calendar.GetComponents("VTIMEZONE")));
        meeting.SetProperty("SUMMARY", "OfficeIMO managed edit — Zażółć 日本語");

        IcsDocument reopened = IcsDocument.Parse(calendar.Serialize());
        Assert.DoesNotContain(reopened.Validate(), issue => issue.Severity == ContentLineValidationSeverity.Error);
        Assert.Equal(zone, SerializeZone(Assert.Single(reopened.GetComponents("VTIMEZONE"))));
        ContentLineComponent updated = Assert.Single(reopened.GetComponents("VEVENT"));
        Assert.Equal("OfficeIMO managed edit — Zażółć 日本語", updated.GetFirstProperty("SUMMARY")!.Value);
        Assert.Equal("FREQ=WEEKLY;COUNT=4", updated.GetFirstProperty("RRULE")!.Value);
        Assert.Equal("Europe/Warsaw", updated.GetTemporalValue("DTSTART")!.Value.TimeZoneId);
        Assert.Equal(alarmIdentity, Assert.Single(updated.GetComponents("VALARM")).GetFirstProperty("X-WR-ALARMUID")!.Value);
    }

    [Fact]
    public void ContactExportRetainsUnicodeFoldedPhotoAndRepeatedTypeParametersWhenEdited() {
        VerifyFixtures("contacts/");
        VCardDocument document = VCardDocument.Load(Path.Combine(FixtureRoot, "contacts", "native-contact-edited.vcf"));
        Assert.DoesNotContain(document.Validate(), issue => issue.Severity == ContentLineValidationSeverity.Error);
        ContentLineComponent card = Assert.Single(document.Cards);
        Assert.Equal("OfficeIMO Zażółć 日本語 Validation", card.GetVCardText("FN"));
        Assert.Equal("Validation;OfficeIMO Zażółć 日本語;;;", card.GetFirstProperty("N")!.Value);
        Assert.Equal(ContactPhotoHash, Hash(Convert.FromBase64String(card.GetFirstProperty("PHOTO")!.Value)));
        ContentLineProperty email = card.GetFirstProperty("EMAIL")!;
        Assert.Equal("officeimo-validation@example.test", email.Value);
        Assert.Equal(new[] { "INTERNET", "WORK", "pref" }, email.Parameters
            .Where(parameter => parameter.Name.Equals("TYPE", StringComparison.OrdinalIgnoreCase)).SelectMany(parameter => parameter.Values));
        card.SetVCardText("FN", "OfficeIMO managed edit — Zażółć 日本語");

        ContentLineComponent reopened = Assert.Single(VCardDocument.Parse(document.Serialize()).Cards);
        Assert.Equal("OfficeIMO managed edit — Zażółć 日本語", reopened.GetVCardText("FN"));
        Assert.Equal(ContactPhotoHash, Hash(Convert.FromBase64String(reopened.GetFirstProperty("PHOTO")!.Value)));
        Assert.Equal(email.Parameters.SelectMany(parameter => parameter.Values),
            reopened.GetFirstProperty("EMAIL")!.Parameters.SelectMany(parameter => parameter.Values));
    }

    private static string FixtureRoot => Path.Combine(EmailTestRepository.FindRoot(), "OfficeIMO.Email.Tests", "Corpora", "native-apple-20261003");
    private const string ContactPhotoHash = "81a3243cc7e9270f8d3c197fe7573ac999ac4de9a24e87ce32f7ec2f13b020f6";

    private static string SerializeZone(ContentLineComponent zone) {
        var document = new IcsDocument();
        document.Calendars[0].Components.Add(zone);
        return document.Serialize();
    }
    private static readonly string[] AttachmentHashes = {
        "a31bc9e0d54f843f4a3c088c88f7542e3ca553a0d9f2ec5956bb1e3ae23621d5",
        "3d5b3a5d3cbfe3b002f15d1b73331b9ab4f528831c9631cc09e41e032d8ceb2f"
    };

    private static void VerifyFixtures(string prefix) {
        using JsonDocument manifest = JsonDocument.Parse(File.ReadAllText(Path.Combine(FixtureRoot, "manifest.json")));
        foreach (JsonElement fixture in manifest.RootElement.GetProperty("files").EnumerateArray()) {
            string path = fixture.GetProperty("path").GetString()!;
            if (path.StartsWith(prefix, StringComparison.Ordinal))
                Assert.Equal(fixture.GetProperty("sha256").GetString(), Hash(File.ReadAllBytes(Path.Combine(FixtureRoot, path))));
        }
    }

    private static string HashAttachment(EmailAttachment attachment) {
        using Stream input = attachment.OpenContentStream();
        using SHA256 hash = SHA256.Create();
        return BitConverter.ToString(hash.ComputeHash(input)).Replace("-", string.Empty).ToLowerInvariant();
    }

    private static string Hash(byte[] bytes) {
        using SHA256 hash = SHA256.Create();
        return BitConverter.ToString(hash.ComputeHash(bytes)).Replace("-", string.Empty).ToLowerInvariant();
    }
}
