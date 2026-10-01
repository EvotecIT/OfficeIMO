using OfficeIMO.Email;
using OfficeIMO.Email.AddressBook;
using OfficeIMO.Email.AddressBook.Tests;
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class EmailContactConsolidationTests {
    [Fact]
    public void PreservesConflictsProvenanceAndRequiresSelectionWithoutChangingSources() {
        EmailDocument first = EmailPortableContentExportTests.Contact("Ada", "ada@example.com");
        EmailDocument second = EmailPortableContentExportTests.Contact("Ada Lovelace", "ADA@example.com");
        first.Contact!.Phones.Mobile = "111"; second.Contact!.Phones.Mobile = "222";
        var review = EmailContactConsolidation.Review(new[] { new EmailContactSource("archive:a", first), new EmailContactSource("book:b", second) });
        EmailContactReviewGroup group = Assert.Single(review.Groups);
        Assert.Equal(new[] { "archive:a", "book:b" }, group.SourceIds);
        EmailContactReviewField phone = Assert.Single(group.Fields, field => field.Choices[0].Name == "TEL");
        Assert.True(phone.HasConflict);
        Assert.Equal(new[] { "111", "222" }, phone.Choices.Select(choice => choice.Value));
        Assert.Equal("book:b", Assert.Single(phone.Choices[1].SourceIds));
        Assert.Contains("TYPE=CELL,VOICE", phone.Choices[1].Parameters);
        Assert.Throws<InvalidOperationException>(() => EmailContactConsolidation.Consolidate(review, 0));
        var choices = group.Fields.Where(field => field.HasConflict).ToDictionary(field => field.Key, _ => 1);
        VCardDocument selected = EmailContactConsolidation.Consolidate(review, 0, choices);
        Assert.Equal("Ada Lovelace", selected.Cards[0].GetVCardText("FN"));
        Assert.Equal("222", selected.Cards[0].GetFirstProperty("TEL")!.Value);
        Assert.Equal("111", first.Contact.Phones.Mobile);
        second.Contact.DisplayName = "changed after review";
        Assert.Equal("Ada Lovelace", EmailContactConsolidation.Consolidate(review, 0, choices).Cards[0].GetVCardText("FN"));
        Assert.Throws<ArgumentException>(() => EmailContactConsolidation.Consolidate(review, 0, new Dictionary<string, int> { ["wrong"] = 0 }));
    }

    [Fact]
    public void DoesNotGroupEqualNamesAndRejectsAmbiguousOrIncompleteDirectoryMatches() {
        var first = EmailPortableContentExportTests.Contact("Ada", "ada@example.test");
        var alias = EmailPortableContentExportTests.Contact("Ada alias", "alias-ada@example.test");
        var sameName = EmailPortableContentExportTests.Contact("Ada", "unrelated@example.test");
        EmailContactSource[] sources = { new EmailContactSource("one", first), new EmailContactSource("alias", alias), new EmailContactSource("unrelated", sameName) };
        Assert.Equal(3, EmailContactConsolidation.Review(sources).Groups.Count);
        using var stream = new MemoryStream(new OabV4Fixture().Build());
        using var session = OfflineAddressBookSession.Open(stream, "contacts.oab");
        Assert.Equal(2, EmailContactConsolidation.Review(sources, directory: session.BuildIdentityIndex()).Groups.Count);
        var incomplete = EmailContactConsolidation.Review(sources, directory: session.BuildIdentityIndex(new OfflineAddressBookIdentityIndexOptions(maxEntries: 1)));
        Assert.Equal(3, incomplete.Groups.Count);
        Assert.Contains(incomplete.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_CONTACT_IDENTITY_UNRESOLVED");
        using var duplicateStream = new MemoryStream(new OabV4Fixture().AddPerson("Duplicate", "ada@example.test", "dup", "Ada", "Duplicate", "Research").Build());
        using var duplicateSession = OfflineAddressBookSession.Open(duplicateStream, "duplicate.oab");
        var ambiguous = EmailContactConsolidation.Review(sources, directory: duplicateSession.BuildIdentityIndex());
        Assert.Equal(3, ambiguous.Groups.Count);
        Assert.Contains(ambiguous.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_CONTACT_IDENTITY_UNRESOLVED");
    }

    [Fact]
    public void UsesCanonicalOabMappingAndEnforcesInputOutputAndPropertyBounds() {
        using var stream = new MemoryStream(new OabV4Fixture().Build());
        using var session = OfflineAddressBookSession.Open(stream, "contacts.oab");
        OfflineAddressBookEntry entry = session.EnumerateEntries().First();
        EmailContactSource source = EmailContactSource.FromAddressBook(entry, "oab:person");
        Assert.Equal(entry.DisplayName, source.Document.Contact!.DisplayName);
        var review = EmailContactConsolidation.Review(new[] { source });
        Assert.Single(EmailContactConsolidation.Consolidate(review, 0).Cards);
        Assert.Throws<ArgumentException>(() => EmailContactConsolidation.Review(new[] { source, source }));
        var other = new EmailContactSource("other", EmailPortableContentExportTests.Contact("Other", "other@example.com"));
        Assert.Throws<EmailLimitExceededException>(() => EmailContactConsolidation.Review(new[] { source, other }, new EmailContactReviewOptions(maxContacts: 1)));
        Assert.Throws<EmailLimitExceededException>(() => EmailContactConsolidation.Review(new[] { source }, new EmailContactReviewOptions(maxPropertiesPerCard: 1)));
        Assert.Throws<EmailLimitExceededException>(() => EmailContactConsolidation.Review(new[] { source }, new EmailContactReviewOptions(maxTotalBytes: 10, maxContactBytes: 10)));
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => EmailContactConsolidation.Review(new[] { source }, cancellationToken: cancellation.Token));
    }

    [Fact]
    public void CollectionExportsPreserveSeparateCalendarsAndCardsWithAggregateLimits() {
        string Calendar(string uid) => "MIME-Version: 1.0\r\nContent-Type: text/calendar\r\n\r\nBEGIN:VCALENDAR\r\nVERSION:2.0\r\nPRODID:test\r\nBEGIN:VTIMEZONE\r\nTZID:Local\r\nX-SOURCE:" + uid + "\r\nEND:VTIMEZONE\r\nBEGIN:VEVENT\r\nUID:" + uid + "\r\nDTSTART:20261001T090000Z\r\nEND:VEVENT\r\nEND:VCALENDAR\r\n";
        using var a = new EmailDocumentReader().Read(Encoding.UTF8.GetBytes(Calendar("a")));
        using var b = new EmailDocumentReader().Read(Encoding.UTF8.GetBytes(Calendar("b")));
        var collection = EmailPortableContentCollection.ToCalendars(new[] { a.Document, b.Document });
        Assert.True(collection.RetainedSemanticSource);
        Assert.Equal(2, IcsDocument.Load(collection.ToBytes()).Calendars.Count);
        Assert.Equal(new[] { "a", "b" }, collection.Document.GetComponents("VTIMEZONE").Select(zone => zone.GetFirstProperty("X-SOURCE")!.Value));
        long oneBytes = EmailPortableContentExport.ToCalendar(a.Document).ToBytes().LongLength;
        Assert.Throws<EmailLimitExceededException>(() => EmailPortableContentCollection.ToCalendars(new[] { a.Document, b.Document }, new EmailPortableContentExportOptions(maxOutputBytes: oneBytes)));
        Assert.Throws<EmailLimitExceededException>(() => EmailPortableContentCollection.ToCalendars(new[] { a.Document, b.Document }, maxItems: 1));
        Assert.Equal(2, EmailPortableContentCollection.ToVCards(new[] { EmailPortableContentExportTests.Contact("A", "a@example.com"), EmailPortableContentExportTests.Contact("B", "b@example.com") }).Document.Cards.Count);
    }

    [Fact]
    public void RepeatedValuesOnOneContactRemainASetRatherThanAConflict() {
        var contact = EmailPortableContentExportTests.Contact("Ada", "ada@example.com");
        contact.Contact!.Phones.Business = "111"; contact.Contact.Phones.Business2 = "222";
        var review = EmailContactConsolidation.Review(new[] { new EmailContactSource("one", contact) });
        var field = Assert.Single(review.Groups[0].Fields, candidate => candidate.Choices[0].Name == "TEL");
        Assert.False(field.HasConflict);
        Assert.Equal(new[] { "111", "222" }, Assert.Single(field.Choices).Values);
        Assert.Equal(new[] { "111", "222" }, EmailContactConsolidation.Consolidate(review, 0).Cards[0].GetProperties("TEL").Select(property => property.Value));
    }

    [Fact]
    public void MatchesDecodedEmailSemanticsRatherThanQuotedPrintableAddressSyntax() {
        string eml = "MIME-Version: 1.0\r\nContent-Type: text/vcard\r\n\r\nBEGIN:VCARD\r\nVERSION:2.1\r\nFN:Ada\r\nEMAIL;ENCODING=QUOTED-PRINTABLE:ada=2Bwork@example.com\r\nEND:VCARD\r\n";
        using var read = new EmailDocumentReader().Read(Encoding.UTF8.GetBytes(eml));
        var review = EmailContactConsolidation.Review(new[] {
            new EmailContactSource("encoded", read.Document),
            new EmailContactSource("same", EmailPortableContentExportTests.Contact("Ada", "ada+work@example.com")),
            new EmailContactSource("unrelated", EmailPortableContentExportTests.Contact("Other", "ada=2Bwork@example.com"))
        });
        Assert.Equal(new[] { "encoded", "same" }, review.Groups[0].SourceIds);
        Assert.Equal("unrelated", Assert.Single(review.Groups[1].SourceIds));
    }

    [Fact]
    public void ResolvesDirectoryAliasesInRetainedEmailSlotsBeyondTheOutlookProjection() {
        string eml = "MIME-Version: 1.0\r\nContent-Type: text/vcard\r\n\r\nBEGIN:VCARD\r\nVERSION:3.0\r\nFN:Ada\r\nEMAIL:a@elsewhere.test\r\nEMAIL:b@elsewhere.test\r\nEMAIL:c@elsewhere.test\r\nEMAIL:alias-ada@example.test\r\nEND:VCARD\r\n";
        using var read = new EmailDocumentReader().Read(Encoding.UTF8.GetBytes(eml));
        using var stream = new MemoryStream(new OabV4Fixture().Build());
        using var directory = OfflineAddressBookSession.Open(stream, "identity.oab");
        var review = EmailContactConsolidation.Review(new[] {
            new EmailContactSource("retained", read.Document),
            new EmailContactSource("primary", EmailPortableContentExportTests.Contact("Ada", "ada@example.test"))
        }, directory: directory.BuildIdentityIndex());
        Assert.Equal(new[] { "retained", "primary" }, Assert.Single(review.Groups).SourceIds);
    }

    [Theory]
    [InlineData("")]
    [InlineData(" ")]
    [InlineData("<>")]
    [InlineData("SMTP:")]
    public void RetainsEmptyOrInvalidEmailFieldsWithoutResolvingThem(string invalid) {
        string eml = "MIME-Version: 1.0\r\nContent-Type: text/vcard\r\n\r\nBEGIN:VCARD\r\nVERSION:3.0\r\nFN:Ada\r\nEMAIL:" + invalid + "\r\nEMAIL:ada@example.test\r\nEND:VCARD\r\n";
        using var read = new EmailDocumentReader().Read(Encoding.UTF8.GetBytes(eml));
        using var stream = new MemoryStream(new OabV4Fixture().Build());
        using var directory = OfflineAddressBookSession.Open(stream, "identity.oab");
        var review = EmailContactConsolidation.Review(new[] { new EmailContactSource("retained", read.Document) }, directory: directory.BuildIdentityIndex());
        var field = Assert.Single(Assert.Single(review.Groups).Fields, candidate => candidate.Choices[0].Name == "EMAIL");
        Assert.Equal(new[] { invalid, "ada@example.test" }, field.Choices[0].Values);
    }

    [Fact]
    public void RecoveredMalformedQuotedPrintableAddressesDoNotCreateIdentityMatches() {
        string eml = "MIME-Version: 1.0\r\nContent-Type: text/vcard\r\n\r\nBEGIN:VCARD\r\nVERSION:2.1\r\nFN:Ada\r\nEMAIL;ENCODING=QUOTED-PRINTABLE:ada=ZZ@example.test\r\nEND:VCARD\r\n";
        using var read = new EmailDocumentReader().Read(Encoding.UTF8.GetBytes(eml));
        var review = EmailContactConsolidation.Review(new[] {
            new EmailContactSource("recovered", read.Document),
            new EmailContactSource("literal", EmailPortableContentExportTests.Contact("Other", "ada=ZZ@example.test"))
        });
        Assert.Equal(2, review.Groups.Count);
        Assert.Contains(review.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_MIME_QUOTED_PRINTABLE_INVALID");
    }
}
