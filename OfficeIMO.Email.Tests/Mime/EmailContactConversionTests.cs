using MimeKit;
using OfficeIMO.Email;
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class EmailContactConversionTests {
    [Fact]
    public void ConvertsContactThroughVcardWithoutLosingCommonFields() {
        var contact = new OutlookContact {
            DisplayName = "Ada Lovelace",
            GivenName = "Ada",
            Surname = "Lovelace",
            CompanyName = "Analytical Engines",
            Department = "Research",
            JobTitle = "Mathematician",
            Birthday = new DateTimeOffset(1815, 12, 10, 0, 0, 0, TimeSpan.Zero),
            BusinessHomePage = "https://example.com/ada"
        };
        contact.Email1.Address = "ada@example.com";
        contact.Phones.Business = "+44 20 0000 0000";
        contact.Phones.PrimaryFax = "+44 20 0000 0001";
        contact.Phones.Assistant = "+44 20 0000 0002";
        contact.BusinessAddress.Street = "1 Engine Way";
        contact.BusinessAddress.City = "London";
        contact.BusinessAddress.Formatted = "1 Engine Way, London";
        contact.BusinessAddress.CountryCode = "GB";
        contact.HasPicture = true;
        var source = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            OutlookItemKind = OutlookItemKind.Contact,
            MessageClass = "IPM.Contact",
            Subject = "Ada Lovelace",
            Contact = contact
        };
        source.Body.Text = "Contact notes";

        byte[] eml = new EmailDocumentWriter().ToBytes(source, EmailFileFormat.Eml);
        EmailDocument result = new EmailDocumentReader().Read(eml).Document;
        using var oracleStream = new MemoryStream(eml);
        MimeMessage oracle = MimeMessage.Load(oracleStream);

        Assert.Equal("text/vcard", Assert.IsAssignableFrom<MimeEntity>(oracle.Body).ContentType.MimeType);
        Assert.Null(oracle.TextBody);
        Assert.Contains(oracle.BodyParts.OfType<MimePart>(), part => part.ContentType.MimeType == "text/vcard" &&
            !part.IsAttachment);
        Assert.Equal(OutlookItemKind.Contact, result.OutlookItemKind);
        Assert.Equal("Ada Lovelace", result.Contact!.DisplayName);
        Assert.Equal("Ada", result.Contact.GivenName);
        Assert.Equal("Lovelace", result.Contact.Surname);
        Assert.Equal("Analytical Engines", result.Contact.CompanyName);
        Assert.Equal("Research", result.Contact.Department);
        Assert.Equal("ada@example.com", result.Contact.Email1.Address);
        Assert.Equal("+44 20 0000 0000", result.Contact.Phones.Business);
        Assert.Equal("+44 20 0000 0001", result.Contact.Phones.PrimaryFax);
        Assert.Equal("+44 20 0000 0002", result.Contact.Phones.Assistant);
        Assert.Equal("London", result.Contact.BusinessAddress.City);
        Assert.Equal("1 Engine Way, London", result.Contact.BusinessAddress.Formatted);
        Assert.Equal("GB", result.Contact.BusinessAddress.CountryCode);
        Assert.True(result.Contact.HasPicture);
        Assert.Equal("Contact notes", result.Body.Text);
    }

    [Fact]
    public void BlocksOpaqueMapiEntryIdentifiersDuringVcardConversion() {
        var contact = new OutlookContact();
        contact.Email1.Address = "/o=Example/ou=Exchange Administrative Group/cn=Recipients/cn=person";
        contact.Email1.AddressType = "EX";
        contact.Email1.OriginalEntryId = new byte[] { 1, 2, 3 };
        var source = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            OutlookItemKind = OutlookItemKind.Contact,
            Contact = contact
        };

        EmailConversionReport report = new EmailDocumentWriter().AnalyzeConversion(source, EmailFileFormat.Eml);

        Assert.False(report.CanWrite);
        Assert.Contains(report.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_VCARD_OPAQUE_CONTACT_IDENTITY");
    }

    [Fact]
    public void AccumulatesRepeatedVcardTypeParameters() {
        byte[] eml = Encoding.ASCII.GetBytes(
            "Content-Type: text/vcard; charset=utf-8\r\n\r\nBEGIN:VCARD\r\nVERSION:3.0\r\n" +
            "FN:Ada Lovelace\r\nTEL;TYPE=HOME;TYPE=VOICE:+44 20 0000 0000\r\nEND:VCARD\r\n");

        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        Assert.Equal("+44 20 0000 0000", document.Contact!.Phones.Home);
        Assert.Null(document.Contact.Phones.Primary);
    }

    [Fact]
    public void ParsesQuotedVcardParameterDelimiters() {
        byte[] eml = Encoding.ASCII.GetBytes(
            "Content-Type: text/vcard; charset=utf-8\r\n\r\nBEGIN:VCARD\r\nVERSION:3.0\r\n" +
            "FN:Ada Lovelace\r\n" +
            "ADR;TYPE=HOME;LABEL=\"12: Main; Apt\":;;Main Street;Warsaw;Mazovia;00-001;Poland\r\n" +
            "END:VCARD\r\n");

        OutlookPostalAddress address = new EmailDocumentReader().Read(eml).Document.Contact!.HomeAddress;

        Assert.Equal("12: Main; Apt", address.Formatted);
        Assert.Equal("Main Street", address.Street);
        Assert.Equal("Warsaw", address.City);
        Assert.Equal("Mazovia", address.StateOrProvince);
        Assert.Equal("00-001", address.PostalCode);
        Assert.Equal("Poland", address.Country);
    }

    [Fact]
    public void DecodesQuotedPrintableVcardValuesThroughStoreConversion() {
        byte[] eml = Encoding.ASCII.GetBytes(
            "Content-Type: text/vcard; charset=utf-8\r\n\r\nBEGIN:VCARD\r\nVERSION:3.0\r\n" +
            "FN;CHARSET=UTF-8;ENCODING=QUOTED-PRINTABLE:Jos=C3=\r\n =A9\r\n" +
            "N;CHARSET=UTF-8;ENCODING=QUOTED-PRINTABLE:Example;Jos=C3=A9;;;\r\nEND:VCARD\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        EmailDocument roundTrip = new EmailDocumentReader().Read(
            new EmailDocumentWriter().ToBytes(document, EmailFileFormat.OutlookMsg)).Document;

        Assert.Equal("José", document.Contact!.DisplayName);
        Assert.Equal("José", document.Contact.GivenName);
        Assert.Equal("José", roundTrip.Contact!.DisplayName);
        Assert.Equal("José", roundTrip.Contact.GivenName);
    }

    [Fact]
    public void PreservesQuotedVcardParameterBackslashesThroughStoreConversion() {
        byte[] eml = Encoding.ASCII.GetBytes(
            "Content-Type: text/vcard; charset=utf-8\r\n\r\nBEGIN:VCARD\r\nVERSION:3.0\r\n" +
            "FN:Ada Lovelace\r\nADR;TYPE=HOME;LABEL=\"12 \\\"Main\\\" \\\\ St\":;;Main Street;;;;\r\n" +
            "END:VCARD\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        EmailDocument roundTrip = new EmailDocumentReader().Read(
            new EmailDocumentWriter().ToBytes(document, EmailFileFormat.OutlookMsg)).Document;

        Assert.Equal("12 \"Main\" \\ St", document.Contact!.HomeAddress.Formatted);
        Assert.Equal("12 \"Main\" \\ St", roundTrip.Contact!.HomeAddress.Formatted);
    }

    [Fact]
    public void DecodesLegacyTextEscapesInVcardAddressLabels() {
        byte[] eml = Encoding.ASCII.GetBytes(
            "Content-Type: text/vcard; charset=utf-8\r\n\r\nBEGIN:VCARD\r\nVERSION:3.0\r\n" +
            "FN:Ada Lovelace\r\n" +
            "ADR;TYPE=HOME;LABEL=\"Line 1\\nLine 2\\, Suite\\; East\":;;Main Street;;;;\r\n" +
            "END:VCARD\r\n");

        OutlookPostalAddress address = new EmailDocumentReader().Read(eml).Document.Contact!.HomeAddress;

        Assert.Equal("Line 1\nLine 2, Suite; East", address.Formatted);
    }

    [Fact]
    public void MimeVcardProjectionPreservesLegacyLiteralCaretParameters() {
        byte[] eml = Encoding.ASCII.GetBytes(
            "Content-Type: text/vcard; charset=utf-8\r\n\r\nBEGIN:VCARD\r\nVERSION:3.0\r\n" +
            "FN:Ada Lovelace\r\n" +
            "ADR;TYPE=HOME;LABEL=alpha^nbeta:;;Main Street;;;;\r\nEND:VCARD\r\n");

        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        Assert.Equal("alpha^nbeta", document.Contact!.HomeAddress.Formatted);
    }

    [Fact]
    public void ProjectsTextDirectoryVcards() {
        byte[] eml = Encoding.ASCII.GetBytes(
            "Content-Type: text/directory; profile=vCard; charset=utf-8\r\n\r\n" +
            "BEGIN:VCARD\r\nVERSION:3.0\r\nFN:Ada Lovelace\r\n" +
            "EMAIL;TYPE=WORK:ada@example.com\r\nEND:VCARD\r\n");

        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        Assert.Equal(OutlookItemKind.Contact, document.OutlookItemKind);
        Assert.Equal("Ada Lovelace", document.Contact!.DisplayName);
        Assert.Equal("ada@example.com", document.Contact.Email1.Address);
        Assert.True(Assert.Single(document.Attachments).IsProjectedSemanticContent);
    }

    [Fact]
    public void DecodesVcardProjectionUsingTheDeclaredCharset() {
        byte[] prefix = Encoding.ASCII.GetBytes("Content-Type: text/vcard; charset=windows-1252\r\n\r\n");
        byte[] vcard = Encoding.ASCII.GetBytes(
            "BEGIN:VCARD\r\nVERSION:3.0\r\nFN:Andr#\r\nN:Example;Andr#;;;\r\nEND:VCARD\r\n");
        for (int index = 0; index < vcard.Length; index++) {
            if (vcard[index] == (byte)'#') vcard[index] = 0xe9;
        }

        EmailDocument document = new EmailDocumentReader().Read(prefix.Concat(vcard).ToArray()).Document;
        EmailDocument roundTrip = new EmailDocumentReader().Read(
            new EmailDocumentWriter().ToBytes(document, EmailFileFormat.Eml)).Document;

        Assert.Equal("André", document.Contact!.DisplayName);
        Assert.Equal("André", document.Contact.GivenName);
        Assert.Equal("André", roundTrip.Contact!.DisplayName);
        EmailAttachment retained = Assert.Single(roundTrip.Attachments,
            attachment => VCardContentType(attachment.ContentType));
        Assert.Equal("windows-1252", retained.ContentTypeParameters["charset"]);
    }

    [Fact]
    public void MapsHomeAndOtherVcardEmailsToTheirOutlookSlots() {
        byte[] eml = Encoding.ASCII.GetBytes(
            "Content-Type: text/vcard; charset=utf-8\r\n\r\nBEGIN:VCARD\r\nVERSION:3.0\r\n" +
            "FN:Ada Lovelace\r\nEMAIL;TYPE=HOME:ada@home.example\r\n" +
            "X-OFFICEIMO-EMAIL2-DISPLAY-NAME:Home Ada\r\n" +
            "X-OFFICEIMO-EMAIL2-ORIGINAL-DISPLAY-NAME:Ada Original\r\n" +
            "EMAIL;TYPE=OTHER:ada@other.example\r\nEND:VCARD\r\n");

        OutlookContact contact = new EmailDocumentReader().Read(eml).Document.Contact!;

        Assert.Null(contact.Email1.Address);
        Assert.Equal("ada@home.example", contact.Email2.Address);
        Assert.Equal("Home Ada", contact.Email2.DisplayName);
        Assert.Equal("Ada Original", contact.Email2.OriginalDisplayName);
        Assert.Equal("ada@other.example", contact.Email3.Address);
    }

    [Fact]
    public void MarksMultiContactVcardProjectionIncompleteForStoreConversion() {
        byte[] eml = Encoding.ASCII.GetBytes(
            "Content-Type: text/vcard; charset=utf-8\r\n\r\n" +
            "BEGIN:VCARD\r\nVERSION:3.0\r\nFN:First\r\nEND:VCARD\r\n" +
            "BEGIN:VCARD\r\nVERSION:3.0\r\nFN:Second\r\nEND:VCARD\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        EmailConversionReport report = new EmailDocumentWriter().AnalyzeConversion(
            document, EmailFileFormat.OutlookMsg);

        Assert.False(report.CanWrite);
        Assert.Contains(report.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_STORE_SEMANTIC_PROJECTION_INCOMPLETE");
    }

    [Fact]
    public void BlocksVcardsWithMoreThanThreeEmailAddressesBeforeStoreConversion() {
        byte[] eml = Encoding.ASCII.GetBytes(
            "Content-Type: text/vcard; charset=utf-8\r\n\r\nBEGIN:VCARD\r\nVERSION:3.0\r\n" +
            "FN:Ada Lovelace\r\nEMAIL:first@example.com\r\nEMAIL:second@example.com\r\n" +
            "EMAIL:third@example.com\r\nEMAIL:fourth@example.com\r\nEND:VCARD\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        EmailConversionReport report = new EmailDocumentWriter().AnalyzeConversion(
            document, EmailFileFormat.OutlookMsg);

        Assert.False(report.CanWrite);
        Assert.Contains(report.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_STORE_SEMANTIC_PROJECTION_INCOMPLETE");
    }

    [Fact]
    public void BlocksVcardAddressSlotOverflowBeforeStoreConversion() {
        byte[] eml = Encoding.ASCII.GetBytes(
            "Content-Type: text/vcard; charset=utf-8\r\n\r\nBEGIN:VCARD\r\nVERSION:3.0\r\n" +
            "FN:Ada Lovelace\r\nADR;TYPE=HOME:;;First Street;Warsaw;;;\r\n" +
            "ADR;TYPE=HOME:;;Second Street;London;;;\r\nEND:VCARD\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        EmailConversionReport report = new EmailDocumentWriter().AnalyzeConversion(
            document, EmailFileFormat.OutlookMsg);

        Assert.False(report.CanWrite);
        Assert.Contains(report.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_STORE_SEMANTIC_PROJECTION_INCOMPLETE");
    }

    [Fact]
    public void BlocksVcardExtendedAddressComponentsBeforeStoreConversion() {
        byte[] eml = Encoding.ASCII.GetBytes(
            "Content-Type: text/vcard; charset=utf-8\r\n\r\nBEGIN:VCARD\r\nVERSION:3.0\r\n" +
            "FN:Ada Lovelace\r\nADR;TYPE=HOME:;Apt 4;123 Main;Warsaw;;;\r\nEND:VCARD\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        EmailConversionReport report = new EmailDocumentWriter().AnalyzeConversion(
            document, EmailFileFormat.OutlookMsg);

        Assert.Equal("123 Main", document.Contact!.HomeAddress.Street);
        Assert.False(report.CanWrite);
        Assert.Contains(report.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_STORE_SEMANTIC_PROJECTION_INCOMPLETE");
    }

    [Fact]
    public void BlocksExtraVcardOrganizationUnitsBeforeStoreConversion() {
        byte[] eml = Encoding.ASCII.GetBytes(
            "Content-Type: text/vcard; charset=utf-8\r\n\r\nBEGIN:VCARD\r\nVERSION:3.0\r\n" +
            "FN:Ada Lovelace\r\nORG:Example Corp;Sales;West\r\nEND:VCARD\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        EmailConversionReport report = new EmailDocumentWriter().AnalyzeConversion(
            document, EmailFileFormat.OutlookMsg);

        Assert.Equal("Example Corp", document.Contact!.CompanyName);
        Assert.Equal("Sales", document.Contact.Department);
        Assert.False(report.CanWrite);
        Assert.Contains(report.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_STORE_SEMANTIC_PROJECTION_INCOMPLETE");
    }

    [Theory]
    [InlineData("HOME")]
    [InlineData("OTHER")]
    public void BlocksPreferredNonWorkVcardEmailsBeforeStoreConversion(string type) {
        byte[] eml = Encoding.ASCII.GetBytes(
            "Content-Type: text/vcard; charset=utf-8\r\n\r\nBEGIN:VCARD\r\nVERSION:3.0\r\n" +
            "FN:Ada Lovelace\r\nEMAIL;TYPE=" + type + ",PREF:ada@example.com\r\nEND:VCARD\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        EmailConversionReport report = new EmailDocumentWriter().AnalyzeConversion(
            document, EmailFileFormat.OutlookMsg);

        Assert.False(report.CanWrite);
        Assert.Contains(report.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_STORE_SEMANTIC_PROJECTION_INCOMPLETE");
    }

    [Fact]
    public void BlocksVcardUrlSlotOverflowBeforeStoreConversion() {
        byte[] eml = Encoding.ASCII.GetBytes(
            "Content-Type: text/vcard; charset=utf-8\r\n\r\nBEGIN:VCARD\r\nVERSION:3.0\r\n" +
            "FN:Ada Lovelace\r\nURL;TYPE=WORK:https://first.example\r\n" +
            "URL;TYPE=WORK:https://second.example\r\nEND:VCARD\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        EmailConversionReport report = new EmailDocumentWriter().AnalyzeConversion(
            document, EmailFileFormat.OutlookMsg);

        Assert.False(report.CanWrite);
        Assert.Contains(report.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_STORE_SEMANTIC_PROJECTION_INCOMPLETE");
    }

    [Fact]
    public void BlocksVcardPhoneSlotOverflowBeforeStoreConversion() {
        byte[] eml = Encoding.ASCII.GetBytes(
            "Content-Type: text/vcard; charset=utf-8\r\n\r\nBEGIN:VCARD\r\nVERSION:3.0\r\n" +
            "FN:Ada Lovelace\r\nTEL;TYPE=HOME:first\r\nTEL;TYPE=HOME:second\r\n" +
            "TEL;TYPE=HOME:third\r\nEND:VCARD\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        EmailConversionReport report = new EmailDocumentWriter().AnalyzeConversion(
            document, EmailFileFormat.OutlookMsg);

        Assert.False(report.CanWrite);
        Assert.Equal("first", document.Contact!.Phones.Home);
        Assert.Equal("third", document.Contact.Phones.Home2);
        Assert.Contains(report.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_STORE_SEMANTIC_PROJECTION_INCOMPLETE");
    }

    [Theory]
    [InlineData("BDAY")]
    [InlineData("ANNIVERSARY")]
    public void BlocksDateTimeVcardDatesBeforeStoreConversion(string property) {
        byte[] eml = Encoding.ASCII.GetBytes(
            "Content-Type: text/vcard; charset=utf-8\r\n\r\nBEGIN:VCARD\r\nVERSION:3.0\r\n" +
            "FN:Ada Lovelace\r\n" + property + ":19960415T231000Z\r\nEND:VCARD\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        EmailConversionReport report = new EmailDocumentWriter().AnalyzeConversion(
            document, EmailFileFormat.OutlookMsg);

        Assert.False(report.CanWrite);
        Assert.Contains(report.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_STORE_SEMANTIC_PROJECTION_INCOMPLETE");
    }

    [Fact]
    public void BlocksVcardUidBeforeStoreConversion() {
        byte[] eml = Encoding.ASCII.GetBytes(
            "Content-Type: text/vcard; charset=utf-8\r\n\r\nBEGIN:VCARD\r\nVERSION:3.0\r\n" +
            "UID:contact-123\r\nFN:Ada Lovelace\r\nEND:VCARD\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        EmailConversionReport report = new EmailDocumentWriter().AnalyzeConversion(
            document, EmailFileFormat.OutlookMsg);

        Assert.False(report.CanWrite);
        Assert.Contains(report.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_STORE_SEMANTIC_PROJECTION_INCOMPLETE");
    }

    [Theory]
    [InlineData("REV:20260715T080000Z")]
    [InlineData("KIND:group")]
    [InlineData("KIND:org")]
    public void BlocksUnprojectedVcardIdentityMetadataBeforeStoreConversion(string property) {
        byte[] eml = Encoding.ASCII.GetBytes(
            "Content-Type: text/vcard; charset=utf-8\r\n\r\nBEGIN:VCARD\r\nVERSION:3.0\r\n" +
            "FN:Ada Lovelace\r\n" + property + "\r\nEND:VCARD\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        EmailConversionReport report = new EmailDocumentWriter().AnalyzeConversion(
            document, EmailFileFormat.OutlookMsg);

        Assert.False(report.CanWrite);
        Assert.Contains(report.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_STORE_SEMANTIC_PROJECTION_INCOMPLETE");
    }

    [Theory]
    [InlineData("SOURCE:https://example.com/contact.vcf")]
    [InlineData("FBURL:https://example.com/freebusy")]
    [InlineData("CALURI:https://example.com/calendar")]
    [InlineData("CALADRURI:mailto:calendar@example.com")]
    [InlineData("SOUND:https://example.com/name.wav")]
    public void BlocksUnprojectedVcardSourceAndLinkFieldsBeforeStoreConversion(string property) {
        byte[] eml = Encoding.ASCII.GetBytes(
            "Content-Type: text/vcard; charset=utf-8\r\n\r\nBEGIN:VCARD\r\nVERSION:3.0\r\n" +
            "FN:Ada Lovelace\r\n" + property + "\r\nEND:VCARD\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        EmailConversionReport report = new EmailDocumentWriter().AnalyzeConversion(
            document, EmailFileFormat.OutlookMsg);

        Assert.False(report.CanWrite);
        Assert.Contains(report.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_STORE_SEMANTIC_PROJECTION_INCOMPLETE");
    }

    [Theory]
    [InlineData("2.1")]
    [InlineData("4.0")]
    public void BlocksUnsupportedVcardVersionsBeforeStoreConversion(string version) {
        byte[] eml = Encoding.ASCII.GetBytes(
            "Content-Type: text/vcard; charset=utf-8\r\n\r\nBEGIN:VCARD\r\nVERSION:" + version + "\r\n" +
            "FN:Ada Lovelace\r\nEND:VCARD\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        EmailConversionReport report = new EmailDocumentWriter().AnalyzeConversion(
            document, EmailFileFormat.OutlookMsg);

        Assert.False(report.CanWrite);
        Assert.Contains(report.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_STORE_SEMANTIC_PROJECTION_INCOMPLETE");
    }

    [Fact]
    public void BlocksPreferredTypedVcardPhonesBeforeStoreConversion() {
        byte[] eml = Encoding.ASCII.GetBytes(
            "Content-Type: text/vcard; charset=utf-8\r\n\r\nBEGIN:VCARD\r\nVERSION:3.0\r\n" +
            "FN:Ada Lovelace\r\nTEL;TYPE=HOME,PREF:+48-123-456-789\r\nEND:VCARD\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        EmailConversionReport report = new EmailDocumentWriter().AnalyzeConversion(
            document, EmailFileFormat.OutlookMsg);

        Assert.Equal("+48-123-456-789", document.Contact!.Phones.Home);
        Assert.False(report.CanWrite);
        Assert.Contains(report.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_STORE_SEMANTIC_PROJECTION_INCOMPLETE");
    }

    [Fact]
    public void BlocksVcardEmailTypeSlotFallbackBeforeStoreConversion() {
        byte[] eml = Encoding.ASCII.GetBytes(
            "Content-Type: text/vcard; charset=utf-8\r\n\r\nBEGIN:VCARD\r\nVERSION:3.0\r\n" +
            "FN:Ada Lovelace\r\nEMAIL;TYPE=WORK:first@example.com\r\n" +
            "EMAIL;TYPE=WORK:second@example.com\r\nEND:VCARD\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        EmailConversionReport report = new EmailDocumentWriter().AnalyzeConversion(
            document, EmailFileFormat.OutlookMsg);

        Assert.Equal("first@example.com", document.Contact!.Email1.Address);
        Assert.Equal("second@example.com", document.Contact.Email2.Address);
        Assert.False(report.CanWrite);
        Assert.Contains(report.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_STORE_SEMANTIC_PROJECTION_INCOMPLETE");
    }

    [Theory]
    [InlineData("EMAIL;TYPE=WORK:ada@example.com")]
    [InlineData("EMAIL:ada@example.com")]
    [InlineData("TEL;TYPE=VOICE:+48-123-456-789")]
    [InlineData("TEL:+48-123-456-789")]
    [InlineData("URL;TYPE=OTHER:https://example.com")]
    [InlineData("ADR;TYPE=HOME,POSTAL:;;123 Main;Town;;;")]
    public void BlocksVcardTypeSemanticsThatOutlookSlotsCannotPreserve(string property) {
        byte[] eml = Encoding.ASCII.GetBytes(
            "Content-Type: text/vcard; charset=utf-8\r\n\r\nBEGIN:VCARD\r\nVERSION:3.0\r\n" +
            "FN:Ada Lovelace\r\n" + property + "\r\nEND:VCARD\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        EmailConversionReport report = new EmailDocumentWriter().AnalyzeConversion(
            document, EmailFileFormat.OutlookMsg);

        Assert.False(report.CanWrite);
        Assert.Contains(report.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_STORE_SEMANTIC_PROJECTION_INCOMPLETE");
    }

    [Theory]
    [InlineData("NICKNAME:Bob,Rob")]
    [InlineData("IMPP:xmpp:first@example.com\r\nIMPP:xmpp:second@example.com")]
    public void BlocksUnrepresentableMultiValueVcardPropertiesBeforeStoreConversion(string properties) {
        byte[] eml = Encoding.ASCII.GetBytes(
            "Content-Type: text/vcard; charset=utf-8\r\n\r\nBEGIN:VCARD\r\nVERSION:3.0\r\n" +
            "FN:Robert Example\r\n" + properties + "\r\nEND:VCARD\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        EmailConversionReport report = new EmailDocumentWriter().AnalyzeConversion(
            document, EmailFileFormat.OutlookMsg);

        Assert.False(report.CanWrite);
        Assert.Contains(report.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_STORE_SEMANTIC_PROJECTION_INCOMPLETE");
    }

    [Fact]
    public void BlocksDistinctMimeBodyAndVcardNoteBeforeStoreConversion() {
        byte[] eml = Encoding.ASCII.GetBytes(
            "MIME-Version: 1.0\r\nContent-Type: multipart/alternative; boundary=x\r\n\r\n" +
            "--x\r\nContent-Type: text/plain; charset=utf-8\r\n\r\nWrapper text\r\n" +
            "--x\r\nContent-Type: text/vcard; charset=utf-8\r\n\r\n" +
            "BEGIN:VCARD\r\nVERSION:3.0\r\nFN:Ada Lovelace\r\nNOTE:Contact notes\r\n" +
            "END:VCARD\r\n--x--\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        EmailConversionReport report = new EmailDocumentWriter().AnalyzeConversion(
            document, EmailFileFormat.OutlookMsg);

        Assert.False(report.CanWrite);
        Assert.Equal("Wrapper text", document.Body.Text!.Trim());
        Assert.Contains(report.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_STORE_SEMANTIC_PROJECTION_INCOMPLETE");
    }

    [Fact]
    public void KeepsSemanticVcardAsTheMimeBodyWhenOtherAttachmentsExist() {
        byte[] eml = Encoding.ASCII.GetBytes(
            "Subject: Contact\r\nMIME-Version: 1.0\r\nContent-Type: multipart/mixed; boundary=x\r\n\r\n" +
            "--x\r\nContent-Type: text/vcard; charset=utf-8\r\n\r\n" +
            "BEGIN:VCARD\r\nVERSION:3.0\r\nFN:Ada Lovelace\r\nNOTE:Contact notes\r\nEND:VCARD\r\n" +
            "--x\r\nContent-Type: text/plain; charset=utf-8; name=notes.txt\r\n" +
            "Content-Disposition: attachment; filename=notes.txt\r\n\r\nOrdinary attachment\r\n--x--\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        byte[] rewritten = new EmailDocumentWriter().ToBytes(document, EmailFileFormat.Eml);
        using var stream = new MemoryStream(rewritten);
        Multipart mixed = Assert.IsAssignableFrom<Multipart>(MimeMessage.Load(stream).Body);

        Assert.Equal(2, mixed.Count);
        MimePart contactBody = Assert.IsAssignableFrom<MimePart>(mixed[0]);
        Assert.True(VCardContentType(contactBody.ContentType.MimeType));
        Assert.False(contactBody.IsAttachment);
        MimePart attachment = Assert.IsAssignableFrom<MimePart>(mixed[1]);
        Assert.True(attachment.IsAttachment);
        Assert.Equal("notes.txt", attachment.FileName);
    }

    [Fact]
    public void DecodesEscapedVcardTextSequentially() {
        byte[] eml = Encoding.ASCII.GetBytes(
            "Content-Type: text/vcard; charset=utf-8\r\n\r\nBEGIN:VCARD\r\nVERSION:3.0\r\n" +
            "FN:Ada Lovelace\r\nNOTE:Literal \\\\n value\r\nEND:VCARD\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        EmailDocument stored = new EmailDocumentReader().Read(
            new EmailDocumentWriter().ToBytes(document, EmailFileFormat.OutlookMsg)).Document;
        EmailDocument roundTrip = new EmailDocumentReader().Read(
            new EmailDocumentWriter().ToBytes(stored, EmailFileFormat.Eml)).Document;

        Assert.Equal("Literal \\n value", document.Body.Text);
        Assert.Equal("Literal \\n value", roundTrip.Body.Text);
    }

    [Fact]
    public void KeepsOrdinaryVcardAttachmentAlongsideGeneratedContact() {
        var source = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            OutlookItemKind = OutlookItemKind.Contact,
            Subject = "Generated contact",
            Contact = new OutlookContact { DisplayName = "Generated contact" }
        };
        byte[] ordinary = Encoding.ASCII.GetBytes(
            "BEGIN:VCARD\r\nVERSION:3.0\r\nFN:Ordinary file\r\nEND:VCARD\r\n");
        source.Attachments.Add(new EmailAttachment {
            FileName = "ordinary.vcf",
            ContentType = "text/vcard",
            Content = ordinary,
            Length = ordinary.LongLength
        });

        byte[] output = new EmailDocumentWriter().ToBytes(source, EmailFileFormat.Eml);
        using var stream = new MemoryStream(output);
        MimePart[] vcards = MimeMessage.Load(stream).BodyParts.OfType<MimePart>()
            .Where(part => VCardContentType(part.ContentType.MimeType)).ToArray();

        Assert.Equal(2, vcards.Length);
        Assert.Contains(vcards, part => part.IsAttachment && part.FileName == "ordinary.vcf");
        Assert.Contains(vcards, part => !part.IsAttachment && part.FileName == null);
    }

    [Fact]
    public void KeepsAttachedVcardOnAnOrdinaryMessage() {
        byte[] eml = Encoding.ASCII.GetBytes(
            "Subject: Attached contact\r\nMIME-Version: 1.0\r\n" +
            "Content-Type: multipart/mixed; boundary=x\r\n\r\n" +
            "--x\r\nContent-Type: text/plain; charset=utf-8\r\n\r\nPlease import this contact.\r\n" +
            "--x\r\nContent-Type: text/vcard; charset=utf-8\r\n" +
            "Content-Disposition: attachment; filename=ada.vcf\r\n\r\n" +
            "BEGIN:VCARD\r\nVERSION:3.0\r\nFN:Ada Lovelace\r\nEND:VCARD\r\n--x--\r\n");

        EmailDocument document = new EmailDocumentReader().Read(eml).Document;
        EmailDocument roundTrip = new EmailDocumentReader().Read(
            new EmailDocumentWriter().ToBytes(document, EmailFileFormat.Eml)).Document;

        Assert.Equal(OutlookItemKind.Message, document.OutlookItemKind);
        Assert.Null(document.Contact);
        EmailAttachment attachment = Assert.Single(document.Attachments);
        Assert.Equal("ada.vcf", attachment.FileName);
        Assert.Equal(OutlookItemKind.Message, roundTrip.OutlookItemKind);
        Assert.Null(roundTrip.Contact);
        Assert.Equal("ada.vcf", Assert.Single(roundTrip.Attachments).FileName);
    }

    [Fact]
    public void PreservesVcardCategoriesThroughMsgConversion() {
        byte[] eml = Encoding.ASCII.GetBytes(
            "Content-Type: text/vcard; charset=utf-8\r\n\r\nBEGIN:VCARD\r\nVERSION:3.0\r\n" +
            "FN:Ada Lovelace\r\nCATEGORIES:Blue,Project\\, X\r\nEND:VCARD\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        byte[] msg = new EmailDocumentWriter().ToBytes(document, EmailFileFormat.OutlookMsg);
        EmailDocument roundTrip = new EmailDocumentReader().Read(msg).Document;
        EmailDocument regeneratedVcard = new EmailDocumentReader().Read(
            new EmailDocumentWriter().ToBytes(roundTrip, EmailFileFormat.Eml)).Document;

        Assert.Equal(new[] { "Blue", "Project, X" }, document.MessageMetadata.Categories);
        Assert.Equal(new[] { "Blue", "Project, X" }, roundTrip.MessageMetadata.Categories);
        Assert.Equal(new[] { "Blue", "Project, X" }, regeneratedVcard.MessageMetadata.Categories);
    }

    [Fact]
    public void PreservesConfidentialVcardClassThroughStoreConversion() {
        byte[] eml = Encoding.ASCII.GetBytes(
            "Content-Type: text/vcard; charset=utf-8\r\n\r\nBEGIN:VCARD\r\nVERSION:3.0\r\n" +
            "FN:Ada Lovelace\r\nCLASS:CONFIDENTIAL\r\nEND:VCARD\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        EmailDocument storeRoundTrip = new EmailDocumentReader().Read(
            new EmailDocumentWriter().ToBytes(document, EmailFileFormat.OutlookMsg)).Document;
        byte[] regenerated = new EmailDocumentWriter().ToBytes(storeRoundTrip, EmailFileFormat.Eml);
        using var stream = new MemoryStream(regenerated);
        MimePart vcard = Assert.IsAssignableFrom<MimePart>(MimeMessage.Load(stream).Body);
        using var content = new MemoryStream();
        vcard.Content!.DecodeTo(content);

        Assert.True(document.Contact!.IsPrivate);
        Assert.Equal(3, document.MessageMetadata.Sensitivity);
        Assert.True(storeRoundTrip.Contact!.IsPrivate);
        Assert.Equal(3, storeRoundTrip.MessageMetadata.Sensitivity);
        Assert.Contains("CLASS:CONFIDENTIAL", Encoding.UTF8.GetString(content.ToArray()),
            StringComparison.Ordinal);
    }

    private static bool VCardContentType(string? contentType) =>
        string.Equals(contentType, "text/vcard", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(contentType, "text/x-vcard", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(contentType, "application/vcard", StringComparison.OrdinalIgnoreCase);
}
