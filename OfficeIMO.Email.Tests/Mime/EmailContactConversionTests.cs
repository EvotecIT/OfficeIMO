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

        byte[] eml = new EmailDocumentWriter().ToBytes(source, EmailFileFormat.Eml);
        EmailDocument result = new EmailDocumentReader().Read(eml).Document;
        using var oracleStream = new MemoryStream(eml);
        MimeMessage oracle = MimeMessage.Load(oracleStream);

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

    private static bool VCardContentType(string? contentType) =>
        string.Equals(contentType, "text/vcard", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(contentType, "text/x-vcard", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(contentType, "application/vcard", StringComparison.OrdinalIgnoreCase);
}
