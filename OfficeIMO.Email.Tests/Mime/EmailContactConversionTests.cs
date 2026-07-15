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

        Assert.Contains(oracle.Attachments.OfType<MimePart>(), part => part.ContentType.MimeType == "text/vcard");
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

        Assert.Equal("André", document.Contact!.DisplayName);
        Assert.Equal("André", document.Contact.GivenName);
    }
}
