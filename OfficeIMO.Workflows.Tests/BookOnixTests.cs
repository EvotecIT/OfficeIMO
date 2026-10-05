using OfficeIMO.Epub;
using OfficeIMO.Html;
using System.Security.Cryptography;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Schema;
using Xunit;

namespace OfficeIMO.Workflows.Tests;

public sealed class BookOnixTests {
    [Fact]
    public void ExplicitMetadataAndSelectedIsbnAreBoundToTheExactPublication() {
        var project = Project();
        project.Publication.AddIdentifier("other-isbn", new EpubIdentifierMetadata { Value = "9781861972712", Kind = EpubIdentifierKind.Isbn13 });
        byte[] before = project.ToProjectBytes();
        var result = project.ExportOnix(Options() with {
            Subtitle = "A & B", PublicationDate = new DateOnly(2026, 10, 5),
            Contributors = [new("Author", BookOnixContributorRole.Author), new("Studio & Co", BookOnixContributorRole.Illustrator, true)]
        }, TestSchema());
        var xml = XDocument.Load(new MemoryStream(result.Bytes));
        XNamespace ns = BookProject.OnixNamespace;
        Assert.Equal(ns + "ONIXMessage", xml.Root!.Name);
        Assert.Equal("3.1", (string?)xml.Root.Attribute("release"));
        Assert.Equal("9780306406157", xml.Descendants(ns + "IDValue").Single().Value);
        Assert.Equal("Book & title", xml.Descendants(ns + "TitleText").Single().Value);
        Assert.Equal("A & B", xml.Descendants(ns + "Subtitle").Single().Value);
        Assert.Equal("20261005T100000Z", xml.Descendants(ns + "SentDateTime").Single().Value);
        Assert.Equal("20261005", xml.Descendants(ns + "Date").Single().Value);
        Assert.Equal("Studio & Co", xml.Descendants(ns + "CorporateName").Single().Value);
        Assert.Equal("Author", xml.Descendants(ns + "PersonName").Single().Value);
        Assert.Equal(project.Export().Bytes, result.Publication.Bytes);
        Assert.Equal(Convert.ToHexString(SHA256.HashData(result.Publication.Bytes)), result.PublicationSha256);
        Assert.Equal(before, project.ToProjectBytes());
        Assert.Equal(project.ExportOnix(Options(), TestSchema()).Bytes, project.ExportOnix(Options(), TestSchema()).Bytes);
    }

    [Fact]
    public void ImportReviewAndCancellationCannotBeBypassed() {
        var project = BookProject.FromImport(EpubManuscript.ImportHtml(HtmlConversionDocument.Parse(
            "<title>Book</title><p>Text</p><script>ignored()</script>")));
        project.Publication.AddIdentifier("isbn", new EpubIdentifierMetadata { Value = "9780306406157", Kind = EpubIdentifierKind.Isbn13 });
        Assert.Throws<InvalidOperationException>(() => project.ExportOnix(Options(), TestSchema()));
        project.AcknowledgeImportLoss();
        var result = project.ExportOnix(Options(), TestSchema());
        Assert.True(result.ImportLossAcknowledged);
        Assert.Contains(result.ImportDiagnostics, item => item.LossKind == OfficeConversionLossKind.Omission);
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => project.ExportOnix(Options(), TestSchema(), cancellationToken: cancellation.Token));
    }

    [Theory]
    [InlineData("missing")]
    [InlineData("checksum")]
    [InlineData("opaque")]
    public void ProductIdentityMustBeAnExplicitValidIsbn13(string kind) {
        var project = BookProject.Create("Book");
        if (kind != "missing") project.Publication.AddDublinCoreMetadata("identifier", kind == "checksum" ? "9780306406158" : "urn:uuid:book", "isbn");
        Assert.Throws<ArgumentException>(() => project.ExportOnix(Options(), TestSchema()));
    }

    [Theory]
    [InlineData("unknown-credit")]
    [InlineData("contradictory-credit")]
    [InlineData("too-many-credits")]
    [InlineData("invalid-role")]
    [InlineData("notification")]
    [InlineData("language")]
    [InlineData("oversized")]
    public void UnsupportedAssertionsAreRejectedWithoutChangingTheProject(string kind) {
        var project = Project(); byte[] before = project.ToProjectBytes();
        var options = kind switch {
            "unknown-credit" => Options() with { Contributors = [] },
            "contradictory-credit" => Options() with { NoContributors = true },
            "too-many-credits" => Options() with { Contributors = Enumerable.Repeat(new BookOnixContributor("Name", BookOnixContributorRole.Author), 101).ToArray() },
            "invalid-role" => Options() with { Contributors = [new("Name", (BookOnixContributorRole)99)] },
            "notification" => Options() with { Notification = (BookOnixNotification)99 },
            "language" => Options() with { LanguageCode = "en-US" },
            _ => Options() with { RecordReference = new string('x', 4097) }
        };
        Assert.ThrowsAny<ArgumentException>(() => project.ExportOnix(options, TestSchema()));
        Assert.Equal(before, project.ToProjectBytes());
    }

    [Fact]
    public void TheSuppliedSchemaIsRequiredAndActuallyEnforced() {
        var project = Project();
        Assert.Throws<ArgumentException>(() => project.ExportOnix(Options(), new XmlSchemaSet()));
        // A schema that declares the right root but forbids children must reject the emitted record.
        Assert.Throws<InvalidDataException>(() => project.ExportOnix(Options(), TestSchema(rejectChildren: true)));
        var xml = XDocument.Load(new MemoryStream(project.ExportOnix(Options() with { Contributors = [], NoContributors = true }, TestSchema()).Bytes));
        Assert.Single(xml.Descendants(XName.Get("NoContributor", BookProject.OnixNamespace)));
    }

    [Fact]
    public void SerializedXmlBoundIncludesEscapingExpansion() {
        var project = Project(); byte[] before = project.ToProjectBytes();
        // Each supplied value is within its character/count bounds, but XML escaping expands the output.
        var options = Options() with { Contributors = Enumerable.Repeat(
            new BookOnixContributor(new string('<', 4096), BookOnixContributorRole.Other), 100).ToArray() };
        Assert.Throws<InvalidDataException>(() => project.ExportOnix(options, TestSchema()));
        Assert.Equal(before, project.ToProjectBytes());
    }

    private static BookProject Project() {
        var project = BookProject.Create("Book & title");
        project.Publication.AddIdentifier("isbn", new EpubIdentifierMetadata { Value = "978-0-306-40615-7", Kind = EpubIdentifierKind.Isbn13 });
        return project;
    }
    private static BookOnixExportOptions Options() => new() {
        SenderName = "Example Press", RecordReference = "digital-edition-1", IdentifierId = "isbn", LanguageCode = "eng",
        PublisherName = "Example Press", SentAt = new DateTimeOffset(2026, 10, 5, 12, 0, 0, TimeSpan.FromHours(2)),
        Notification = BookOnixNotification.Confirmed, Contributors = [new("Author", BookOnixContributorRole.Author)]
    };
    // This test double proves schema enforcement, not ONIX conformance. The opt-in Onix fixture runner
    // validates authored products with the unchanged EDItEUR schema and an independent XML validator.
    private static XmlSchemaSet TestSchema(bool rejectChildren = false) {
        string content = rejectChildren ? "" : "<xs:sequence><xs:any minOccurs='0' maxOccurs='unbounded' processContents='skip'/></xs:sequence>";
        using var input = new StringReader($"<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema' targetNamespace='{BookProject.OnixNamespace}' elementFormDefault='qualified'><xs:element name='ONIXMessage'><xs:complexType>{content}<xs:attribute name='release' type='xs:string'/></xs:complexType></xs:element></xs:schema>");
        using var reader = XmlReader.Create(input, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null });
        var schemas = new XmlSchemaSet { XmlResolver = null }; schemas.Add(null, reader); schemas.Compile(); return schemas;
    }
}
