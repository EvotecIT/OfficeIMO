using OfficeIMO.Core.Internal;
using OfficeIMO.Epub;
using OfficeIMO.Provenance;
using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Schema;

namespace OfficeIMO.Workflows;

public sealed partial class BookProject {
    /// <summary>Namespace of the supported ONIX 3.1 reference-tag export profile.</summary>
    public const string OnixNamespace = "http://ns.editeur.org/onix/3.1/reference";

    /// <summary>
    /// Exports a single-product digital-download ONIX bibliographic record, validated against a caller-supplied
    /// compiled schema set. The schema must declare ONIXMessage in the ONIX 3.1 reference namespace.
    /// Callers own schema provenance and must not mutate the schema set during export.
    /// The result retains the exact validated EPUB from which the title and selected ISBN were obtained.
    /// This is a complete-record notification, not a block update; commercial fields require explicit caller assertions.
    /// </summary>
    public BookOnixExportResult ExportOnix(BookOnixExportOptions options, XmlSchemaSet schemas,
        EpubWriteOptions? epubOptions = null, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(options);
        ArgumentNullException.ThrowIfNull(schemas);
        cancellationToken.ThrowIfCancellationRequested();
        if (!schemas.IsCompiled || schemas.GlobalElements[new XmlQualifiedName("ONIXMessage", OnixNamespace)] == null)
            throw new ArgumentException("Supply a compiled schema set declaring ONIX 3.1 reference ONIXMessage.", nameof(schemas));
        string notification = options.Notification switch {
            BookOnixNotification.Early => "01", BookOnixNotification.Advance => "02", BookOnixNotification.Confirmed => "03",
            _ => throw new ArgumentOutOfRangeException(nameof(options.Notification))
        };
        RequireOnixText(options.SenderName, nameof(options.SenderName));
        RequireOnixText(options.RecordReference, nameof(options.RecordReference));
        RequireOnixText(options.IdentifierId, nameof(options.IdentifierId));
        RequireOnixText(options.PublisherName, nameof(options.PublisherName));
        RequireOnixText(options.LanguageCode, nameof(options.LanguageCode));
        if (options.LanguageCode.Length != 3 || options.LanguageCode.Any(c => c < 'a' || c > 'z'))
            throw new ArgumentException("Supply a three-letter ONIX list 74 language code.", nameof(options.LanguageCode));
        if (options.Subtitle != null) RequireOnixText(options.Subtitle, nameof(options.Subtitle));
        ArgumentNullException.ThrowIfNull(options.Contributors);
        if (options.Contributors.Count > 100 || options.NoContributors == (options.Contributors.Count != 0))
            throw new ArgumentException("Supply 1-100 credits, or explicitly select NoContributors.", nameof(options));
        var credits = options.Contributors.ToArray();
        XNamespace onix = OnixNamespace;
        var contributorElements = new List<XElement>();
        foreach (BookOnixContributor credit in credits) {
            cancellationToken.ThrowIfCancellationRequested();
            ArgumentNullException.ThrowIfNull(credit);
            RequireOnixText(credit.Name, nameof(credit.Name));
            string role = credit.Role switch {
                BookOnixContributorRole.Author => "A01", BookOnixContributorRole.Editor => "B01",
                BookOnixContributorRole.Translator => "B06", BookOnixContributorRole.Illustrator => "A12",
                BookOnixContributorRole.Other => "Z99", _ => throw new ArgumentOutOfRangeException(nameof(credit.Role))
            };
            contributorElements.Add(new XElement(onix + "Contributor",
                new XElement(onix + "SequenceNumber", contributorElements.Count + 1),
                new XElement(onix + "ContributorRole", role),
                new XElement(onix + (credit.IsOrganization ? "CorporateName" : "PersonName"), credit.Name)));
        }

        OnixCommercialParts commercial = BuildOnixCommercial(options.Commercial, options.PublicationDate, cancellationToken);
        EpubWriteResult publication = Export(epubOptions ?? new EpubWriteOptions(), cancellationToken);
        // Inspect the actual exported metadata, including any writer normalization, without rereading chapters.
        XDocument package;
        using (var input = new MemoryStream(publication.Bytes, false))
        using (var archive = new ZipArchive(input, ZipArchiveMode.Read)) {
            var entry = archive.GetEntry(_publication.PackagePath)
                ?? throw new InvalidDataException("The exported EPUB does not contain its selected package.");
            byte[] bytes = Read(entry, MaximumDeliveryMetadataBytes, cancellationToken);
            using var metadata = new MemoryStream(bytes, false);
            using var reader = XmlReader.Create(metadata, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null });
            package = XDocument.Load(reader);
        }
        XNamespace opf = "http://www.idpf.org/2007/opf", dc = "http://purl.org/dc/elements/1.1/";
        XElement metadataElement = package.Root!.Element(opf + "metadata")!;
        string title = metadataElement.Elements(dc + "title").First().Value;
        RequireOnixText(title, "title");
        XElement identifier = metadataElement.Elements(dc + "identifier").SingleOrDefault(e => (string?)e.Attribute("id") == options.IdentifierId)
            ?? throw new ArgumentException("The selected ISBN identifier does not exist in the exported package.", nameof(options.IdentifierId));
        string isbn = OfficeIsbn.Normalize(identifier.Value, isbn13: true);

        var titleElement = new XElement(onix + "TitleElement", new XElement(onix + "TitleElementLevel", "01"),
            new XElement(onix + "TitleText", title));
        if (options.Subtitle != null) titleElement.Add(new XElement(onix + "Subtitle", options.Subtitle));
        var descriptive = new XElement(onix + "DescriptiveDetail",
            new XElement(onix + "ProductComposition", "00"), new XElement(onix + "ProductForm", "ED"),
            new XElement(onix + "ProductFormDetail", "E101"),
            new XElement(onix + "TitleDetail", new XElement(onix + "TitleType", "01"), titleElement));
        if (options.NoContributors) descriptive.Add(new XElement(onix + "NoContributor"));
        else descriptive.Add(contributorElements);
        descriptive.Add(new XElement(onix + "Language", new XElement(onix + "LanguageRole", "01"),
            new XElement(onix + "LanguageCode", options.LanguageCode)));
        var publishing = new XElement(onix + "PublishingDetail", new XElement(onix + "Publisher",
            new XElement(onix + "PublishingRole", "01"), new XElement(onix + "PublisherName", options.PublisherName)));
        if (commercial.Status != null) publishing.Add(commercial.Status);
        if (options.PublicationDate is { } date) publishing.Add(new XElement(onix + "PublishingDate",
            new XElement(onix + "PublishingDateRole", "01"),
            new XElement(onix + "Date", new XAttribute("dateformat", "00"), date.ToString("yyyyMMdd", CultureInfo.InvariantCulture))));
        publishing.Add(commercial.Rights);
        var message = new XDocument(new XElement(onix + "ONIXMessage", new XAttribute("release", "3.1"),
            new XElement(onix + "Header", new XElement(onix + "Sender", new XElement(onix + "SenderName", options.SenderName)),
                new XElement(onix + "SentDateTime", options.SentAt.UtcDateTime.ToString("yyyyMMdd'T'HHmmss'Z'", CultureInfo.InvariantCulture))),
            new XElement(onix + "Product", new XElement(onix + "RecordReference", options.RecordReference),
                new XElement(onix + "NotificationType", notification),
                new XElement(onix + "ProductIdentifier", new XElement(onix + "ProductIDType", "15"), new XElement(onix + "IDValue", isbn)),
                descriptive, publishing, commercial.Supplies)));
        message.Validate(schemas, (_, args) => {
            cancellationToken.ThrowIfCancellationRequested();
            throw new InvalidDataException("ONIX schema validation failed: " + args.Message, args.Exception);
        });
        cancellationToken.ThrowIfCancellationRequested();
        using var output = new OfficeProvenanceBoundedMemoryStream(1024L * 1024);
        using (var writer = XmlWriter.Create(output, new XmlWriterSettings { Encoding = new UTF8Encoding(false, true), CloseOutput = false }))
            message.Save(writer);
        cancellationToken.ThrowIfCancellationRequested();
        return new BookOnixExportResult(output.ToArray(), publication, ImportDiagnostics, ImportLossAcknowledged);
    }

    private static void RequireOnixText(string value, string name) {
        ArgumentException.ThrowIfNullOrWhiteSpace(value, name);
        if (value.Length > 4096) throw new ArgumentException("ONIX text fields cannot exceed 4096 characters in this profile.", name);
        XmlConvert.VerifyXmlChars(value);
    }
}
