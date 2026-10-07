using OfficeIMO.Core.Internal;
using OfficeIMO.Epub;
using System.Globalization;
using System.IO.Compression;
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
        BookOnixMessage.RequireSchema(schemas);
        string notification = options.Notification switch {
            BookOnixNotification.Early => "01", BookOnixNotification.Advance => "02", BookOnixNotification.Confirmed => "03",
            _ => throw new ArgumentOutOfRangeException(nameof(options.Notification))
        };
        RequireOnixText(options.SenderName, nameof(options.SenderName));
        RequireOnixText(options.RecordReference, nameof(options.RecordReference));
        RequireOnixText(options.IdentifierId, nameof(options.IdentifierId));
        RequireOnixText(options.PublisherName, nameof(options.PublisherName));
        RequireOnixLanguageCode(options.LanguageCode, nameof(options.LanguageCode));
        if (options.TitleId != null) RequireOnixText(options.TitleId, nameof(options.TitleId));
        if (options.Subtitle != null) RequireOnixText(options.Subtitle, nameof(options.Subtitle));
        XNamespace onix = OnixNamespace;
        IReadOnlyList<XElement> contributorElements = BuildOnixContributors(options.Contributors,
            options.NoContributors, requireDeclaration: true, cancellationToken);

        OnixCommercialParts commercial = BuildOnixCommercial(options.Commercial, options.PublicationDate, cancellationToken);
        IReadOnlyList<XElement> accessibility = BuildOnixAccessibility(options.Accessibility, cancellationToken);
        IReadOnlyList<XElement> subjects = BuildOnixSubjects(options.Subjects, cancellationToken);
        IReadOnlyList<XElement> alternativeTitles = BuildOnixAlternativeTitles(options.AlternativeTitles, cancellationToken);
        IReadOnlyList<XElement> edition = BuildOnixEdition(options.Edition, cancellationToken);
        IReadOnlyList<XElement> collections = BuildOnixCollections(options, cancellationToken);
        IReadOnlyList<XElement> audience = BuildOnixAudience(options.Audience, cancellationToken);
        XElement? collateral = BuildOnixCollateral(options.CollateralTexts, options.SupportingResources, contributorElements, cancellationToken);
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
        XElement titleRecord = options.TitleId == null ? metadataElement.Elements(dc + "title").First() :
            metadataElement.Elements(dc + "title").SingleOrDefault(e => (string?)e.Attribute("id") == options.TitleId)
                ?? throw new ArgumentException("The selected title does not exist in the exported package.", nameof(options.TitleId));
        string title = titleRecord.Value;
        RequireOnixText(title, "title");
        XElement identifier = metadataElement.Elements(dc + "identifier").SingleOrDefault(e => (string?)e.Attribute("id") == options.IdentifierId)
            ?? throw new ArgumentException("The selected ISBN identifier does not exist in the exported package.", nameof(options.IdentifierId));
        string isbn = OfficeIsbn.Normalize(identifier.Value, isbn13: true);

        var titleElement = new XElement(onix + "TitleElement", new XElement(onix + "TitleElementLevel", "01"),
            BuildOnixTitleText(title, options.TitleSorting));
        if (options.Subtitle != null) titleElement.Add(new XElement(onix + "Subtitle", options.Subtitle));
        var descriptive = new XElement(onix + "DescriptiveDetail",
            new XElement(onix + "ProductComposition", "00"), new XElement(onix + "ProductForm", "ED"),
            new XElement(onix + "ProductFormDetail", "E101"),
            accessibility, collections,
            new XElement(onix + "TitleDetail", new XElement(onix + "TitleType", "01"), titleElement));
        descriptive.Add(alternativeTitles);
        descriptive.Add(contributorElements);
        descriptive.Add(edition);
        descriptive.Add(new XElement(onix + "Language", new XElement(onix + "LanguageRole", "01"),
            new XElement(onix + "LanguageCode", options.LanguageCode)));
        descriptive.Add(subjects);
        descriptive.Add(audience);
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
                descriptive, collateral, publishing, commercial.Supplies)));
        byte[] onixBytes = BookOnixMessage.ValidateAndSerialize(message, schemas, 1024L * 1024, cancellationToken);
        return new BookOnixExportResult(onixBytes, publication, ImportDiagnostics, ImportLossAcknowledged);
    }

    private static void RequireOnixHttpUrl(string url, string name) {
        RequireOnixText(url, name);
        if (url.Any(char.IsWhiteSpace) || !Uri.TryCreate(url, UriKind.Absolute, out var parsed) ||
            (parsed.Scheme != Uri.UriSchemeHttps && parsed.Scheme != Uri.UriSchemeHttp) || parsed.UserInfo.Length != 0)
            throw new ArgumentException("Supply an absolute HTTP(S) URL without credentials.", name);
    }

    private static void RequireOnixText(string value, string name) {
        ArgumentException.ThrowIfNullOrWhiteSpace(value, name);
        if (value.Length > 4096) throw new ArgumentException("ONIX text fields cannot exceed 4096 characters in this profile.", name);
        XmlConvert.VerifyXmlChars(value);
    }
}
