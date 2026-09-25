using OfficeIMO.OpenDocument;
using OfficeIMO.PowerPoint;
using System.Xml.Linq;

namespace OfficeIMO.PowerPoint.OpenDocument;

public static partial class PowerPointOpenDocumentConversionExtensions {
    private static void CopyPowerPointMetadata(PowerPointPresentation source,
        OdpPresentation target, OdfConversionReport report) {
        PowerPointBuiltinDocumentProperties properties = source.BuiltinDocumentProperties;
        target.Metadata.Title = properties.Title;
        target.Metadata.Subject = properties.Subject;
        target.Metadata.Description = properties.Description;
        target.Metadata.Creator = properties.Creator;
        target.Metadata.LastModifiedBy = properties.LastModifiedBy;
        target.Metadata.Language = source.OpenXmlDocument.PackageProperties.Language;
        if (properties.Created.HasValue)
            target.Metadata.CreationDate = AsUtc(properties.Created.Value);
        if (properties.Modified.HasValue)
            target.Metadata.ModifiedDate = AsUtc(properties.Modified.Value);

        int unsupported = CountNonempty(properties.Keywords, properties.Category,
            properties.Revision, properties.Version);
        if (properties.LastPrinted.HasValue) unsupported++;
        unsupported += CountNonempty(source.OpenXmlDocument.PackageProperties.ContentStatus,
            source.OpenXmlDocument.PackageProperties.ContentType,
            source.OpenXmlDocument.PackageProperties.Identifier);
        unsupported += source.OpenXmlDocument.CustomFilePropertiesPart?.Properties?
            .ChildElements.Count ?? 0;
        unsupported += CountNonempty(source.ApplicationProperties.Company,
            source.ApplicationProperties.Manager);
        AddUnsupported(report, "document-metadata", unsupported,
            "PowerPoint keywords, category, revision, version, last-print time, company, manager, package content fields, and custom properties have no exact mapping in the current ODP metadata surface.");
    }

    private static void CopyOdpMetadata(OdpPresentation source,
        PowerPointPresentation target, OdfConversionReport report) {
        target.BuiltinDocumentProperties.Title = source.Metadata.Title;
        target.BuiltinDocumentProperties.Subject = source.Metadata.Subject;
        target.BuiltinDocumentProperties.Description = source.Metadata.Description;
        target.BuiltinDocumentProperties.Creator = source.Metadata.Creator;
        target.BuiltinDocumentProperties.LastModifiedBy = source.Metadata.LastModifiedBy;
        target.OpenXmlDocument.PackageProperties.Language = source.Metadata.Language;
        target.BuiltinDocumentProperties.Created = source.Metadata.CreationDate?.UtcDateTime;
        target.BuiltinDocumentProperties.Modified = source.Metadata.ModifiedDate?.UtcDateTime;

        var mapped = new HashSet<System.Xml.Linq.XName> {
            OdfNamespaces.Dc + "title", OdfNamespaces.Dc + "subject",
            OdfNamespaces.Dc + "description", OdfNamespaces.Meta + "initial-creator",
            OdfNamespaces.Dc + "creator",
            OdfNamespaces.Meta + "creation-date", OdfNamespaces.Dc + "date",
            OdfNamespaces.Meta + "generator", OdfNamespaces.Dc + "language"
        };
        int unsupported = source.Package.GetXml("meta.xml")
            .Descendants(OdfNamespaces.Office + "meta")
            .Elements().Count(element => !mapped.Contains(element.Name));
        AddUnsupported(report, "document-metadata", unsupported,
            "ODF keywords, custom properties, and other metadata outside the shared core fields were not transferred to PowerPoint.");
    }

    private static DateTimeOffset AsUtc(DateTime value) => new(
        value.Kind == DateTimeKind.Unspecified
            ? DateTime.SpecifyKind(value, DateTimeKind.Utc)
            : value.ToUniversalTime());

    private static int CountNonempty(params string?[] values) =>
        values.Count(value => !string.IsNullOrWhiteSpace(value));
}
