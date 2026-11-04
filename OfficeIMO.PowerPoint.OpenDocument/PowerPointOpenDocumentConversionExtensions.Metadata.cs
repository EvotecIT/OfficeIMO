using OfficeIMO.OpenDocument;
using OfficeIMO.PowerPoint;
using System.Xml.Linq;

namespace OfficeIMO.PowerPoint.OpenDocument;

public static partial class PowerPointOpenDocumentConversionExtensions {
    private static readonly Lazy<IReadOnlyDictionary<string, string>> DefaultPowerPointExtendedMetadata = new(() => {
        using PowerPointDocument baseline = PowerPointDocument.Create(new MemoryStream(),
            new PowerPointCreateOptions());
        return ReadExtendedMetadata(baseline);
    });

    private static IReadOnlyDictionary<string, string> ReadExtendedMetadata(PowerPointDocument presentation) =>
        presentation.OpenXmlDocument.ExtendedFilePropertiesPart?.Properties?.ChildElements
            .Where(element => element.LocalName is not (
                "Company" or "Manager" or "Slides" or "Notes" or "HiddenSlides" or
                "Words" or "Paragraphs" or "Characters" or "CharactersWithSpaces" or
                "Lines" or "Pages" or "Bytes" or "HeadingPairs" or "TitlesOfParts"))
            .ToDictionary(element => element.LocalName, element => element.InnerText, StringComparer.Ordinal)
        ?? new Dictionary<string, string>(StringComparer.Ordinal);

    private static void CopyPowerPointMetadata(PowerPointDocument source,
        OdpPresentation target, OdfConversionReport report) {
        PowerPointBuiltinDocumentProperties properties = source.Properties;
        target.Properties.Title = properties.Title;
        target.Properties.Subject = properties.Subject;
        target.Properties.Description = properties.Description;
        target.Properties.Creator = properties.Creator;
        target.Properties.LastModifiedBy = properties.LastModifiedBy;
        target.Properties.Language = source.OpenXmlDocument.PackageProperties.Language;
        if (properties.Created.HasValue)
            target.Properties.CreationDate = AsUtc(properties.Created.Value);
        if (properties.Modified.HasValue)
            target.Properties.ModifiedDate = AsUtc(properties.Modified.Value);

        int unsupported = CountNonempty(properties.Keywords, properties.Category,
            properties.Revision, properties.Version);
        if (properties.LastPrinted.HasValue) unsupported++;
        unsupported += CountNonempty(source.OpenXmlDocument.PackageProperties.ContentStatus,
            source.OpenXmlDocument.PackageProperties.ContentType,
            source.OpenXmlDocument.PackageProperties.Identifier);
        unsupported += source.OpenXmlDocument.CustomFilePropertiesPart?.Properties?
            .ChildElements.Count ?? 0;
        var applicationProperties = source.OpenXmlDocument.ExtendedFilePropertiesPart?.Properties;
        unsupported += CountNonempty(applicationProperties?.Company?.Text,
            applicationProperties?.Manager?.Text);
        foreach (KeyValuePair<string, string> property in ReadExtendedMetadata(source)) {
            if (!string.IsNullOrWhiteSpace(property.Value) &&
                (!DefaultPowerPointExtendedMetadata.Value.TryGetValue(property.Key, out string? baseline) ||
                 !string.Equals(property.Value, baseline, StringComparison.Ordinal))) unsupported++;
        }
        AddUnsupported(report, "document-metadata", unsupported,
            "PowerPoint keywords, category, revision, version, last-print time, extended application properties, package content fields, and custom properties have no exact mapping in the current ODP metadata surface.");
    }

    private static void CopyOdpMetadata(OdpPresentation source,
        PowerPointDocument target, OdfConversionReport report) {
        target.Properties.Title = source.Properties.Title;
        target.Properties.Subject = source.Properties.Subject;
        target.Properties.Description = source.Properties.Description;
        target.Properties.Creator = source.Properties.Creator;
        target.Properties.LastModifiedBy = source.Properties.LastModifiedBy;
        target.OpenXmlDocument.PackageProperties.Language = source.Properties.Language;
        target.Properties.Created = source.Properties.CreationDate?.UtcDateTime;
        target.Properties.Modified = source.Properties.ModifiedDate?.UtcDateTime;

        var mapped = new HashSet<System.Xml.Linq.XName> {
            OdfNamespaces.Dc + "title", OdfNamespaces.Dc + "subject",
            OdfNamespaces.Dc + "description", OdfNamespaces.Meta + "initial-creator",
            OdfNamespaces.Dc + "creator",
            OdfNamespaces.Meta + "creation-date", OdfNamespaces.Dc + "date",
            OdfNamespaces.Dc + "language"
        };
        int unsupported = source.Package.GetXml("meta.xml")
            .Descendants(OdfNamespaces.Office + "meta")
            .Elements().Count(element => !mapped.Contains(element.Name) &&
                (element.Name != OdfNamespaces.Meta + "generator" ||
                 !string.Equals(element.Value, "OfficeIMO.OpenDocument", StringComparison.Ordinal)));
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
