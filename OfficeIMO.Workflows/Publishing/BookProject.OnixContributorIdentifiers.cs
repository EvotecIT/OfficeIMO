using System.Globalization;
using System.Xml.Linq;

namespace OfficeIMO.Workflows;

public sealed partial class BookProject {
    private static IReadOnlyList<XElement> BuildOnixContributorIdentifiers(IReadOnlyList<BookOnixContributorIdentifier> identifiers) {
        ArgumentNullException.ThrowIfNull(identifiers);
        if (identifiers.Count > 16) throw new ArgumentException("At most 16 contributor identifiers are supported.", nameof(identifiers));
        XNamespace ns = OnixNamespace;
        var result = new List<XElement>();
        var schemes = new HashSet<(BookOnixContributorIdentifierType, string?)>();
        foreach (var identifier in identifiers) {
            ArgumentNullException.ThrowIfNull(identifier);
            if (!Enum.IsDefined(identifier.Type)) throw new ArgumentOutOfRangeException(nameof(identifier.Type));
            RequireOnixText(identifier.Value, nameof(identifier.Value));
            if (identifier.Value.Length > 100) throw new ArgumentException("Contributor identifiers are limited to 100 characters.", nameof(identifiers));
            if (identifier.Type == BookOnixContributorIdentifierType.Proprietary) {
                RequireOnixText(identifier.SchemeName!, nameof(identifier.SchemeName));
                if (identifier.SchemeName!.Length > 100) throw new ArgumentException("Contributor scheme names are limited to 100 characters.", nameof(identifiers));
            } else {
                if (identifier.SchemeName != null) throw new ArgumentException("Only proprietary identifiers carry a scheme name.", nameof(identifiers));
                if (identifier.Value.Length != 16 || identifier.Value.Take(15).Any(c => c < '0' || c > '9') ||
                    !(identifier.Value[15] is >= '0' and <= '9' or 'X'))
                    throw new ArgumentException("ISNI and ORCID require 15 ASCII digits and a final digit or uppercase X, without separators.", nameof(identifiers));
            }
            if (!schemes.Add((identifier.Type, identifier.SchemeName))) throw new ArgumentException("Contributor identifier schemes must be distinct.", nameof(identifiers));
            result.Add(new XElement(ns + "NameIdentifier",
                new XElement(ns + "NameIDType", ((int)identifier.Type).ToString("00", CultureInfo.InvariantCulture)),
                identifier.SchemeName != null ? new XElement(ns + "IDTypeName", identifier.SchemeName) : null,
                new XElement(ns + "IDValue", identifier.Value)));
        }
        return result;
    }

    private static IReadOnlyList<XElement> BuildOnixResourceContributorReferences(
        IReadOnlyList<BookOnixContributorIdentifier> references, IReadOnlyList<XElement> contributors,
        ref int textBudget, CancellationToken token) {
        ArgumentNullException.ThrowIfNull(references);
        if (references.Count > 16) throw new ArgumentException("At most 16 resource contributor references are supported.", nameof(references));
        XNamespace ns = OnixNamespace;
        var result = new List<XElement>();
        var seen = new HashSet<(string, string)>();
        foreach (var reference in references) {
            token.ThrowIfCancellationRequested();
            XElement identity = BuildOnixContributorIdentifiers([reference]).Single();
            string type = identity.Element(ns + "NameIDType")!.Value;
            string value = identity.Element(ns + "IDValue")!.Value;
            string? scheme = (string?)identity.Element(ns + "IDTypeName");
            bool SameValue(XElement identifier) => (string?)identifier.Element(ns + "NameIDType") == type &&
                (string?)identifier.Element(ns + "IDValue") == value;
            var matches = contributors.Where(c => c.Elements(ns + "NameIdentifier").Any(SameValue)).ToArray();
            if (matches.Length == 0 || !matches.SelectMany(c => c.Elements(ns + "NameIdentifier")).Where(SameValue)
                .All(i => (string?)i.Element(ns + "IDTypeName") == scheme))
                throw new ArgumentException("Each resource identity must match a product contributor without proprietary scheme ambiguity.", nameof(references));
            if (!seen.Add((type, value))) throw new ArgumentException("Resource contributor references must be distinct.", nameof(references));
            ConsumeOnixResourceText(value, ref textBudget);
            result.Add(new XElement(ns + "ResourceFeature", new XElement(ns + "ResourceFeatureType", type switch {
                "16" => "05", "01" => "06", "21" => "11", _ => throw new InvalidOperationException()
            }), new XElement(ns + "FeatureValue", value)));
        }
        return result;
    }
}
