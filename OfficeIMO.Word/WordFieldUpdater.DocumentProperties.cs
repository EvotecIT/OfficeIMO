using OfficeIMO.Core.Internal;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.Word.OpenXml;
using System.Globalization;

namespace OfficeIMO.Word;

internal static partial class WordFieldUpdater {
    private enum CoreFieldProperty {
        Creator, Title, Subject, Keywords, Description, LastModifiedBy,
        Created, Modified, LastPrinted, Revision, Category, Version
    }

    private static bool TryEvaluateCoreProperty(WordDocument document, CoreFieldProperty property,
        WordFieldInventory.ParsedFieldInstruction parsed, out string? value, out WordFieldUpdateStatus status,
        out string message) {
        object? propertyValue = document._ownedModel is { } model
            ? ReadOwnedCoreProperty(model, property)
            : ReadSdkCoreProperty(document.Properties, property);
        string success = $"Updated from built-in document property {property}.";
        return IsCoreDate(property)
            ? TrySetDate((DateTime?)propertyValue, parsed, success, out value, out status, out message)
            : TrySetTextValue((string?)propertyValue, parsed, success, out value, out status, out message);
    }

    private static bool IsCoreDate(CoreFieldProperty property) =>
        property is CoreFieldProperty.Created or CoreFieldProperty.Modified or CoreFieldProperty.LastPrinted;

    private static object? ReadOwnedCoreProperty(WordPackageModel model, CoreFieldProperty property) {
        const string relationship = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
        const int Core = OfficeXmlNamespaces.CoreProperties;
        const int Dublin = OfficeXmlNamespaces.DublinCore;
        const int Terms = OfficeXmlNamespaces.DublinCoreTerms;
        var (namespaceId, name) = property switch {
            CoreFieldProperty.Creator => (Dublin, "creator"),
            CoreFieldProperty.Title => (Dublin, "title"),
            CoreFieldProperty.Subject => (Dublin, "subject"),
            CoreFieldProperty.Keywords => (Core, "keywords"),
            CoreFieldProperty.Description => (Dublin, "description"),
            CoreFieldProperty.LastModifiedBy => (Core, "lastModifiedBy"),
            CoreFieldProperty.Created => (Terms, "created"),
            CoreFieldProperty.Modified => (Terms, "modified"),
            CoreFieldProperty.LastPrinted => (Core, "lastPrinted"),
            CoreFieldProperty.Revision => (Core, "revision"),
            CoreFieldProperty.Category => (Core, "category"),
            CoreFieldProperty.Version => (Core, "version"),
            _ => throw new ArgumentOutOfRangeException(nameof(property))
        };

        OfficeXmlNode element = ReadPackagePropertyPart(model, relationship).Child(namespaceId, name);
        if (element.IsNone) return null;
        string value = element.InnerText;
        if (!IsCoreDate(property)) return value;
        // Package properties expose timestamps in the local time zone.
        return DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime date)
            ? date : (DateTime?)null;
    }

    private static object? ReadSdkCoreProperty(WordBuiltinDocumentProperties properties, CoreFieldProperty property) =>
        property switch {
            CoreFieldProperty.Creator => properties.Creator,
            CoreFieldProperty.Title => properties.Title,
            CoreFieldProperty.Subject => properties.Subject,
            CoreFieldProperty.Keywords => properties.Keywords,
            CoreFieldProperty.Description => properties.Description,
            CoreFieldProperty.LastModifiedBy => properties.LastModifiedBy,
            CoreFieldProperty.Created => properties.Created,
            CoreFieldProperty.Modified => properties.Modified,
            CoreFieldProperty.LastPrinted => properties.LastPrinted,
            CoreFieldProperty.Revision => properties.Revision,
            CoreFieldProperty.Category => properties.Category,
            CoreFieldProperty.Version => properties.Version,
            _ => throw new ArgumentOutOfRangeException(nameof(property))
        };

    private static OfficeXmlNode ReadPackagePropertyPart(WordPackageModel model, string relationshipType) {
        if (!model.Package.PartExists("/_rels/.rels")) return default;
        OfficeXmlTree relationships = model.Package.GetRelationships(null);
        var root = new OfficeXmlNode(relationships, relationships.Root);
        foreach (OfficeXmlNode relationship in root.Children()) {
            string? type = relationship.Attribute("Type");
            if (type == null || OfficeXmlNamespaces.NormalizeRelationshipType(type) != relationshipType ||
                relationship.Attribute("TargetMode") == "External" || relationship.Attribute("Id") is not { } id)
                continue;

            string? target = model.Package.GetRelationshipTarget(null, id);
            if (target != null && model.Package.PartExists(target)) {
                OfficeXmlTree part = model.Package.GetXml(target);
                return new OfficeXmlNode(part, part.Root);
            }
        }
        return default;
    }

    private static int ReadFieldStatistic(WordDocument document, string statistic) {
        if (document._ownedModel is { } model && ReadStoredFieldStatistic(model, statistic) is int stored) return stored;
        return statistic == "Words" ? document.Statistics.Words : document.Statistics.Characters;
    }

    private static int? ReadStoredFieldStatistic(WordPackageModel model, string statistic) {
        const string relationship = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "extended-properties";
        OfficeXmlNode value = ReadPackagePropertyPart(model, relationship).Child(OfficeXmlNamespaces.ExtendedProperties, statistic);
        return int.TryParse(value.InnerText, out int count) ? count : (int?)null;
    }

    private static bool CustomFieldUsesUnsupportedVariant(WordPackageModel model, WordFieldInventory.ParsedFieldInstruction parsed) {
        string? name = parsed.Instructions.FirstOrDefault();
        if (string.IsNullOrWhiteSpace(name)) return false;
        OfficeXmlNode property = FindOwnedFieldCustomProperty(model, TrimQuotes(name!));
        return property.Child(OfficeXmlNamespaces.DocPropsVTypes, "vector").Exists;
    }

    private static OfficeXmlNode FindOwnedFieldCustomProperty(WordPackageModel model, string name) {
        const string relationship = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "custom-properties";
        const int Ns = OfficeXmlNamespaces.CustomProperties;
        OfficeXmlNode properties = ReadPackagePropertyPart(model, relationship);
        foreach (OfficeXmlNode property in properties.Children(Ns, "property")) {
            if (string.Equals(property.Attribute("name"), name, StringComparison.OrdinalIgnoreCase)) return property;
        }
        return default;
    }

    private static KeyValuePair<string, object?> ReadFieldCustomProperty(WordDocument document, string name) {
        if (document._ownedModel is not { } model) {
            KeyValuePair<string, WordCustomProperty> match = document.CustomProperties.FirstOrDefault(pair =>
                string.Equals(pair.Key, name, StringComparison.OrdinalIgnoreCase));
            return string.IsNullOrEmpty(match.Key) ? default : new KeyValuePair<string, object?>(match.Key, match.Value.Value);
        }

        OfficeXmlNode property = FindOwnedFieldCustomProperty(model, name);
        if (property.Attribute("name") is not { } key) return default;
        const int Variant = OfficeXmlNamespaces.DocPropsVTypes;
        OfficeXmlNode value = property.FirstChild;
        if (value.NamespaceId != Variant) return new KeyValuePair<string, object?>(key, null);
        string text = value.InnerText;
        object? parsed = value.LocalName switch {
            "i4" => int.Parse(text),
            "i8" => long.Parse(text),
            "filetime" or "date" => DateTime.Parse(text).ToUniversalTime(),
            "r4" or "r8" => OfficeRoundTripText.ParseDouble(text, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture),
            "lpwstr" => text,
            "bool" => bool.Parse(text),
            "blob" or "oblob" => string.IsNullOrEmpty(text) ? Array.Empty<byte>() : Convert.FromBase64String(text),
            _ => null
        };
        return new KeyValuePair<string, object?>(key, parsed);
    }

    private static KeyValuePair<string, string> ReadFieldVariable(WordDocument document, string variableName) {
        if (document._ownedModel == null) {
            return document.DocumentVariables.FirstOrDefault(pair =>
                string.Equals(pair.Key, variableName, StringComparison.OrdinalIgnoreCase));
        }

        const int W = FieldNamespace;
        OfficeXmlNode settings = WordPackageView.For(document).Main?.FirstPart(WordPackageView.SettingsRelationshipType)
            ?.RootNamed(W, "settings") ?? default;
        string? key = null;
        string value = string.Empty;
        foreach (OfficeXmlNode variable in settings.Child(W, "docVars").Children(W, "docVar")) {
            string? name = variable.Attribute(W, "name");
            if (key == null && string.Equals(name, variableName, StringComparison.OrdinalIgnoreCase)) key = name;
            if (key != null && key == name) value = variable.Attribute(W, "val") ?? string.Empty;
        }
        return key == null ? default : new KeyValuePair<string, string>(key, value);
    }
}
