using OfficeIMO.Core.Internal;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.Word.OpenXml;
using System.Globalization;

namespace OfficeIMO.Word;

internal static partial class WordFieldUpdater {
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
        return FindOwnedCustomPropertyVariant(property).Is(OfficeXmlNamespaces.DocPropsVTypes, "vector");
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
        OfficeXmlNode value = FindOwnedCustomPropertyVariant(property);
        if (value.IsNone) return new KeyValuePair<string, object?>(key, null);
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

    private static readonly string[] CustomPropertyVariantOrder = {
        "i4", "filetime", "r4", "lpwstr", "bool", "r8", "i8", "blob", "oblob", "vector", "empty", "date"
    };

    private static OfficeXmlNode FindOwnedCustomPropertyVariant(OfficeXmlNode property) {
        // The default package reader exposes direct typed values beside ignorable children.
        // Explicit compatibility processing already uses the SDK's processed package.
        OfficeXmlNode value = default;
        int priority = CustomPropertyVariantOrder.Length;
        foreach (OfficeXmlNode child in property.Children()) {
            if (child.NamespaceId != OfficeXmlNamespaces.DocPropsVTypes) continue;
            int candidate = Array.IndexOf(CustomPropertyVariantOrder, child.LocalName);
            if (candidate < 0 || candidate >= priority) continue;
            if (candidate == 0) return child;
            value = child;
            priority = candidate;
        }
        return value;
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
