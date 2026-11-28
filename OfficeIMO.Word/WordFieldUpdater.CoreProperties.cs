using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.Word.OpenXml;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Xml;

namespace OfficeIMO.Word;

internal static partial class WordFieldUpdater {
    private const int Core = OfficeXmlNamespaces.CoreProperties;
    private const int Dublin = OfficeXmlNamespaces.DublinCore;
    private const int Terms = OfficeXmlNamespaces.DublinCoreTerms;

    private enum CoreFieldProperty {
        Creator, Title, Subject, Keywords, Description, LastModifiedBy,
        Created, Modified, LastPrinted, Revision, Category, Version,
        Identifier, Language, ContentType, ContentStatus
    }

    private static bool TryEvaluateCoreProperty(WordDocument document, CoreFieldProperty property,
        WordFieldInventory.ParsedFieldInstruction parsed, out string? value, out WordFieldUpdateStatus status,
        out string message) {
        if (document._ownedModel is not { } model || !TryReadOwnedCoreProperty(model, property, out object? propertyValue)) {
            propertyValue = ReadSdkCoreProperty(document.Properties, property);
        }
        string success = $"Updated from built-in document property {property}.";
        return IsCoreDate(property)
            ? TrySetDate((DateTime?)propertyValue, parsed, success, out value, out status, out message)
            : TrySetTextValue((string?)propertyValue, parsed, success, out value, out status, out message);
    }

    private static bool IsCoreDate(CoreFieldProperty property) =>
        property is CoreFieldProperty.Created or CoreFieldProperty.Modified or CoreFieldProperty.LastPrinted;

    private static readonly ConditionalWeakTable<OfficeXmlTree, Dictionary<CoreFieldProperty, object>> OwnedCoreProperties = new();

    private static bool TryReadOwnedCoreProperty(WordPackageModel model, CoreFieldProperty property, out object? value) {
        value = null;
        if (!model.Package.PartExists("/_rels/.rels")) return true;

        const string relationship = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
        const int Relationships = OfficeXmlNamespaces.PackageRelationships;
        OfficeXmlTree relationships = model.Package.GetRelationships(null);
        var root = new OfficeXmlNode(relationships, relationships.Root);
        OfficeXmlNode reference = default;
        foreach (OfficeXmlNode candidate in root.Children(Relationships, "Relationship")) {
            if (candidate.Attribute("Type") != relationship) continue;
            // Let the package reader report its format error before updating the field.
            if (reference.Exists) return false;
            reference = candidate;
        }

        if (reference.Attribute("Id") is not { } id ||
            model.Package.GetRelationshipTarget(null, id) is not { } target ||
            !model.Package.PartExists(target)) return true;

        // Like the package reader, validate the part on first use. Keep the result while the parsed part is unchanged.
        OfficeXmlTree tree = model.Package.GetXml(target);
        var values = tree.Version == 0 ? OwnedCoreProperties.GetValue(tree, ReadOwnedCoreProperties) : ReadOwnedCoreProperties(tree);
        values.TryGetValue(property, out value);
        return true;
    }

    private static Dictionary<CoreFieldProperty, object> ReadOwnedCoreProperties(OfficeXmlTree tree) {
        var root = new OfficeXmlNode(tree, tree.Root);
        if (!root.Is(Core, "coreProperties") || WordXmlNodes.AnyAttribute(root, (_, _, _) => true))
            throw new XmlException("Invalid core properties element.");

        var values = new Dictionary<CoreFieldProperty, object>();
        for (int node = tree.FirstChildOf(tree.Root); node != OfficeXmlTree.NoNode; node = tree.NextSiblingOf(node)) {
            if (tree.KindOf(node) == OfficeXmlTreeNodeKind.Text) {
                if (tree.IsCData(node) || !IsCoreXmlWhitespace(tree.GetText(node)))
                    throw new XmlException("Expected a core property element.");
                continue;
            }

            var element = new OfficeXmlNode(tree, node);
            CoreFieldProperty? property = IdentifyCoreFieldProperty(element);
            if (property is not { } key || values.ContainsKey(key))
                throw new XmlException($"Unexpected core property '{element.LocalName}'.");

            ValidateCorePropertyAttributes(element);
            string text = ReadCorePropertyText(tree, node);
            if (!IsCoreDate(key)) {
                values.Add(key, text);
            } else if (DateTime.TryParseExact(text, CorePropertyTimestampFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime date)) {
                values.Add(key, date);
            } else {
                throw new XmlException($"Core property '{element.LocalName}' does not contain a valid package timestamp.");
            }
        }
        return values;
    }

    private static string ReadCorePropertyText(OfficeXmlTree tree, int property) {
        string? value = null;
        for (int node = tree.FirstChildOf(property); node != OfficeXmlTree.NoNode; node = tree.NextSiblingOf(node)) {
            if (tree.KindOf(node) != OfficeXmlTreeNodeKind.Text || tree.IsCData(node))
                throw new XmlException("Core properties can only contain plain text.");

            string text = tree.GetText(node);
            if (IsCoreXmlWhitespace(text)) continue;
            if (value != null) throw new XmlException("A core property cannot contain multiple text nodes.");
            value = text;
        }
        return value ?? string.Empty;
    }

    private static bool IsCoreXmlWhitespace(string value) => value.All(character => character is ' ' or '\t' or '\r' or '\n');

    private static void ValidateCorePropertyAttributes(OfficeXmlNode property) {
        if (property.NamespaceId != Terms) {
            if (WordXmlNodes.AnyAttribute(property, (_, _, _) => true))
                throw new XmlException($"Unexpected attribute on core property '{property.LocalName}'.");
            return;
        }

        const string schemaInstance = "http://www.w3.org/2001/XMLSchema-instance";
        const string terms = "http://purl.org/dc/terms/";
        var attributes = WordXmlNodes.Attributes(property);
        if (attributes.Count == 1 && attributes[0].NamespaceUri == schemaInstance && attributes[0].LocalName == "type") {
            string type = attributes[0].Value;
            int colon = type.IndexOf(':');
            if (colon > 0 && type.Substring(colon + 1) == "W3CDTF" &&
                WordXmlNodes.LookupNamespace(property, type.Substring(0, colon)) == terms) return;
        }
        throw new XmlException($"Core property '{property.LocalName}' requires a Dublin Core timestamp type.");
    }

    private static CoreFieldProperty? IdentifyCoreFieldProperty(OfficeXmlNode element) => element.NamespaceId switch {
        Dublin => element.LocalName switch {
            "creator" => CoreFieldProperty.Creator,
            "title" => CoreFieldProperty.Title,
            "subject" => CoreFieldProperty.Subject,
            "description" => CoreFieldProperty.Description,
            "identifier" => CoreFieldProperty.Identifier,
            "language" => CoreFieldProperty.Language,
            _ => null
        },
        Terms => element.LocalName switch {
            "created" => CoreFieldProperty.Created,
            "modified" => CoreFieldProperty.Modified,
            _ => null
        },
        Core => element.LocalName switch {
            "keywords" => CoreFieldProperty.Keywords,
            "lastModifiedBy" => CoreFieldProperty.LastModifiedBy,
            "lastPrinted" => CoreFieldProperty.LastPrinted,
            "revision" => CoreFieldProperty.Revision,
            "category" => CoreFieldProperty.Category,
            "version" => CoreFieldProperty.Version,
            "contentType" => CoreFieldProperty.ContentType,
            "contentStatus" => CoreFieldProperty.ContentStatus,
            _ => null
        },
        _ => null
    };

    private static readonly string[] CorePropertyTimestampFormats = CreateCorePropertyTimestampFormats();

    private static string[] CreateCorePropertyTimestampFormats() {
        const string time = "yyyy-MM-ddTHH:mm:ss";
        var formats = new List<string>();
        for (int decimals = 0; decimals <= 7; decimals++) {
            string format = decimals == 0 ? time : time + "." + new string('f', decimals);
            string signed = @"\-" + format;
            // The package reader also accepts a leading minus and excludes a lone fractional digit without a zone on unsigned dates.
            if (decimals != 1) formats.Add(format);
            formats.Add(format + "Z");
            formats.Add(format + "zzz");
            formats.Add(signed);
            formats.Add(signed + "Z");
            formats.Add(signed + "zzz");
        }
        return formats.ToArray();
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

}
