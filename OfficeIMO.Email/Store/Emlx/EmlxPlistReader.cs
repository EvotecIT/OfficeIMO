using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.Email.Store;

/// <summary>Small, bounded XML property-list reader for EMLX metadata.</summary>
internal static class EmlxPlistReader {
    internal static bool LooksLikeBinaryPlist(byte[] data, int offset) {
        byte[] signature = Encoding.ASCII.GetBytes("bplist00");
        if (data.Length - offset < signature.Length) return false;
        for (int index = 0; index < signature.Length; index++) {
            if (data[offset + index] != signature[index]) return false;
        }
        return true;
    }

    internal static IReadOnlyDictionary<string, object?> Read(byte[] data, int offset,
        EmailStoreReaderOptions options, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        var settings = new XmlReaderSettings {
            DtdProcessing = DtdProcessing.Ignore,
            XmlResolver = null,
            MaxCharactersInDocument = options.MaxXmlCharactersPerItem,
            MaxCharactersFromEntities = 0,
            IgnoreComments = true
        };
        XDocument document;
        using (var stream = new MemoryStream(data, offset, data.Length - offset, writable: false))
        using (XmlReader reader = XmlReader.Create(stream, settings)) {
            document = XDocument.Load(reader, LoadOptions.None);
        }

        XElement? root = document.Root;
        if (root == null || !string.Equals(root.Name.LocalName, "plist", StringComparison.OrdinalIgnoreCase)) {
            throw new InvalidDataException("The EMLX metadata trailer is not an Apple property list.");
        }
        XElement? dictionary = root.Elements().FirstOrDefault(element =>
            string.Equals(element.Name.LocalName, "dict", StringComparison.OrdinalIgnoreCase));
        if (dictionary == null) throw new InvalidDataException("The EMLX property list has no root dictionary.");
        int propertyCount = 0;
        object? value = ParseValue(dictionary, options, cancellationToken, ref propertyCount, depth: 0);
        if (value is Dictionary<string, object?> values) return values;
        throw new InvalidDataException("The EMLX property-list root is not a dictionary.");
    }

    private static object? ParseValue(XElement element, EmailStoreReaderOptions options,
        CancellationToken cancellationToken, ref int propertyCount, int depth) {
        cancellationToken.ThrowIfCancellationRequested();
        if (depth > options.MaxBTreeDepth) {
            throw new EmailStoreLimitExceededException(nameof(EmailStoreReaderOptions.MaxBTreeDepth),
                depth, options.MaxBTreeDepth);
        }
        string name = element.Name.LocalName;
        if (string.Equals(name, "dict", StringComparison.OrdinalIgnoreCase)) {
            return ParseDictionary(element, options, cancellationToken, ref propertyCount, depth);
        }
        if (string.Equals(name, "array", StringComparison.OrdinalIgnoreCase)) {
            var values = new List<object?>();
            foreach (XElement child in element.Elements()) {
                propertyCount++;
                EnforcePropertyCount(propertyCount, options);
                values.Add(ParseValue(child, options, cancellationToken, ref propertyCount, depth + 1));
            }
            return values.ToArray();
        }
        if (string.Equals(name, "true", StringComparison.OrdinalIgnoreCase)) return true;
        if (string.Equals(name, "false", StringComparison.OrdinalIgnoreCase)) return false;
        if (string.Equals(name, "integer", StringComparison.OrdinalIgnoreCase)) {
            if (!long.TryParse(element.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out long integer)) {
                throw new FormatException("An EMLX property-list integer is invalid.");
            }
            return integer;
        }
        if (string.Equals(name, "real", StringComparison.OrdinalIgnoreCase)) {
            if (!double.TryParse(element.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double real)) {
                throw new FormatException("An EMLX property-list real value is invalid.");
            }
            return real;
        }
        if (string.Equals(name, "date", StringComparison.OrdinalIgnoreCase)) {
            if (DateTimeOffset.TryParse(element.Value, CultureInfo.InvariantCulture,
                    DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.AssumeUniversal,
                    out DateTimeOffset date)) return date;
            throw new FormatException("An EMLX property-list date is invalid.");
        }
        if (string.Equals(name, "data", StringComparison.OrdinalIgnoreCase)) {
            try { return Convert.FromBase64String(element.Value); }
            catch (FormatException) { throw new FormatException("An EMLX property-list data value is invalid base64."); }
        }
        return element.Value;
    }

    private static Dictionary<string, object?> ParseDictionary(XElement dictionary,
        EmailStoreReaderOptions options, CancellationToken cancellationToken, ref int propertyCount, int depth) {
        var result = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
        List<XElement> elements = dictionary.Elements().ToList();
        for (int index = 0; index < elements.Count; index += 2) {
            cancellationToken.ThrowIfCancellationRequested();
            XElement key = elements[index];
            if (!string.Equals(key.Name.LocalName, "key", StringComparison.OrdinalIgnoreCase) ||
                index + 1 >= elements.Count) {
                throw new InvalidDataException("An EMLX property-list dictionary contains an unpaired key or value.");
            }
            propertyCount++;
            EnforcePropertyCount(propertyCount, options);
            if (result.ContainsKey(key.Value)) {
                throw new InvalidDataException("An EMLX property-list dictionary contains a duplicate key.");
            }
            result.Add(key.Value, ParseValue(elements[index + 1], options,
                cancellationToken, ref propertyCount, depth + 1));
        }
        return result;
    }

    private static void EnforcePropertyCount(int propertyCount, EmailStoreReaderOptions options) {
        if (propertyCount > options.MaxPropertiesPerItem) {
            throw new EmailStoreLimitExceededException(nameof(EmailStoreReaderOptions.MaxPropertiesPerItem),
                propertyCount, options.MaxPropertiesPerItem);
        }
    }
}
