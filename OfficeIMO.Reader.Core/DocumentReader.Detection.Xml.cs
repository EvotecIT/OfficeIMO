using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;
using System.Xml;

namespace OfficeIMO.Reader;

internal static partial class DocumentReaderEngine {
    private const string FlatOpcPackageNamespace = "http://schemas.microsoft.com/office/2006/xmlPackage";

    // This is bounded identification only. Word validates the root relationship and the whole package before loading.
    private static bool ContainsFlatOpcWordMainPart(string prefix) {
        try {
            using var text = new StringReader(prefix);
            using XmlReader reader = XmlReader.Create(text, new XmlReaderSettings {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null
            });
            while (reader.Read()) {
                if (reader.NodeType != XmlNodeType.Element || reader.Depth != 1 ||
                    reader.LocalName != "part" || reader.NamespaceURI != FlatOpcPackageNamespace) continue;
                string? contentType = reader.GetAttribute("contentType", FlatOpcPackageNamespace);
                if (contentType is "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml" or
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml" or
                    "application/vnd.ms-word.document.macroEnabled.main+xml" or
                    "application/vnd.ms-word.template.macroEnabledTemplate.main+xml") return true;
            }
        } catch (XmlException) {
            // A bounded prefix can stop inside an XML token. Leave incomplete evidence as generic XML.
        }
        return false;
    }

    /// <summary>Checks an XML root before a format-owned import and restores the input position.</summary>
    internal static bool HasXmlRoot(Stream input, string localName, string namespaceUri,
        int maximumBytes, CancellationToken token) {
        if (!input.CanSeek) return false;
        token.ThrowIfCancellationRequested();
        long start = input.Position;
        try {
            using var prefix = new MemoryStream(ReadDetectionPrefix(input, maximumBytes), writable: false);
            using XmlReader reader = XmlReader.Create(prefix, new XmlReaderSettings {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null
            });
            while (reader.Read()) {
                token.ThrowIfCancellationRequested();
                if (reader.NodeType == XmlNodeType.Element)
                    return reader.LocalName == localName && reader.NamespaceURI == namespaceUri;
            }
        } catch (XmlException) {
            // Binary files and incomplete XML roots must use their normal owner admission.
        } finally {
            input.Position = start;
        }
        return false;
    }

    private static bool TryResolveXmlRootNamespace(
        string qualifiedName,
        string rootTag,
        string documentType,
        out string namespaceUri) {
        namespaceUri = string.Empty;
        try {
            XmlConvert.VerifyName(qualifiedName);
        } catch (XmlException) {
            return false;
        }

        int separator = qualifiedName.IndexOf(':');
        if (separator == 0 || separator == qualifiedName.Length - 1 ||
            separator >= 0 && qualifiedName.IndexOf(':', separator + 1) >= 0) return false;
        string namespaceAttribute = separator < 0 ? "xmlns" : "xmlns:" + qualifiedName.Substring(0, separator);
        int position = 1 + qualifiedName.Length;
        bool foundNamespace = false;
        var attributeNames = new HashSet<string>(StringComparer.Ordinal);
        while (position < rootTag.Length) {
            while (position < rootTag.Length && char.IsWhiteSpace(rootTag[position])) position++;
            if (position >= rootTag.Length) return false;
            if (rootTag[position] == '>') break;
            if (rootTag[position] == '/' && position + 1 < rootTag.Length && rootTag[position + 1] == '>') break;

            int nameStart = position;
            while (position < rootTag.Length && !char.IsWhiteSpace(rootTag[position]) &&
                   rootTag[position] != '=' && rootTag[position] != '>' && rootTag[position] != '/') position++;
            if (position == nameStart) return false;
            string attributeName = rootTag.Substring(nameStart, position - nameStart);
            try {
                XmlConvert.VerifyName(attributeName);
            } catch (XmlException) {
                return false;
            }
            if (!attributeNames.Add(attributeName)) return false;
            while (position < rootTag.Length && char.IsWhiteSpace(rootTag[position])) position++;
            if (position >= rootTag.Length || rootTag[position++] != '=') return false;
            while (position < rootTag.Length && char.IsWhiteSpace(rootTag[position])) position++;
            if (position >= rootTag.Length || rootTag[position] != '\'' && rootTag[position] != '"') return false;
            char quote = rootTag[position++];
            int valueStart = position;
            while (position < rootTag.Length && rootTag[position] != quote) {
                if (rootTag[position] == '<') return false;
                position++;
            }
            if (position >= rootTag.Length) return false;
            string attributeValue = rootTag.Substring(valueStart, position - valueStart);
            position++;
            if (!string.Equals(attributeName, namespaceAttribute, StringComparison.Ordinal)) continue;
            if (foundNamespace) return false;
            if (!TryDecodeXmlAttributeValue(attributeValue, out namespaceUri)) {
                return TryResolveXmlRootNamespaceWithInternalEntities(
                    qualifiedName, rootTag, documentType, out namespaceUri);
            }
            foundNamespace = true;
        }

        return separator < 0 || foundNamespace && namespaceUri.Length > 0;
    }

    private static bool TryResolveXmlRootNamespaceWithInternalEntities(
        string qualifiedName,
        string rootTag,
        string documentType,
        out string namespaceUri) {
        namespaceUri = string.Empty;
        if (documentType.Length == 0) return false;
        string closingTag = rootTag.EndsWith("/>", StringComparison.Ordinal)
            ? string.Empty
            : "</" + qualifiedName + ">";
        string probe = documentType + rootTag + closingTag;
        var settings = new XmlReaderSettings {
            DtdProcessing = DtdProcessing.Parse,
            XmlResolver = null,
            MaxCharactersFromEntities = 4096,
            MaxCharactersInDocument = Math.Max(4096, probe.Length + 4096)
        };
        try {
            using var input = new StringReader(probe);
            using XmlReader reader = XmlReader.Create(input, settings);
            while (reader.Read()) {
                if (reader.NodeType != XmlNodeType.Element || reader.Depth != 0) continue;
                if (!string.Equals(reader.Name, qualifiedName, StringComparison.Ordinal)) return false;
                namespaceUri = reader.NamespaceURI;
                return qualifiedName.IndexOf(':') < 0 || namespaceUri.Length > 0;
            }
        } catch (XmlException) {
            return false;
        }
        return false;
    }

    private static bool TryDecodeXmlAttributeValue(string value, out string decoded) {
        var result = new StringBuilder(value.Length);
        for (int position = 0; position < value.Length; position++) {
            char character = value[position];
            if (character != '&') {
                result.Append(character);
                continue;
            }
            int end = value.IndexOf(';', position + 1);
            if (end < 0) {
                decoded = string.Empty;
                return false;
            }
            string reference = value.Substring(position + 1, end - position - 1);
            switch (reference) {
                case "amp": result.Append('&'); break;
                case "apos": result.Append('\''); break;
                case "gt": result.Append('>'); break;
                case "lt": result.Append('<'); break;
                case "quot": result.Append('"'); break;
                default:
                    int radix = reference.StartsWith("#x", StringComparison.Ordinal) ? 16 : 10;
                    int digits = reference.StartsWith("#x", StringComparison.Ordinal) ? 2 : 1;
                    if (!reference.StartsWith("#", StringComparison.Ordinal) || reference.Length <= digits ||
                        !int.TryParse(reference.Substring(digits),
                            radix == 16 ? System.Globalization.NumberStyles.AllowHexSpecifier : System.Globalization.NumberStyles.None,
                            System.Globalization.CultureInfo.InvariantCulture, out int codePoint) ||
                        !IsXmlCodePoint(codePoint)) {
                        decoded = string.Empty;
                        return false;
                    }
                    result.Append(char.ConvertFromUtf32(codePoint));
                    break;
            }
            position = end;
        }
        decoded = result.ToString();
        return true;
    }

    private static bool IsXmlCodePoint(int value) =>
        value == 0x9 || value == 0xA || value == 0xD ||
        value >= 0x20 && value <= 0xD7FF ||
        value >= 0xE000 && value <= 0xFFFD ||
        value >= 0x10000 && value <= 0x10FFFF;
}
