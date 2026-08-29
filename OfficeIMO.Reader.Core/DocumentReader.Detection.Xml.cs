using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace OfficeIMO.Reader;

internal static partial class DocumentReaderEngine {
    private static bool TryResolveXmlRootNamespace(string qualifiedName, string rootTag, out string namespaceUri) {
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
            if (foundNamespace || !TryDecodeXmlAttributeValue(attributeValue, out namespaceUri)) return false;
            foundNamespace = true;
        }

        return separator < 0 || foundNamespace && namespaceUri.Length > 0;
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
