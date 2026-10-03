using System.Globalization;
using System.Threading;

namespace OfficeIMO.AsciiDoc;

/// <summary>Derives section targets locally from the supported visible title semantics.</summary>
internal static class AsciiDocSectionIdentifier {
    internal static string Title(AsciiDocInlineSequence inlines, AsciiDocDocumentAttributes attributes, CancellationToken token, out bool approximate) {
        var output = new StringBuilder();
        bool simplified = false;
        Append(inlines, 0);
        approximate = simplified;
        return output.ToString();

        void Append(AsciiDocInlineSequence sequence, int depth) {
            if (depth >= 128) throw new InvalidDataException("Section title traversal exceeds the inline nesting limit.");
            foreach (AsciiDocInline inline in sequence.Items) {
                token.ThrowIfCancellationRequested();
                switch (inline) {
                    case AsciiDocTextInline text: output.Append(text.Text.Replace("<", "&lt;").Replace(">", "&gt;")); break;
                    case AsciiDocFormattedInline formatted: Append(formatted.Content, depth + 1); break;
                    case AsciiDocAttributeReferenceInline reference:
                        var substitution = AsciiDocAttributeSubstitutor.Substitute("{" + reference.Name + "}", attributes);
                        simplified |= substitution.Diagnostics.Count > 0;
                        if (substitution.Value == "{" + reference.Name + "}") { output.Append(substitution.Value); break; }
                        Append(AsciiDocInlineSequence.Parse(substitution.Value, cancellationToken: token), depth + 1);
                        break;
                    case AsciiDocAnchorInline _: break;
                    case AsciiDocMacroInline macro when macro.Name == "http" || macro.Name == "https" || macro.Name == "ftp" || macro.Name == "link":
                        string label = macro.Attributes.Entries.FirstOrDefault(entry => entry.Kind == AsciiDocElementAttributeKind.Positional)?.Value ??
                            ((macro.Name == "http" || macro.Name == "https" || macro.Name == "ftp") ? macro.Name + ":" + macro.Target : macro.Target);
                        var expanded = AsciiDocAttributeSubstitutor.Substitute(label, attributes);
                        simplified |= expanded.Diagnostics.Count > 0;
                        Append(AsciiDocInlineSequence.Parse(expanded.Value, cancellationToken: token), depth + 1);
                        break;
                    case AsciiDocPassthroughInline passthrough: output.Append(passthrough.Content); simplified = true; break;
                    case AsciiDocMacroInline macro when macro.Name == "pass": output.Append(macro.AttributeList); simplified = true; break;
                    case AsciiDocCrossReferenceInline reference: output.Append(reference.Text ?? reference.Target); simplified = true; break;
                    default: output.Append(inline.OriginalText); simplified = true; break;
                }
            }
        }
    }

    internal static string Label(string title, CancellationToken token) {
        var output = new StringBuilder(title.Length);
        bool noClosingTag = false;
        for (int index = 0; index < title.Length; index++) {
            if ((index & 4095) == 0) token.ThrowIfCancellationRequested();
            if (title[index] == '<' && !noClosingTag) {
                int end = title.IndexOf('>', index + 1);
                if (end >= 0) { index = end; continue; }
                noClosingTag = true;
            }
            output.Append(title[index]);
        }
        return System.Net.WebUtility.HtmlDecode(output.ToString());
    }

    internal static string Create(string title, string prefix, string separator, CancellationToken token) {
        string lower = title.ToLowerInvariant();
        var cleaned = new StringBuilder(lower.Length);
        bool noClosingTag = false;
        for (int index = 0; index < lower.Length; index++) {
            if ((index & 4095) == 0) token.ThrowIfCancellationRequested();
            char character = lower[index];
            if (character == '<' && !noClosingTag) {
                int end = lower.IndexOf('>', index + 1);
                if (end >= 0) { index = end; continue; }
                noClosingTag = true;
            }
            if (character == '&') {
                int end = index + 1;
                while (end < lower.Length && (char.IsLetterOrDigit(lower[end]) || lower[end] == '#')) {
                    if ((end & 4095) == 0) token.ThrowIfCancellationRequested();
                    end++;
                }
                if (end > index + 1 && end < lower.Length && lower[end] == ';') { index = end; continue; }
            }
            if (character == '.' && index + 2 < lower.Length && lower[index + 1] == '.' && lower[index + 2] == '.' &&
                (index == 0 || lower[index - 1] != '\\')) {
                index += 2;
                continue;
            }
            if (character == '-' && index > 0 && index + 2 < lower.Length && lower[index + 1] == '-' &&
                char.IsLetterOrDigit(lower[index - 1]) && char.IsLetterOrDigit(lower[index + 2])) {
                index++;
                continue;
            }
            // Asciidoctor's spaced em dash replaces its surrounding spaces with
            // character references, which section ID generation removes too.
            if (character == '-' && index > 0 && lower[index - 1] == ' ' && index + 2 < lower.Length && lower[index + 1] == '-' && lower[index + 2] == ' ') {
                if (cleaned.Length > 0 && cleaned[cleaned.Length - 1] == ' ') cleaned.Length--;
                index += 2;
                continue;
            }
            if (character == ' ' || character == '-' || character == '.') { cleaned.Append(character); continue; }
            UnicodeCategory category = CharUnicodeInfo.GetUnicodeCategory(lower, index);
            if (category == UnicodeCategory.UppercaseLetter || category == UnicodeCategory.LowercaseLetter ||
                category == UnicodeCategory.TitlecaseLetter || category == UnicodeCategory.ModifierLetter ||
                category == UnicodeCategory.OtherLetter || category == UnicodeCategory.DecimalDigitNumber ||
                category == UnicodeCategory.NonSpacingMark || category == UnicodeCategory.SpacingCombiningMark ||
                category == UnicodeCategory.ConnectorPunctuation) {
                cleaned.Append(character);
                if (char.IsHighSurrogate(character) && index + 1 < lower.Length && char.IsLowSurrogate(lower[index + 1])) cleaned.Append(lower[++index]);
            }
        }
        string value = prefix + cleaned;
        if (separator.Length == 0) return value.Replace(" ", string.Empty);
        var result = new StringBuilder(value.Length);
        bool precedingSeparator = false;
        for (int index = 0; index < value.Length; index++) {
            if ((index & 4095) == 0) token.ThrowIfCancellationRequested();
            bool isSeparator = value[index] == ' ' || value[index] == '-' || value[index] == '.' ||
                string.CompareOrdinal(value, index, separator, 0, separator.Length) == 0;
            if (isSeparator) {
                if (!precedingSeparator) result.Append(separator);
                precedingSeparator = true;
                if (separator.Length > 1 && string.CompareOrdinal(value, index, separator, 0, separator.Length) == 0) index += separator.Length - 1;
            } else { result.Append(value[index]); precedingSeparator = false; }
        }
        if (precedingSeparator && result.Length > 0) result.Length -= separator.Length;
        if (prefix.Length == 0 && result.Length >= separator.Length && string.CompareOrdinal(result.ToString(), 0, separator, 0, separator.Length) == 0)
            result.Remove(0, separator.Length);
        return result.ToString();
    }
}
