namespace OfficeIMO.Word;

/// <summary>
/// Normalizes HTML list markup before it is handed to Word's AltChunk importer.
/// </summary>
/// <remarks>
/// Word interprets a <c>value</c> attribute on an item in an unordered list as
/// numbering metadata. Browsers ignore that attribute for bullet semantics, so
/// remove it only when the nearest containing list is a <c>ul</c>. Ordered-list
/// values remain intact.
/// </remarks>
internal static class HtmlListSemanticsNormalizer {
    internal static string Normalize(string html) {
        if (string.IsNullOrEmpty(html)
            || html.IndexOf("<li", StringComparison.OrdinalIgnoreCase) < 0
            || html.IndexOf("<ul", StringComparison.OrdinalIgnoreCase) < 0) {
            return html;
        }

        var output = new StringBuilder(html.Length);
        var listAncestors = new List<string>();
        string? rawTextElement = null;
        int position = 0;

        while (position < html.Length) {
            if (rawTextElement != null) {
                int rawEnd = html.IndexOf("</" + rawTextElement, position, StringComparison.OrdinalIgnoreCase);
                if (rawEnd < 0) {
                    output.Append(html, position, html.Length - position);
                    break;
                }

                output.Append(html, position, rawEnd - position);
                position = rawEnd;
                rawTextElement = null;
            }

            int tagStart = html.IndexOf('<', position);
            if (tagStart < 0) {
                output.Append(html, position, html.Length - position);
                break;
            }

            output.Append(html, position, tagStart - position);

            if (html.IndexOf("<!--", tagStart, StringComparison.Ordinal) == tagStart) {
                int commentEnd = html.IndexOf("-->", tagStart + 4, StringComparison.Ordinal);
                if (commentEnd < 0) {
                    output.Append(html, tagStart, html.Length - tagStart);
                    break;
                }

                int commentLength = commentEnd + 3 - tagStart;
                output.Append(html, tagStart, commentLength);
                position = commentEnd + 3;
                continue;
            }

            int tagEnd = FindTagEnd(html, tagStart);
            if (tagEnd < 0) {
                output.Append(html, tagStart, html.Length - tagStart);
                break;
            }

            string tag = html.Substring(tagStart, tagEnd - tagStart + 1);
            if (TryReadTag(tag, out string tagName, out bool closing, out bool selfClosing)) {
                if (!closing
                    && string.Equals(tagName, "li", StringComparison.OrdinalIgnoreCase)
                    && listAncestors.Count > 0
                    && string.Equals(listAncestors[listAncestors.Count - 1], "ul", StringComparison.OrdinalIgnoreCase)) {
                    tag = RemoveValueAttributes(tag);
                }

                if (string.Equals(tagName, "ul", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(tagName, "ol", StringComparison.OrdinalIgnoreCase)) {
                    if (closing) {
                        RemoveClosedList(listAncestors, tagName);
                    } else if (!selfClosing) {
                        listAncestors.Add(tagName.ToLowerInvariant());
                    }
                } else if (!closing && !selfClosing
                    && (string.Equals(tagName, "script", StringComparison.OrdinalIgnoreCase)
                        || string.Equals(tagName, "style", StringComparison.OrdinalIgnoreCase))) {
                    rawTextElement = tagName;
                }
            }

            output.Append(tag);
            position = tagEnd + 1;
        }

        return output.ToString();
    }

    private static int FindTagEnd(string html, int tagStart) {
        char quote = '\0';
        for (int i = tagStart + 1; i < html.Length; i++) {
            char current = html[i];
            if (quote != '\0') {
                if (current == quote) {
                    quote = '\0';
                }
                continue;
            }

            if (current == '\'' || current == '"') {
                quote = current;
            } else if (current == '>') {
                return i;
            }
        }

        return -1;
    }

    private static bool TryReadTag(string tag, out string name, out bool closing, out bool selfClosing) {
        name = string.Empty;
        closing = false;
        selfClosing = false;

        int position = 1;
        while (position < tag.Length && char.IsWhiteSpace(tag[position])) {
            position++;
        }

        if (position >= tag.Length || tag[position] == '!' || tag[position] == '?') {
            return false;
        }

        if (tag[position] == '/') {
            closing = true;
            position++;
            while (position < tag.Length && char.IsWhiteSpace(tag[position])) {
                position++;
            }
        }

        int nameStart = position;
        while (position < tag.Length && IsNameCharacter(tag[position])) {
            position++;
        }

        if (position == nameStart) {
            return false;
        }

        name = tag.Substring(nameStart, position - nameStart);
        int tail = tag.Length - 2;
        while (tail >= 0 && char.IsWhiteSpace(tag[tail])) {
            tail--;
        }
        selfClosing = tail >= 0 && tag[tail] == '/';
        return true;
    }

    private static bool IsNameCharacter(char value) =>
        char.IsLetterOrDigit(value) || value == '-' || value == ':' || value == '_';

    private static void RemoveClosedList(List<string> listAncestors, string tagName) {
        for (int i = listAncestors.Count - 1; i >= 0; i--) {
            if (!string.Equals(listAncestors[i], tagName, StringComparison.OrdinalIgnoreCase)) {
                continue;
            }

            listAncestors.RemoveRange(i, listAncestors.Count - i);
            return;
        }
    }

    private static string RemoveValueAttributes(string tag) {
        var ranges = new List<Tuple<int, int>>();
        int position = 1;
        while (position < tag.Length && char.IsWhiteSpace(tag[position])) {
            position++;
        }
        while (position < tag.Length && IsNameCharacter(tag[position])) {
            position++;
        }

        int tagContentEnd = tag.Length - 1;
        while (position < tagContentEnd) {
            int whitespaceStart = position;
            while (position < tagContentEnd && char.IsWhiteSpace(tag[position])) {
                position++;
            }

            if (position >= tagContentEnd || tag[position] == '/') {
                break;
            }

            int attributeStart = position;
            while (position < tagContentEnd
                && !char.IsWhiteSpace(tag[position])
                && tag[position] != '='
                && tag[position] != '/'
                && tag[position] != '>') {
                position++;
            }

            if (position == attributeStart) {
                position++;
                continue;
            }

            string attributeName = tag.Substring(attributeStart, position - attributeStart);
            int attributeNameEnd = position;
            while (position < tagContentEnd && char.IsWhiteSpace(tag[position])) {
                position++;
            }

            if (position < tagContentEnd && tag[position] == '=') {
                position++;
                while (position < tagContentEnd && char.IsWhiteSpace(tag[position])) {
                    position++;
                }

                if (position < tagContentEnd && (tag[position] == '\'' || tag[position] == '"')) {
                    char quote = tag[position++];
                    while (position < tagContentEnd && tag[position] != quote) {
                        position++;
                    }
                    if (position < tagContentEnd) {
                        position++;
                    }
                } else {
                    while (position < tagContentEnd
                        && !char.IsWhiteSpace(tag[position])
                        && tag[position] != '/'
                        && tag[position] != '>') {
                        position++;
                    }
                }
            } else {
                position = attributeNameEnd;
            }

            if (string.Equals(attributeName, "value", StringComparison.OrdinalIgnoreCase)) {
                ranges.Add(Tuple.Create(whitespaceStart, position));
            }
        }

        if (ranges.Count == 0) {
            return tag;
        }

        var normalized = new StringBuilder(tag.Length);
        int copiedThrough = 0;
        foreach (var range in ranges) {
            normalized.Append(tag, copiedThrough, range.Item1 - copiedThrough);
            copiedThrough = range.Item2;
        }
        normalized.Append(tag, copiedThrough, tag.Length - copiedThrough);
        return normalized.ToString();
    }
}
