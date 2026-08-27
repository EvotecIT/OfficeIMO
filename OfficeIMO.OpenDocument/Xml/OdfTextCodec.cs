namespace OfficeIMO.OpenDocument;

internal static class OdfTextCodec {
    private const int MaximumDecodedCharacters = 16 * 1024 * 1024;

    internal static string Read(XElement element) {
        if (element == null) throw new ArgumentNullException(nameof(element));
        XNode? first = element.FirstNode;
        if (first == null) return string.Empty;
        if (first is XText text && first.NextNode == null) {
            string value = text.Value;
            if (value.Length > MaximumDecodedCharacters) {
                throw new InvalidDataException($"Decoded OpenDocument text exceeds the {MaximumDecodedCharacters}-character safety limit.");
            }
            return value;
        }
        return ReadNodes(element.Nodes());
    }

    internal static string ReadNodes(IEnumerable<XNode> nodes) {
        if (nodes == null) throw new ArgumentNullException(nameof(nodes));
        var builder = new StringBuilder();
        AppendValue(nodes, builder, MaximumDecodedCharacters);
        return builder.ToString();
    }

    internal static void Replace(XElement element, string? text) {
        if (element == null) throw new ArgumentNullException(nameof(element));
        element.RemoveNodes();
        Append(element, text);
    }

    internal static void Append(XElement element, string? text) {
        if (element == null) throw new ArgumentNullException(nameof(element));
        if (string.IsNullOrEmpty(text)) return;

        var plain = new StringBuilder();
        int spaces = 0;
        Action flushPlain = () => {
            if (plain.Length == 0) return;
            element.Add(new XText(plain.ToString()));
            plain.Clear();
        };
        Action flushSpaces = () => {
            if (spaces == 0) return;
            flushPlain();
            var space = new XElement(OdfNamespaces.Text + "s");
            if (spaces != 1) space.SetAttributeValue(OdfNamespaces.Text + "c", spaces);
            element.Add(space);
            spaces = 0;
        };

        foreach (char character in text!) {
            if (character == ' ') {
                spaces++;
                continue;
            }
            flushSpaces();
            if (character == '\t') {
                flushPlain();
                element.Add(new XElement(OdfNamespaces.Text + "tab"));
            } else if (character == '\n') {
                flushPlain();
                element.Add(new XElement(OdfNamespaces.Text + "line-break"));
            } else if (character != '\r') {
                plain.Append(character);
            }
        }
        flushSpaces();
        flushPlain();
    }

    internal static void TransformTextCase(
        XElement element,
        OfficeIMO.Drawing.OfficeTextCase textCase,
        CultureInfo? culture = null) {
        if (element == null) throw new ArgumentNullException(nameof(element));
        if (textCase == OfficeIMO.Drawing.OfficeTextCase.None) return;

        string source = Read(element);
        string transformed = OfficeIMO.Drawing.OfficeTextCaseTransformer.Apply(source, textCase, culture);
        if (transformed.Length == source.Length) {
            int offset = 0;
            AssignTransformedText(element.Nodes(), transformed, ref offset);
            return;
        }

        // Culture-sensitive casing can occasionally change UTF-16 length. Preserve every
        // inline node in that uncommon case, even though its casing context is node-local.
        foreach (XText text in element.DescendantNodes().OfType<XText>().ToList()) {
            text.Value = OfficeIMO.Drawing.OfficeTextCaseTransformer.Apply(text.Value, textCase, culture);
        }
    }

    internal static void TransformTextCase(
        IReadOnlyList<XElement> elements,
        OfficeIMO.Drawing.OfficeTextCase textCase,
        CultureInfo? culture = null) {
        if (elements == null) throw new ArgumentNullException(nameof(elements));
        if (textCase == OfficeIMO.Drawing.OfficeTextCase.None || elements.Count == 0) return;

        var source = new StringBuilder();
        for (int index = 0; index < elements.Count; index++) {
            if (elements[index] == null) throw new ArgumentException("Text elements cannot contain null entries.", nameof(elements));
            string paragraphText = Read(elements[index]);
            int separatorLength = index > 0 ? 1 : 0;
            EnsureCapacity(source, separatorLength + paragraphText.Length, MaximumDecodedCharacters);
            if (separatorLength != 0) source.Append('\n');
            source.Append(paragraphText);
        }

        string transformed = OfficeIMO.Drawing.OfficeTextCaseTransformer.Apply(source.ToString(), textCase, culture);
        if (transformed.Length == source.Length) {
            int offset = 0;
            for (int index = 0; index < elements.Count; index++) {
                AssignTransformedText(elements[index].Nodes(), transformed, ref offset);
                if (index < elements.Count - 1) offset++;
            }
            return;
        }

        // Preserve each paragraph's inline nodes when culture-sensitive casing changes UTF-16 length.
        // Sentence and word context remains shared in the common length-preserving path above.
        foreach (XElement element in elements) TransformTextCase(element, textCase, culture);
    }

    private static void AssignTransformedText(IEnumerable<XNode> nodes, string transformed, ref int offset) {
        foreach (XNode node in nodes) {
            if (node is XText text) {
                int length = text.Value.Length;
                text.Value = transformed.Substring(offset, length);
                offset += length;
                continue;
            }
            if (!(node is XElement element)) continue;
            if (element.Name == OdfNamespaces.Text + "s") {
                offset += ParsePositiveCount((string?)element.Attribute(OdfNamespaces.Text + "c"));
            } else if (element.Name == OdfNamespaces.Text + "tab" ||
                       element.Name == OdfNamespaces.Text + "line-break") {
                offset++;
            } else {
                AssignTransformedText(element.Nodes(), transformed, ref offset);
            }
        }
    }

    private static void AppendValue(IEnumerable<XNode> nodes, StringBuilder builder, int maximumCharacters) {
        foreach (XNode node in nodes) {
            if (node is XText text) {
                AppendBounded(builder, text.Value, maximumCharacters);
                continue;
            }
            if (!(node is XElement element)) continue;
            if (element.Name == OdfNamespaces.Text + "s") {
                int count = ParsePositiveCount((string?)element.Attribute(OdfNamespaces.Text + "c"));
                EnsureCapacity(builder, count, maximumCharacters);
                builder.Append(' ', count);
            } else if (element.Name == OdfNamespaces.Text + "tab") {
                EnsureCapacity(builder, 1, maximumCharacters);
                builder.Append('\t');
            } else if (element.Name == OdfNamespaces.Text + "line-break") {
                EnsureCapacity(builder, 1, maximumCharacters);
                builder.Append('\n');
            } else {
                AppendValue(element.Nodes(), builder, maximumCharacters);
            }
        }
    }

    private static void AppendBounded(StringBuilder builder, string value, int maximumCharacters) {
        EnsureCapacity(builder, value.Length, maximumCharacters);
        builder.Append(value);
    }

    private static void EnsureCapacity(StringBuilder builder, int additionalCharacters, int maximumCharacters) {
        if (additionalCharacters > maximumCharacters - builder.Length) {
            throw new InvalidDataException($"Decoded OpenDocument text exceeds the {maximumCharacters}-character safety limit.");
        }
    }

    private static int ParsePositiveCount(string? value) {
        return int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int count) && count > 0 ? count : 1;
    }
}
