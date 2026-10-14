namespace OfficeIMO.OpenDocument;

internal static partial class OdfTextCodec {
    internal const int MaximumDecodedCharacters = 16 * 1024 * 1024;

    internal static string Read(XElement element) {
        if (element == null) throw new ArgumentNullException(nameof(element));
        XNode? first = element.FirstNode;
        if (first == null) return string.Empty;
        if (first is XText text && first.NextNode == null && !text.Value.Any(IsXmlSpace)) {
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
        return DecodeNodes(nodes, MaximumDecodedCharacters).Text;
    }

    internal static string Read(XElement element, ref int remainingCharacters) {
        return ReadNodes(element.Nodes(), ref remainingCharacters);
    }

    internal static string ReadNodes(IEnumerable<XNode> nodes, ref int remainingCharacters) {
        TextSnapshot snapshot = DecodeNodes(nodes, remainingCharacters);
        remainingCharacters -= snapshot.SourceCharacters;
        return snapshot.Text;
    }

    /// <summary>Decodes inline fallback syntax while charging the caller's story-wide node budget.</summary>
    internal static string ReadNodes(IEnumerable<XNode> nodes, ref int remainingCharacters, ref int visitedNodes) {
        if (nodes == null) throw new ArgumentNullException(nameof(nodes));
        TextSnapshot snapshot;
        int visited = visitedNodes;
        try {
            snapshot = DecodeNodes(nodes, remainingCharacters, () => {
                if (++visited > OdfTextTraversal.MaximumVisitedElements)
                    throw new NotSupportedException("Inline text traversal exceeds the projection node limit.");
            });
        } finally {
            visitedNodes = visited;
        }
        remainingCharacters -= snapshot.SourceCharacters;
        return snapshot.Text;
    }

    internal static string ReadJoined(IEnumerable<XElement> elements) {
        if (elements == null) throw new ArgumentNullException(nameof(elements));
        var builder = new StringBuilder();
        bool first = true;
        int remaining = MaximumDecodedCharacters;
        foreach (XElement element in elements) {
            if (element == null) throw new ArgumentException("Text elements cannot contain null entries.", nameof(elements));
            if (!first) {
                EnsureCapacity(builder, 1, MaximumDecodedCharacters);
                if (remaining == 0) throw new InvalidDataException("Decoded OpenDocument text exceeds the character safety limit.");
                builder.Append('\n');
                remaining--;
            }
            first = false;
            builder.Append(Read(element, ref remaining));
        }
        return builder.ToString();
    }

    internal static string JoinBounded(IEnumerable<string> values, string separator = "\n") {
        if (values == null) throw new ArgumentNullException(nameof(values));
        if (separator == null) throw new ArgumentNullException(nameof(separator));
        var builder = new StringBuilder();
        bool first = true;
        foreach (string value in values) {
            if (value == null) throw new ArgumentException("Text values cannot contain null entries.", nameof(values));
            if (!first) {
                EnsureCapacity(builder, separator.Length, MaximumDecodedCharacters);
                builder.Append(separator);
            }
            EnsureCapacity(builder, value.Length, MaximumDecodedCharacters);
            builder.Append(value);
            first = false;
        }
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

        string source = ReadTransformText(element);
        string transformed = OfficeIMO.Drawing.OfficeTextCaseTransformer.Apply(source, textCase, culture);
        if (transformed.Length == source.Length) {
            int offset = 0;
            AssignTransformedText(element.Nodes(), transformed, ref offset);
            return;
        }

        AssignVariableLengthTransformedText(new[] { element }, textCase, culture);
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
            string paragraphText = ReadTransformText(elements[index]);
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

        AssignVariableLengthTransformedText(elements, textCase, culture);
    }

    private static void AssignVariableLengthTransformedText(
        IReadOnlyList<XElement> elements,
        OfficeIMO.Drawing.OfficeTextCase textCase,
        CultureInfo? culture) {
        var segments = new List<string>();
        var targets = new List<XText?>();
        for (int index = 0; index < elements.Count; index++) {
            if (index > 0) {
                segments.Add("\n");
                targets.Add(null);
            }
            CollectTransformSegments(elements[index].Nodes(), segments, targets);
        }

        IReadOnlyList<string> transformed = OfficeIMO.Drawing.OfficeTextCaseTransformer.ApplySegments(segments, textCase, culture);
        for (int index = 0; index < targets.Count; index++) {
            if (targets[index] != null) targets[index]!.Value = transformed[index];
        }
    }

    private static void CollectTransformSegments(
        IEnumerable<XNode> nodes,
        IList<string> segments,
        IList<XText?> targets) {
        foreach (XNode node in VisibleNodes(nodes)) {
            if (node is XText text) {
                segments.Add(text.Value);
                targets.Add(text);
                continue;
            }
            if (!(node is XElement element)) continue;
            if (element.Name == OdfNamespaces.Text + "s") {
                segments.Add(new string(' ', ParsePositiveCount((string?)element.Attribute(OdfNamespaces.Text + "c"))));
                targets.Add(null);
            } else if (element.Name == OdfNamespaces.Text + "tab") {
                segments.Add("\t");
                targets.Add(null);
            } else if (element.Name == OdfNamespaces.Text + "line-break") {
                segments.Add("\n");
                targets.Add(null);
            }
        }
    }

    private static void AssignTransformedText(IEnumerable<XNode> nodes, string transformed, ref int offset) {
        foreach (XNode node in VisibleNodes(nodes)) {
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
            }
        }
    }

    private static string ReadTransformText(XElement element) {
        var builder = new StringBuilder();
        AppendTransformValue(element.Nodes(), builder, MaximumDecodedCharacters);
        return builder.ToString();
    }

    private static void AppendTransformValue(IEnumerable<XNode> nodes, StringBuilder builder, int maximumCharacters) {
        foreach (XNode node in VisibleNodes(nodes)) {
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
            }
        }
    }

    internal static bool IsNonVisibleTextElement(XElement element) =>
        element.Name == OdfNamespaces.Office + "annotation" ||
        element.Name == OdfNamespaces.Presentation + "notes" ||
        element.Name == OdfNamespaces.Text + "note" ||
        element.Name == OdfNamespaces.Text + "ruby-text" ||
        element.Name == OdfNamespaces.Svg + "title" ||
        element.Name == OdfNamespaces.Svg + "desc" ||
        element.Name == OdfNamespaces.Office + "binary-data" ||
        element.Name == OdfNamespaces.Draw + "object" ||
        element.Name == OdfNamespaces.Draw + "object-ole" ||
        element.Name == OdfNamespaces.Draw + "image" ||
        element.Name == OdfNamespaces.Draw + "plugin" ||
        element.Name == OdfNamespaces.Draw + "applet" ||
        element.Name == OdfNamespaces.Draw + "floating-frame";

    /// <summary>
    /// Enumerates visible text and whitespace tokens without entering another text story.
    /// Bounds apply before a container is descended or a text transform can begin assigning values.
    /// </summary>
    private static IEnumerable<XNode> VisibleNodes(IEnumerable<XNode> nodes, Action? onVisit = null, Func<XElement, bool>? atomic = null) {
        var pending = new Stack<IEnumerator<XNode>>();
        pending.Push(nodes.GetEnumerator());
        int visited = 0;
        try {
            while (pending.Count > 0) {
                IEnumerator<XNode> current = pending.Peek();
                if (!current.MoveNext()) {
                    pending.Pop().Dispose();
                    continue;
                }
                if (++visited > OdfTextTraversal.MaximumVisitedElements)
                    throw new NotSupportedException($"OpenDocument text decoding exceeds the {OdfTextTraversal.MaximumVisitedElements}-node safety limit.");
                onVisit?.Invoke();

                XNode node = current.Current;
                if (node is XElement element) {
                    if (IsNonVisibleTextElement(element)) continue;
                    if (atomic?.Invoke(element) == true) {
                        if (pending.Count > OdfTextTraversal.MaximumContainerDepth)
                            throw new NotSupportedException($"OpenDocument text decoding exceeds the {OdfTextTraversal.MaximumContainerDepth}-container nesting limit.");
                        // Scalar projections still charge every source node, including empty
                        // cache text and comments; atomic rendering must not bypass traversal limits.
                        foreach (XNode child in element.Nodes()) {
                            if (++visited > OdfTextTraversal.MaximumVisitedElements)
                                throw new NotSupportedException("OpenDocument scalar text exceeds the node safety limit.");
                            onVisit?.Invoke();
                        }
                        yield return node;
                        continue;
                    }
                    if (element.Name != OdfNamespaces.Text + "s" && element.Name != OdfNamespaces.Text + "tab" &&
                        element.Name != OdfNamespaces.Text + "line-break") {
                        if (pending.Count > OdfTextTraversal.MaximumContainerDepth)
                            throw new NotSupportedException($"OpenDocument text decoding exceeds the {OdfTextTraversal.MaximumContainerDepth}-container nesting limit.");
                        pending.Push(element.Nodes().GetEnumerator());
                        continue;
                    }
                }
                yield return node;
            }
        } finally {
            while (pending.Count > 0) pending.Pop().Dispose();
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
