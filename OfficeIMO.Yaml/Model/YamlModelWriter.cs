namespace OfficeIMO.Yaml;

/// <summary>
/// Writes model nodes through <see cref="YamlEmitter"/>, the same core <see cref="YamlWriter"/> uses: anchors, tags,
/// styles, comments and blank lines of new nodes, and comments read from the source when a whole document is reformatted.
/// </summary>
internal static class YamlModelWriter {
    /// <summary>Renders a node on its own, as saving would write it at the start of a line.</summary>
    public static string Render(YamlNode node) {
        var text = new StringWriter(CultureInfo.InvariantCulture);
        var emitter = new YamlEmitter(text, Options(node.Document));
        emitter.BeginSlot(YamlSlot.Inline, 0, 0);
        WriteNode(emitter, node, ownComments: false);
        emitter.EndFragment(endLine: false);
        emitter.Release();
        return TrimFinalBreak(text.ToString());
    }

    /// <summary>The scalar's text as written, without quotes or a block header, matching what <see cref="YamlScalar.RawText"/> reports for loaded scalars.</summary>
    public static string RenderRawText(YamlScalar scalar) {
        string text = Render(scalar);
        int start = 0;
        if (scalar._properties is { } properties && (properties.Anchor is not null || properties.Tag is not null)) {
            // Skip the anchor and tag in front of the content.
            while (start < text.Length && (text[start] == '&' || text[start] == '!')) {
                while (start < text.Length && text[start] != ' ' && text[start] != '\n') start++;
                while (start < text.Length && text[start] == ' ') start++;
            }
        }
        text = text.Substring(start);
        if (text.Length >= 2 && (text[0] == '\'' || text[0] == '"')) return text.Substring(1, text.Length - 2);
        if (text.Length > 0 && (text[0] == '|' || text[0] == '>')) {
            int line = text.IndexOf('\n');
            return line < 0 ? string.Empty : text.Substring(line + 1);
        }
        return text;
    }

    public static YamlWriterOptions Options(YamlDocument? document) => document is null ? YamlWriterOptions.Default : document.WriterOptions;

    /// <summary>
    /// Writes a node and everything below it. <paramref name="ownComments"/> writes the node's own trailing comment;
    /// comment lines above children are always written where the context allows them.
    /// </summary>
    public static void WriteNode(YamlEmitter emitter, YamlNode node, bool ownComments) {
        bool redefined = node._part?.RedefinesSecondaryHandle ?? false;
        if (node is not YamlAlias) {
            if (node.Anchor is { } anchor) emitter.SetAnchor(anchor);
            if (node.Tag is { } tag) emitter.SetTag(tag, redefined);
        }
        switch (node) {
            case YamlAlias alias:
                emitter.WriteAlias(alias.Name);
                if (ownComments) TrailingComment(emitter, node);
                break;
            case YamlScalar scalar: {
                string value = scalar.Value;
                YamlStyle quote = node._part?._document.Layout.Quote ?? YamlStyle.SingleQuoted;
                YamlStyle style = emitter.ChooseStyle(value, scalar._style, quote);
                bool block = style is YamlStyle.Literal or YamlStyle.Folded;
                if (ownComments && block) TrailingComment(emitter, node);
                emitter.WriteScalar(value, style);
                if (ownComments && !block) TrailingComment(emitter, node);
                break;
            }
            case YamlMapping mapping: {
                emitter.StartCollection(mapping: true, mapping._style);
                bool flow = mapping._style == YamlStyle.Flow || emitter.InFlow;
                for (int i = 0; i < mapping._count; i++) {
                    YamlNode key = mapping.Entries[i].Key, value = mapping.Entries[i].Value;
                    if (!flow) CommentLines(emitter, key);
                    WriteNode(emitter, key, ownComments: false);
                    // A key's comment sits on its line when the value starts below it.
                    if (!flow && value is YamlMapping or YamlSequence && value._style == YamlStyle.Block && (value is YamlMapping { _count: > 0 } || value is YamlSequence { _count: > 0 })) {
                        TrailingComment(emitter, key);
                    }
                    WriteNode(emitter, value, ownComments: !flow);
                }
                emitter.EndCollection(mapping: true);
                if (ownComments && flow) TrailingComment(emitter, node);
                break;
            }
            case YamlSequence sequence: {
                emitter.StartCollection(mapping: false, sequence._style);
                bool flow = sequence._style == YamlStyle.Flow || emitter.InFlow;
                for (int i = 0; i < sequence._count; i++) {
                    YamlNode item = sequence.Items[i];
                    if (!flow) CommentLines(emitter, item);
                    WriteNode(emitter, item, ownComments: !flow);
                }
                emitter.EndCollection(mapping: false);
                if (ownComments && flow) TrailingComment(emitter, node);
                break;
            }
        }
    }

    /// <summary>Writes the blank line and comment lines above an entry whose owner is <paramref name="node"/>.</summary>
    public static void CommentLines(YamlEmitter emitter, YamlNode node) {
        if (!ReferenceEquals(YamlTrivia.Owner(node), node)) return;
        if (node.SpaceBefore) emitter.WriteBlankLine();
        if (node.CommentBefore is { } text) emitter.WriteCommentLines(text);
    }

    private static void TrailingComment(YamlEmitter emitter, YamlNode node) {
        if (!ReferenceEquals(YamlTrivia.TrailingTarget(node), node)) return;
        if (node.Comment is { } text) emitter.SetTrailingComment(text);
    }

    /// <summary>Writes a whole document part: markers, header comments, the root and footer comments.</summary>
    public static void WritePart(YamlEmitter emitter, YamlDocumentPart part, bool marker) {
        if (part.CommentBefore is { } header) {
            emitter.WriteCommentLines(header);
            // Without a marker, a blank line keeps the header from becoming the first entry's comment.
            if (!marker && emitter.Documents == 0) emitter.WriteBlankLine();
        }
        emitter.StartDocument(marker);
        if (marker && part.Comment is { } comment) emitter.SetTrailingComment(comment);
        if (part.RootNode is not null) {
            CommentLines(emitter, part.RootNode);
            WriteNode(emitter, part.RootNode, ownComments: true);
        }
        if (part.CommentAfter is { } footer) emitter.WriteCommentLines(footer);
        if (part.HasEndMarker) emitter.EndDocument();
    }

    public static string TrimFinalBreak(string text) {
        if (text.EndsWith("\r\n", StringComparison.Ordinal)) return text.Substring(0, text.Length - 2);
        if (text.EndsWith("\n", StringComparison.Ordinal)) return text.Substring(0, text.Length - 1);
        return text;
    }
}
