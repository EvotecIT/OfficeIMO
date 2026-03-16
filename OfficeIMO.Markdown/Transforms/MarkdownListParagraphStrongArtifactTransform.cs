namespace OfficeIMO.Markdown;

/// <summary>
/// Repairs malformed strong-marker artifacts inside already-parsed list-item paragraphs.
/// </summary>
/// <remarks>
/// This transform is intended for recoverable list-item inline artifacts that do not require block-boundary repair.
/// It operates directly on the inline AST and only repairs node patterns that the reader already recovered into
/// list-item paragraph content.
/// </remarks>
public sealed class MarkdownListParagraphStrongArtifactTransform : IMarkdownDocumentTransform {
    /// <summary>
    /// Creates a transform with the specified normalization options.
    /// </summary>
    public MarkdownListParagraphStrongArtifactTransform(MarkdownInputNormalizationOptions options) {
        Options = options ?? throw new ArgumentNullException(nameof(options));
    }

    /// <summary>
    /// Normalization switches used by this transform.
    /// </summary>
    public MarkdownInputNormalizationOptions Options { get; }

    /// <inheritdoc />
    public MarkdownDoc Transform(MarkdownDoc document, MarkdownDocumentTransformContext context) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        if (context == null) {
            throw new ArgumentNullException(nameof(context));
        }

        MarkdownDocumentBlockRewriter.RewriteDocument(document, block => RewriteBlock(block));
        return document;
    }

    private IMarkdownBlock RewriteBlock(IMarkdownBlock block) {
        switch (block) {
            case OrderedListBlock ordered:
                NormalizeListItems(ordered.Items);
                break;
            case UnorderedListBlock unordered:
                NormalizeListItems(unordered.Items);
                break;
        }

        return block;
    }

    private void NormalizeListItems(IList<ListItem> items) {
        for (var i = 0; i < items.Count; i++) {
            var item = items[i];
            if (item == null) {
                continue;
            }

            NormalizeSequence(item.Content);
            for (var paragraphIndex = 0; paragraphIndex < item.AdditionalParagraphs.Count; paragraphIndex++) {
                NormalizeSequence(item.AdditionalParagraphs[paragraphIndex]);
            }

            if (item.Children.Count > 0) {
                NormalizeNestedBlocks(item.Children);
            }
        }
    }

    private void NormalizeNestedBlocks(IList<IMarkdownBlock> blocks) {
        for (var i = 0; i < blocks.Count; i++) {
            var block = blocks[i];
            if (block == null) {
                continue;
            }

            blocks[i] = RewriteBlock(block);
        }
    }

    private void NormalizeSequence(InlineSequence? sequence) {
        if (sequence == null || sequence.Nodes.Count == 0) {
            return;
        }

        bool changed = MarkdownReader.NormalizeInlineSequenceInPlace(sequence, Options);
        var working = new List<IMarkdownInline>(sequence.Nodes.Count);

        for (var i = 0; i < sequence.Nodes.Count; i++) {
            var node = sequence.Nodes[i];
            if (node == null) {
                continue;
            }

            if (node is IInlineContainerMarkdownInline container
                && container.NestedInlines != null
                && container.NestedInlines.Nodes.Count > 0) {
                NormalizeSequence(container.NestedInlines);
            }

            working.Add(node);
        }

        bool localChanged;
        do {
            localChanged = false;

            if (Options.NormalizeLooseStrongDelimiters && TryFlattenRepeatedStrongRuns(working)) {
                localChanged = true;
            }

            if (Options.NormalizeDanglingTrailingStrongListClosers && TryRewriteDanglingTrailingStrongToken(working)) {
                localChanged = true;
            }

            if (Options.NormalizeMetricValueStrongRuns) {
                if (TryRewriteOveropenedMetricStrongRuns(working)) {
                    localChanged = true;
                }

                if (TryRewriteAdjacentMetricStrongValue(working)) {
                    localChanged = true;
                }

                if (TryRewriteMissingTrailingStrongClose(working)) {
                    localChanged = true;
                }
            }

            changed |= localChanged;
        } while (localChanged);

        if (!changed) {
            return;
        }

        sequence.ReplaceItems(CoalesceAdjacentTextRuns(working));
        MarkdownReader.NormalizeInlineSequenceInPlace(sequence, Options);
    }

    private static bool TryFlattenRepeatedStrongRuns(List<IMarkdownInline> nodes) {
        bool changed = false;

        for (var i = 0; i < nodes.Count; i++) {
            if (nodes[i] is not BoldSequenceInline outer || outer.Inlines.Nodes.Count != 1) {
                continue;
            }

            switch (outer.Inlines.Nodes[0]) {
                case BoldSequenceInline nestedSequence:
                    nodes[i] = new BoldSequenceInline(nestedSequence.Inlines);
                    changed = true;
                    break;
                case BoldInline nestedBold:
                    nodes[i] = new BoldInline(nestedBold.Text);
                    changed = true;
                    break;
            }
        }

        return changed;
    }

    private static bool TryRewriteDanglingTrailingStrongToken(List<IMarkdownInline> nodes) {
        if (nodes.Count < 2
            || nodes[nodes.Count - 1] is not TextRun markerRun
            || markerRun.Text != "****"
            || nodes[nodes.Count - 2] is not TextRun tokenRun) {
            return false;
        }

        if (!TryExtractTrailingToken(tokenRun.Text, out var prefix, out var token)) {
            return false;
        }

        nodes.RemoveAt(nodes.Count - 1);
        nodes.RemoveAt(nodes.Count - 1);
        if (prefix.Length > 0) {
            nodes.Add(new TextRun(prefix));
        }

        nodes.Add(new BoldInline(token));
        return true;
    }

    private static bool TryRewriteOveropenedMetricStrongRuns(List<IMarkdownInline> nodes) {
        if (nodes.Count < 3 || nodes[nodes.Count - 1] is not IStrongMarkdownInline) {
            return false;
        }

        var openerCount = 0;
        var openerStart = nodes.Count - 1;
        while (openerStart > 0
               && nodes[openerStart - 1] is TextRun opener
               && opener.Text == "**") {
            openerStart--;
            openerCount++;
        }

        if (openerCount < 2) {
            return false;
        }

        nodes.RemoveRange(openerStart, openerCount);
        return true;
    }

    private static bool TryRewriteAdjacentMetricStrongValue(List<IMarkdownInline> nodes) {
        if (nodes.Count < 5
            || nodes[nodes.Count - 1] is not IStrongMarkdownInline
            || nodes[nodes.Count - 2] is not TextRun trailingStrongOpen
            || trailingStrongOpen.Text != "**"
            || nodes[nodes.Count - 3] is not TextRun symbolValue
            || !IsSymbolOnlyValue(symbolValue.Text)
            || nodes[nodes.Count - 4] is not TextRun leadingStrongOpen
            || leadingStrongOpen.Text != "**") {
            return false;
        }

        var symbol = symbolValue.Text.Trim();
        if (symbol.Length == 0) {
            return false;
        }

        nodes.RemoveRange(nodes.Count - 4, 3);
        nodes.Insert(nodes.Count - 1, new TextRun(symbol + " "));
        return true;
    }

    private static bool TryRewriteMissingTrailingStrongClose(List<IMarkdownInline> nodes) {
        if (nodes.Count < 3
            || nodes[nodes.Count - 1] is not ItalicSequenceInline italic
            || italic.Inlines.Nodes.Count == 0
            || nodes[nodes.Count - 2] is not TextRun strayOpen
            || strayOpen.Text != "*") {
            return false;
        }

        nodes.RemoveAt(nodes.Count - 2);
        nodes[nodes.Count - 1] = new BoldSequenceInline(italic.Inlines);
        return true;
    }

    private static bool TryExtractTrailingToken(string text, out string prefix, out string token) {
        prefix = string.Empty;
        token = string.Empty;
        if (string.IsNullOrWhiteSpace(text)) {
            return false;
        }

        var end = text.Length - 1;
        while (end >= 0 && char.IsWhiteSpace(text[end])) {
            end--;
        }

        if (end < 0) {
            return false;
        }

        var start = end;
        while (start >= 0 && IsTokenChar(text[start])) {
            start--;
        }

        token = text.Substring(start + 1, end - start).Trim();
        if (token.Length == 0 || token.IndexOf("**", StringComparison.Ordinal) >= 0) {
            return false;
        }

        prefix = text.Substring(0, start + 1);
        if (prefix.Length == 0) {
            return false;
        }

        return char.IsWhiteSpace(prefix[prefix.Length - 1]) || char.IsPunctuation(prefix[prefix.Length - 1]) || char.IsSymbol(prefix[prefix.Length - 1]);
    }

    private static bool IsTokenChar(char ch) {
        return char.IsLetterOrDigit(ch) || ch == '_' || ch == '.' || ch == '/' || ch == ':' || ch == '-';
    }

    private static bool IsSymbolOnlyValue(string value) {
        if (string.IsNullOrWhiteSpace(value)) {
            return false;
        }

        for (var i = 0; i < value.Length; i++) {
            var ch = value[i];
            if (char.IsWhiteSpace(ch)) {
                continue;
            }

            if (char.IsLetterOrDigit(ch)) {
                return false;
            }
        }

        return true;
    }

    private static List<IMarkdownInline> CoalesceAdjacentTextRuns(List<IMarkdownInline> nodes) {
        if (nodes.Count <= 1) {
            return nodes;
        }

        var compact = new List<IMarkdownInline>(nodes.Count);
        System.Text.StringBuilder? textBuffer = null;

        void FlushTextBuffer() {
            if (textBuffer == null) {
                return;
            }

            compact.Add(new TextRun(textBuffer.ToString()));
            textBuffer = null;
        }

        for (var i = 0; i < nodes.Count; i++) {
            var node = nodes[i];
            if (node is TextRun textRun) {
                textBuffer ??= new System.Text.StringBuilder();
                textBuffer.Append(textRun.Text);
                continue;
            }

            FlushTextBuffer();
            compact.Add(node);
        }

        FlushTextBuffer();
        return compact;
    }
}
