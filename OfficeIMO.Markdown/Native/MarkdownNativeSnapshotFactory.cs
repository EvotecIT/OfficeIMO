namespace OfficeIMO.Markdown;

internal static class MarkdownNativeSnapshotFactory {
    internal static MarkdownNativeDocumentSnapshot FromDocument(MarkdownNativeDocument document) {
        var blocks = new List<MarkdownNativeBlockSnapshot>(document.Blocks.Count);
        for (var i = 0; i < document.Blocks.Count; i++) {
            blocks.Add(FromBlock(document.Blocks[i]));
        }

        var diagnostics = new List<MarkdownNativeDiagnosticSnapshot>(document.Diagnostics.Count);
        for (var i = 0; i < document.Diagnostics.Count; i++) {
            diagnostics.Add(new MarkdownNativeDiagnosticSnapshot(document.Diagnostics[i]));
        }

        return new MarkdownNativeDocumentSnapshot(document.SourceKind, blocks, diagnostics);
    }

    internal static MarkdownNativeBlockSnapshot FromBlock(MarkdownNativeBlock block) {
        var snapshot = new MarkdownNativeBlockSnapshot {
            Id = block.Id,
            Kind = block.Kind,
            SourceSpan = ToSpanSnapshot(block.SourceSpan)
        };

        switch (block) {
            case MarkdownNativeParagraphBlock paragraph:
                snapshot.Text = paragraph.Text;
                snapshot.Markdown = RenderBlock(paragraph.Paragraph);
                snapshot.Inlines = FromInlines(paragraph.InlineRuns);
                break;
            case MarkdownNativeHeadingBlock heading:
                snapshot.Text = heading.Text;
                snapshot.Markdown = RenderBlock(heading.Heading);
                snapshot.Inlines = FromInlines(heading.InlineRuns);
                snapshot.Fields = Fields(("level", heading.Level.ToString(System.Globalization.CultureInfo.InvariantCulture)));
                break;
            case MarkdownNativeCodeBlock code:
                snapshot.Text = code.Content;
                snapshot.Markdown = RenderBlock(code.Code);
                snapshot.Fields = Fields(
                    ("language", code.Language),
                    ("infoString", code.InfoString),
                    ("caption", code.Caption),
                    ("title", code.Title),
                    ("elementId", code.ElementId));
                break;
            case MarkdownNativeVisualBlock visual:
                snapshot.Text = visual.Content;
                snapshot.Markdown = RenderBlock(visual.Visual);
                snapshot.Fields = Fields(
                    ("semanticKind", visual.SemanticKind),
                    ("language", visual.Language),
                    ("infoString", visual.InfoString),
                    ("caption", visual.Caption),
                    ("title", visual.Title),
                    ("elementId", visual.ElementId),
                    ("payloadFormat", visual.Payload.Format.ToString()),
                    ("payloadDetectedSemanticKind", visual.Payload.DetectedSemanticKind),
                    ("payloadJsonType", visual.Payload.JsonType));
                break;
            case MarkdownNativeTableBlock table:
                snapshot.HeaderCells = FromCells(table.HeaderCells);
                snapshot.Rows = FromRows(table.Rows);
                break;
            case MarkdownNativeQuoteBlock quote:
                snapshot.Markdown = RenderBlock(quote.Quote);
                snapshot.Children = FromBlocks(quote.Children);
                break;
            case MarkdownNativeCalloutBlock callout:
                snapshot.Text = callout.Title;
                snapshot.Markdown = RenderBlock(callout.Callout);
                snapshot.Inlines = FromInlines(callout.TitleInlineRuns);
                snapshot.Children = FromBlocks(callout.Children);
                snapshot.Fields = Fields(("calloutKind", callout.CalloutKind));
                break;
            case MarkdownNativeDetailsBlock details:
                snapshot.Text = details.Summary;
                snapshot.Markdown = RenderBlock(details.Details);
                snapshot.Inlines = FromInlines(details.SummaryInlineRuns);
                snapshot.Children = FromBlocks(details.Children);
                snapshot.Fields = Fields(("open", details.Open ? "true" : "false"));
                break;
            case MarkdownNativeListBlock list:
                snapshot.Markdown = list.List.RenderMarkdown();
                snapshot.Items = FromListItems(list.Items);
                snapshot.Fields = Fields(
                    ("isOrdered", list.IsOrdered ? "true" : "false"),
                    ("start", list.Start?.ToString(System.Globalization.CultureInfo.InvariantCulture)));
                break;
            case MarkdownNativeImageBlock image:
                snapshot.Text = image.PlainAlt ?? image.Alt;
                snapshot.Markdown = RenderBlock(image.Image);
                snapshot.Fields = Fields(
                    ("source", image.Source),
                    ("alt", image.Alt),
                    ("plainAlt", image.PlainAlt),
                    ("title", image.Title),
                    ("caption", image.Caption),
                    ("linkUrl", image.LinkUrl));
                break;
            case MarkdownNativeFrontMatterBlock frontMatter:
                snapshot.Markdown = RenderBlock(frontMatter.FrontMatter);
                snapshot.Fields = FromFrontMatter(frontMatter);
                break;
            case MarkdownNativeHtmlBlock html:
                snapshot.Text = html.Html;
                snapshot.Markdown = html.Html;
                snapshot.Fields = Fields(("isComment", html.IsComment ? "true" : "false"));
                break;
            case MarkdownNativeOtherBlock other:
                snapshot.Markdown = other.Markdown;
                break;
        }

        return snapshot;
    }

    private static IReadOnlyDictionary<string, string?> FromFrontMatter(MarkdownNativeFrontMatterBlock frontMatter) {
        var fields = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);
        for (var i = 0; i < frontMatter.Entries.Count; i++) {
            fields[frontMatter.Entries[i].Key] = frontMatter.Entries[i].Value?.ToString();
        }

        return fields;
    }

    private static string RenderBlock(IMarkdownBlock block) => block.RenderMarkdown();

    private static IReadOnlyList<MarkdownNativeBlockSnapshot> FromBlocks(IReadOnlyList<MarkdownNativeBlock> blocks) {
        if (blocks == null || blocks.Count == 0) {
            return Array.Empty<MarkdownNativeBlockSnapshot>();
        }

        var snapshots = new List<MarkdownNativeBlockSnapshot>(blocks.Count);
        for (var i = 0; i < blocks.Count; i++) {
            snapshots.Add(FromBlock(blocks[i]));
        }

        return snapshots;
    }

    private static IReadOnlyList<MarkdownNativeInlineSnapshot> FromInlines(IReadOnlyList<MarkdownNativeInline>? inlines) {
        if (inlines == null || inlines.Count == 0) {
            return Array.Empty<MarkdownNativeInlineSnapshot>();
        }

        var snapshots = new List<MarkdownNativeInlineSnapshot>(inlines.Count);
        for (var i = 0; i < inlines.Count; i++) {
            snapshots.Add(FromInline(inlines[i]));
        }

        return snapshots;
    }

    private static MarkdownNativeInlineSnapshot FromInline(MarkdownNativeInline inline) {
        var metadata = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        for (var i = 0; i < inline.Metadata.Count; i++) {
            metadata[inline.Metadata[i].Name] = inline.Metadata[i].Value;
        }

        return new MarkdownNativeInlineSnapshot(
            inline.Id,
            inline.Kind,
            inline.SyntaxKind,
            inline.Text,
            inline.Markdown,
            inline.Literal,
            ToSpanSnapshot(inline.SourceSpan),
            metadata,
            FromInlines(inline.Children));
    }

    private static IReadOnlyList<MarkdownNativeListItemSnapshot> FromListItems(IReadOnlyList<MarkdownNativeListItem> items) {
        if (items == null || items.Count == 0) {
            return Array.Empty<MarkdownNativeListItemSnapshot>();
        }

        var snapshots = new List<MarkdownNativeListItemSnapshot>(items.Count);
        for (var i = 0; i < items.Count; i++) {
            snapshots.Add(new MarkdownNativeListItemSnapshot(
                items[i].Id,
                items[i].Text,
                items[i].IsTask,
                items[i].Checked,
                items[i].Level,
                ToSpanSnapshot(items[i].SourceSpan),
                FromInlines(items[i].InlineRuns),
                FromBlocks(items[i].Children)));
        }

        return snapshots;
    }

    private static IReadOnlyList<MarkdownNativeTableCellSnapshot> FromCells(IReadOnlyList<MarkdownNativeTableCell> cells) {
        if (cells == null || cells.Count == 0) {
            return Array.Empty<MarkdownNativeTableCellSnapshot>();
        }

        var snapshots = new List<MarkdownNativeTableCellSnapshot>(cells.Count);
        for (var i = 0; i < cells.Count; i++) {
            snapshots.Add(FromCell(cells[i]));
        }

        return snapshots;
    }

    private static IReadOnlyList<IReadOnlyList<MarkdownNativeTableCellSnapshot>> FromRows(IReadOnlyList<IReadOnlyList<MarkdownNativeTableCell>> rows) {
        if (rows == null || rows.Count == 0) {
            return Array.Empty<IReadOnlyList<MarkdownNativeTableCellSnapshot>>();
        }

        var snapshots = new List<IReadOnlyList<MarkdownNativeTableCellSnapshot>>(rows.Count);
        for (var i = 0; i < rows.Count; i++) {
            snapshots.Add(FromCells(rows[i]));
        }

        return snapshots;
    }

    private static MarkdownNativeTableCellSnapshot FromCell(MarkdownNativeTableCell cell) {
        return new MarkdownNativeTableCellSnapshot(
            cell.Text,
            cell.Markdown,
            cell.IsHeader,
            cell.RowIndex,
            cell.ColumnIndex,
            cell.Alignment,
            ToSpanSnapshot(cell.SourceSpan),
            FromInlines(cell.InlineRuns));
    }

    private static MarkdownNativeSourceSpanSnapshot? ToSpanSnapshot(MarkdownSourceSpan? span) =>
        span.HasValue ? new MarkdownNativeSourceSpanSnapshot(span.Value) : null;

    private static IReadOnlyDictionary<string, string?> Fields(params (string Key, string? Value)[] values) {
        var fields = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);
        if (values == null) {
            return fields;
        }

        for (var i = 0; i < values.Length; i++) {
            if (!string.IsNullOrWhiteSpace(values[i].Key) && values[i].Value != null) {
                fields[values[i].Key] = values[i].Value;
            }
        }

        return fields;
    }
}
