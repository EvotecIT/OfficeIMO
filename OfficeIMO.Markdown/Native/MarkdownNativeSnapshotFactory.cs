namespace OfficeIMO.Markdown;

internal static class MarkdownNativeSnapshotFactory {
    internal static MarkdownNativeDocumentSnapshot FromDocument(MarkdownNativeDocument document) {
        var referenceDefinitions = new List<MarkdownNativeReferenceLinkDefinitionSnapshot>(document.ReferenceLinkDefinitions.Count);
        for (var i = 0; i < document.ReferenceLinkDefinitions.Count; i++) {
            referenceDefinitions.Add(new MarkdownNativeReferenceLinkDefinitionSnapshot(document, document.ReferenceLinkDefinitions[i]));
        }

        var abbreviationDefinitions = new List<MarkdownNativeAbbreviationDefinitionSnapshot>(document.AbbreviationDefinitions.Count);
        for (var i = 0; i < document.AbbreviationDefinitions.Count; i++) {
            abbreviationDefinitions.Add(new MarkdownNativeAbbreviationDefinitionSnapshot(document, document.AbbreviationDefinitions[i]));
        }

        var blocks = new List<MarkdownNativeBlockSnapshot>(document.Blocks.Count);
        for (var i = 0; i < document.Blocks.Count; i++) {
            blocks.Add(FromBlock(document, document.Blocks[i]));
        }

        var sourceTrivia = new List<MarkdownNativeSourceTriviaSnapshot>(document.SourceTrivia.Count);
        for (var i = 0; i < document.SourceTrivia.Count; i++) {
            sourceTrivia.Add(FromSourceTrivia(document, document.SourceTrivia[i]));
        }

        var diagnostics = new List<MarkdownNativeDiagnosticSnapshot>(document.Diagnostics.Count);
        for (var i = 0; i < document.Diagnostics.Count; i++) {
            diagnostics.Add(new MarkdownNativeDiagnosticSnapshot(document.Diagnostics[i]));
        }

        return new MarkdownNativeDocumentSnapshot(document.SourceKind, referenceDefinitions, abbreviationDefinitions, sourceTrivia, blocks, diagnostics);
    }

    internal static MarkdownNativeBlockSnapshot FromBlock(MarkdownNativeBlock block) => FromBlock(null, block);

    internal static MarkdownNativeBlockSnapshot FromBlock(MarkdownNativeDocument? document, MarkdownNativeBlock block) {
        var snapshot = new MarkdownNativeBlockSnapshot {
            Id = block.Id,
            Kind = block.Kind,
            SourceSpan = ToSpanSnapshot(block.SourceSpan)
        };

        switch (block) {
            case MarkdownNativeParagraphBlock paragraph:
                snapshot.Text = paragraph.Text;
                snapshot.Markdown = RenderBlock(paragraph.Paragraph);
                snapshot.Inlines = FromInlines(document, paragraph.InlineRuns);
                snapshot.FieldSourceSpans = FieldSpans(("paragraphText", paragraph.TextSourceSpan));
                break;
            case MarkdownNativeHeadingBlock heading:
                snapshot.Text = heading.Text;
                snapshot.Markdown = RenderBlock(heading.Heading);
                snapshot.Inlines = FromInlines(document, heading.InlineRuns);
                snapshot.Fields = Fields(("level", heading.Level.ToString(System.Globalization.CultureInfo.InvariantCulture)));
                snapshot.FieldSourceSpans = FieldSpans(
                    ("level", heading.LevelSourceSpan),
                    ("text", heading.TextSourceSpan));
                break;
            case MarkdownNativeCodeBlock code:
                snapshot.Text = code.Content;
                snapshot.Markdown = RenderBlock(code.Code);
                snapshot.Fields = Fields(
                    ("openingFence", code.OpeningFence),
                    ("language", code.Language),
                    ("infoString", code.InfoString),
                    ("attributes", MarkdownNativeFenceInfoSourceSpans.GetAttributeSourceText(code.FenceInfo)),
                    ("caption", code.Caption),
                    ("title", code.Title),
                    ("elementId", code.ElementId),
                    ("closingFence", code.ClosingFence));
                snapshot.FieldSourceSpans = FieldSpans(
                    ("openingFence", code.OpeningFenceSourceSpan),
                    ("infoString", code.InfoStringSourceSpan),
                    ("attributes", code.AttributeSourceSpan),
                    ("content", code.ContentSourceSpan),
                    ("closingFence", code.ClosingFenceSourceSpan));
                break;
            case MarkdownNativeThematicBreakBlock thematicBreak:
                snapshot.Markdown = thematicBreak.Marker;
                break;
            case MarkdownNativeVisualBlock visual:
                snapshot.Text = visual.Content;
                snapshot.Markdown = RenderBlock(visual.Visual);
                snapshot.Fields = Fields(
                    ("openingFence", visual.OpeningFence),
                    ("semanticKind", visual.SemanticKind),
                    ("language", visual.Language),
                    ("infoString", visual.InfoString),
                    ("attributes", MarkdownNativeFenceInfoSourceSpans.GetAttributeSourceText(visual.FenceInfo)),
                    ("caption", visual.Caption),
                    ("title", visual.Title),
                    ("elementId", visual.ElementId),
                    ("payloadFormat", visual.Payload.Format.ToString()),
                    ("payloadDetectedSemanticKind", visual.Payload.DetectedSemanticKind),
                    ("payloadJsonType", visual.Payload.JsonType),
                    ("closingFence", visual.ClosingFence));
                snapshot.FieldSourceSpans = FieldSpans(
                    ("openingFence", visual.OpeningFenceSourceSpan),
                    ("infoString", visual.InfoStringSourceSpan),
                    ("attributes", visual.AttributeSourceSpan),
                    ("content", visual.ContentSourceSpan),
                    ("closingFence", visual.ClosingFenceSourceSpan));
                break;
            case MarkdownNativeTableBlock table:
                snapshot.HeaderCells = FromCells(document, table.HeaderCells);
                snapshot.Rows = FromRows(document, table.Rows);
                snapshot.HeaderRow = table.HeaderRow == null ? null : FromRow(document, table.HeaderRow);
                snapshot.BodyRows = FromTableRows(document, table.BodyRows);
                snapshot.FieldSourceSpans = FieldSpans(("alignmentRow", table.AlignmentRowSourceSpan));
                break;
            case MarkdownNativeQuoteBlock quote:
                snapshot.Markdown = RenderBlock(quote.Quote);
                snapshot.MarkerSourceSpans = ToSpanSnapshots(quote.MarkerSourceSpans);
                snapshot.Children = FromBlocks(document, quote.Children);
                snapshot.FieldSourceSpans = FieldSpans(("quoteBody", quote.BodySourceSpan));
                break;
            case MarkdownNativeCalloutBlock callout:
                snapshot.Text = callout.Title;
                snapshot.Markdown = RenderBlock(callout.Callout);
                snapshot.Inlines = FromInlines(document, callout.TitleInlineRuns);
                snapshot.Children = FromBlocks(document, callout.Children);
                snapshot.Fields = Fields(("calloutKind", callout.CalloutKind));
                snapshot.FieldSourceSpans = FieldSpans(
                    ("calloutOpeningMarker", callout.OpeningMarkerSourceSpan),
                    ("calloutKind", callout.KindSourceSpan),
                    ("calloutClosingMarker", callout.ClosingMarkerSourceSpan),
                    ("calloutTitle", callout.TitleSourceSpan),
                    ("title", callout.TitleSourceSpan),
                    ("calloutBody", callout.BodySourceSpan));
                break;
            case MarkdownNativeCustomContainerBlock customContainer:
                snapshot.Text = customContainer.Name;
                snapshot.Markdown = RenderBlock(customContainer.Container);
                snapshot.Children = FromBlocks(document, customContainer.Children);
                snapshot.Fields = Fields(
                    ("name", customContainer.Name),
                    ("info", customContainer.Info),
                    ("openingFence", customContainer.OpeningFence),
                    ("closingFence", customContainer.ClosingFence));
                snapshot.FieldSourceSpans = FieldSpans(
                    ("customContainerOpeningFence", customContainer.OpeningFenceSourceSpan),
                    ("customContainerInfo", customContainer.InfoSourceSpan),
                    ("customContainerName", customContainer.NameSourceSpan),
                    ("customContainerBody", customContainer.BodySourceSpan),
                    ("customContainerClosingFence", customContainer.ClosingFenceSourceSpan));
                break;
            case MarkdownNativeDetailsBlock details:
                snapshot.Text = details.Summary;
                snapshot.Markdown = RenderBlock(details.Details);
                snapshot.Inlines = FromInlines(document, details.SummaryInlineRuns);
                snapshot.Children = FromBlocks(document, details.Children);
                snapshot.Fields = Fields(("open", details.Open ? "true" : "false"));
                snapshot.FieldSourceSpans = FieldSpans(
                    ("detailsOpeningTag", details.OpeningTagSourceSpan),
                    ("summary", details.SummarySourceSpan),
                    ("summaryOpeningTag", details.SummaryOpeningTagSourceSpan),
                    ("summaryText", details.SummaryTextSourceSpan),
                    ("summaryClosingTag", details.SummaryClosingTagSourceSpan),
                    ("detailsBody", details.BodySourceSpan),
                    ("detailsClosingTag", details.ClosingTagSourceSpan));
                break;
            case MarkdownNativeDefinitionListBlock definitionList:
                snapshot.Markdown = RenderBlock(definitionList.DefinitionList);
                snapshot.Children = FromBlocks(document, definitionList.Children);
                snapshot.DefinitionGroups = FromDefinitionGroups(document, definitionList.Groups);
                break;
            case MarkdownNativeFootnoteDefinitionBlock footnote:
                snapshot.Text = footnote.Text;
                snapshot.Markdown = RenderBlock(footnote.Footnote);
                snapshot.Children = FromBlocks(document, footnote.Children);
                snapshot.Fields = Fields(("label", footnote.Label));
                snapshot.FieldSourceSpans = FieldSpans(
                    ("footnoteOpeningMarker", footnote.OpeningMarkerSourceSpan),
                    ("label", footnote.LabelSourceSpan),
                    ("footnoteSeparatorMarker", footnote.SeparatorMarkerSourceSpan),
                    ("footnoteBody", footnote.BodySourceSpan));
                break;
            case MarkdownNativeListBlock list:
                snapshot.Markdown = list.List.RenderMarkdown();
                snapshot.Items = FromListItems(document, list.Items);
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
                snapshot.FieldSourceSpans = FieldSpans(
                    ("alt", image.AltSourceSpan),
                    ("source", image.SourceSourceSpan),
                    ("title", image.TitleSourceSpan),
                    ("linkUrl", image.LinkUrlSourceSpan),
                    ("linkTitle", image.LinkTitleSourceSpan));
                break;
            case MarkdownNativeFrontMatterBlock frontMatter:
                snapshot.Markdown = RenderBlock(frontMatter.FrontMatter);
                snapshot.Fields = FromFrontMatter(frontMatter);
                snapshot.FieldSourceSpans = FieldSpans(
                    ("openingFence", frontMatter.OpeningFenceSourceSpan),
                    ("frontMatterBody", frontMatter.BodySourceSpan),
                    ("closingFence", frontMatter.ClosingFenceSourceSpan));
                break;
            case MarkdownNativeHtmlBlock html:
                snapshot.Text = html.Html;
                snapshot.Markdown = html.Html;
                snapshot.Fields = Fields(("isComment", html.IsComment ? "true" : "false"));
                snapshot.FieldSourceSpans = FieldSpans(
                    ("html", html.SourceSpan),
                    ("htmlOpeningTag", html.OpeningTagSourceSpan),
                    ("htmlOpeningMarker", html.RawOpeningMarkerSourceSpan),
                    ("htmlBody", html.RawBodySourceSpan),
                    ("htmlClosingMarker", html.RawClosingMarkerSourceSpan),
                    ("htmlClosingTag", html.ClosingTagSourceSpan),
                    ("htmlCommentOpeningMarker", html.OpeningMarkerSourceSpan),
                    ("htmlCommentBody", html.BodySourceSpan),
                    ("htmlCommentClosingMarker", html.ClosingMarkerSourceSpan));
                break;
            case MarkdownNativeOtherBlock other:
                snapshot.Markdown = other.Markdown;
                break;
        }

        snapshot.SourceFields = FromSourceFields(document, block);
        return snapshot;
    }

    private static MarkdownNativeSourceTriviaSnapshot FromSourceTrivia(
        MarkdownNativeDocument document,
        MarkdownNativeSourceTrivia trivia) {
        string? sourceText = null;
        string? originalSourceText = null;
        MarkdownOriginalSourceSliceFailureReason? originalFailureReason = null;

        if (document.TryCreateSourceSlice(trivia, out var sourceSlice)) {
            sourceText = sourceSlice.Text;
        }

        if (document.TryCreateOriginalSourceSlice(trivia, out var originalSlice, out var failureReason)) {
            originalSourceText = originalSlice.Text;
        } else {
            originalFailureReason = failureReason;
        }

        return new MarkdownNativeSourceTriviaSnapshot(
            trivia,
            sourceText,
            originalSourceText,
            originalFailureReason);
    }

    private static IReadOnlyDictionary<string, string?> FromFrontMatter(MarkdownNativeFrontMatterBlock frontMatter) {
        var fields = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);
        if (frontMatter.RawYaml != null) {
            fields["rawYaml"] = frontMatter.RawYaml;
        }

        for (var i = 0; i < frontMatter.Entries.Count; i++) {
            fields[frontMatter.Entries[i].Key] = frontMatter.Entries[i].Value?.ToString();
        }

        return fields;
    }

    private static string RenderBlock(IMarkdownBlock block) => block.RenderMarkdown();

    private static IReadOnlyList<MarkdownNativeBlockSnapshot> FromBlocks(MarkdownNativeDocument? document, IReadOnlyList<MarkdownNativeBlock> blocks) {
        if (blocks == null || blocks.Count == 0) {
            return Array.Empty<MarkdownNativeBlockSnapshot>();
        }

        var snapshots = new List<MarkdownNativeBlockSnapshot>(blocks.Count);
        for (var i = 0; i < blocks.Count; i++) {
            snapshots.Add(FromBlock(document, blocks[i]));
        }

        return snapshots;
    }

    private static IReadOnlyList<MarkdownNativeInlineSnapshot> FromInlines(
        MarkdownNativeDocument? document,
        IReadOnlyList<MarkdownNativeInline>? inlines) {
        if (inlines == null || inlines.Count == 0) {
            return Array.Empty<MarkdownNativeInlineSnapshot>();
        }

        var snapshots = new List<MarkdownNativeInlineSnapshot>(inlines.Count);
        for (var i = 0; i < inlines.Count; i++) {
            snapshots.Add(FromInline(document, inlines[i]));
        }

        return snapshots;
    }

    private static MarkdownNativeInlineSnapshot FromInline(MarkdownNativeDocument? document, MarkdownNativeInline inline) {
        var metadata = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var metadataSourceSpans = new Dictionary<string, MarkdownNativeSourceSpanSnapshot?>(StringComparer.OrdinalIgnoreCase);
        for (var i = 0; i < inline.Metadata.Count; i++) {
            metadata[inline.Metadata[i].Name] = inline.Metadata[i].Value;
            metadataSourceSpans[inline.Metadata[i].Name] = ToSpanSnapshot(inline.Metadata[i].SourceSpan);
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
            metadataSourceSpans,
            FromInlineMetadata(document, inline.Metadata),
            FromInlines(document, inline.Children));
    }

    private static IReadOnlyList<MarkdownNativeInlineMetadataSnapshot> FromInlineMetadata(
        MarkdownNativeDocument? document,
        IReadOnlyList<MarkdownNativeInlineMetadata>? metadata) {
        if (metadata == null || metadata.Count == 0) {
            return Array.Empty<MarkdownNativeInlineMetadataSnapshot>();
        }

        var orderedMetadata = metadata
            .Select((value, originalIndex) => new { value, originalIndex })
            .OrderBy(item => item.value.SourceSpan.HasValue ? 0 : 1)
            .ThenBy(item => item.value.SourceSpan?.StartLine ?? int.MaxValue)
            .ThenBy(item => item.value.SourceSpan?.StartColumn ?? int.MaxValue)
            .ThenBy(item => item.originalIndex)
            .ToArray();

        var snapshots = new List<MarkdownNativeInlineMetadataSnapshot>(orderedMetadata.Length);
        for (var i = 0; i < orderedMetadata.Length; i++) {
            if (document == null) {
                snapshots.Add(new MarkdownNativeInlineMetadataSnapshot(orderedMetadata[i].value, i));
                continue;
            }

            string? sourceText = null;
            string? originalSourceText = null;
            MarkdownOriginalSourceSliceFailureReason? originalFailureReason = null;

            if (document.TryCreateSourceSlice(orderedMetadata[i].value, out var sourceSlice)) {
                sourceText = sourceSlice.Text;
            }

            if (document.TryCreateOriginalSourceSlice(orderedMetadata[i].value, out var originalSlice, out var failureReason)) {
                originalSourceText = originalSlice.Text;
            } else {
                originalFailureReason = failureReason;
            }

            snapshots.Add(new MarkdownNativeInlineMetadataSnapshot(
                orderedMetadata[i].value,
                i,
                sourceText,
                originalSourceText,
                originalFailureReason));
        }

        return snapshots;
    }

    private static IReadOnlyList<MarkdownNativeListItemSnapshot> FromListItems(
        MarkdownNativeDocument? document,
        IReadOnlyList<MarkdownNativeListItem> items) {
        if (items == null || items.Count == 0) {
            return Array.Empty<MarkdownNativeListItemSnapshot>();
        }

        var snapshots = new List<MarkdownNativeListItemSnapshot>(items.Count);
        for (var i = 0; i < items.Count; i++) {
            string? sourceText = null;
            string? originalSourceText = null;
            MarkdownOriginalSourceSliceFailureReason? originalFailureReason = null;
            var failureReason = MarkdownOriginalSourceSliceFailureReason.None;

            if (document != null && document.TryCreateSourceSlice(items[i], out var sourceSlice)) {
                sourceText = sourceSlice.Text;
            }

            if (document != null && document.TryCreateOriginalSourceSlice(items[i], out var originalSlice, out failureReason)) {
                originalSourceText = originalSlice.Text;
            } else if (document != null) {
                originalFailureReason = failureReason;
            }

            snapshots.Add(new MarkdownNativeListItemSnapshot(
                items[i].Id,
                items[i].Text,
                items[i].IsTask,
                items[i].Checked,
                items[i].Level,
                ToSpanSnapshot(items[i].SourceSpan),
                ToSpanSnapshot(items[i].MarkerSourceSpan),
                items[i].MarkerText,
                ToSpanSnapshot(items[i].TaskMarkerSourceSpan),
                items[i].TaskMarkerText,
                sourceText,
                originalSourceText,
                originalFailureReason,
                FromInlines(document, items[i].InlineRuns),
                FromListItemParagraphs(document, items[i].Paragraphs),
                FromBlocks(document, items[i].Children)));
        }

        return snapshots;
    }

    private static IReadOnlyList<MarkdownNativeListItemParagraphSnapshot> FromListItemParagraphs(
        MarkdownNativeDocument? document,
        IReadOnlyList<MarkdownNativeListItemParagraph> paragraphs) {
        if (paragraphs == null || paragraphs.Count == 0) {
            return Array.Empty<MarkdownNativeListItemParagraphSnapshot>();
        }

        var snapshots = new List<MarkdownNativeListItemParagraphSnapshot>(paragraphs.Count);
        for (var i = 0; i < paragraphs.Count; i++) {
            string? sourceText = null;
            string? originalSourceText = null;
            MarkdownOriginalSourceSliceFailureReason? originalFailureReason = null;
            var failureReason = MarkdownOriginalSourceSliceFailureReason.None;

            if (document != null && document.TryCreateSourceSlice(paragraphs[i], out var sourceSlice)) {
                sourceText = sourceSlice.Text;
            }

            if (document != null && document.TryCreateOriginalSourceSlice(paragraphs[i], out var originalSlice, out failureReason)) {
                originalSourceText = originalSlice.Text;
            } else if (document != null) {
                originalFailureReason = failureReason;
            }

            snapshots.Add(new MarkdownNativeListItemParagraphSnapshot(
                paragraphs[i].Index,
                paragraphs[i].Text,
                ToSpanSnapshot(paragraphs[i].SourceSpan),
                sourceText,
                originalSourceText,
                originalFailureReason,
                FromInlines(document, paragraphs[i].InlineRuns)));
        }

        return snapshots;
    }

    private static IReadOnlyList<MarkdownNativeDefinitionListGroupSnapshot> FromDefinitionGroups(
        MarkdownNativeDocument? document,
        IReadOnlyList<MarkdownNativeDefinitionListGroup> groups) {
        if (groups == null || groups.Count == 0) {
            return Array.Empty<MarkdownNativeDefinitionListGroupSnapshot>();
        }

        var snapshots = new List<MarkdownNativeDefinitionListGroupSnapshot>(groups.Count);
        for (var i = 0; i < groups.Count; i++) {
            string? sourceText = null;
            string? originalSourceText = null;
            MarkdownOriginalSourceSliceFailureReason? originalFailureReason = null;
            var failureReason = MarkdownOriginalSourceSliceFailureReason.None;

            if (document != null && document.TryCreateSourceSlice(groups[i], out var sourceSlice)) {
                sourceText = sourceSlice.Text;
            }

            if (document != null && document.TryCreateOriginalSourceSlice(groups[i], out var originalSlice, out failureReason)) {
                originalSourceText = originalSlice.Text;
            } else if (document != null) {
                originalFailureReason = failureReason;
            }

            snapshots.Add(new MarkdownNativeDefinitionListGroupSnapshot(
                ToSpanSnapshot(groups[i].SourceSpan),
                sourceText,
                originalSourceText,
                originalFailureReason,
                FromDefinitionTerms(document, groups[i].Terms),
                FromDefinitions(document, groups[i].Definitions)));
        }

        return snapshots;
    }

    private static IReadOnlyList<MarkdownNativeDefinitionListTermSnapshot> FromDefinitionTerms(
        MarkdownNativeDocument? document,
        IReadOnlyList<MarkdownNativeDefinitionListTerm> terms) {
        if (terms == null || terms.Count == 0) {
            return Array.Empty<MarkdownNativeDefinitionListTermSnapshot>();
        }

        var snapshots = new List<MarkdownNativeDefinitionListTermSnapshot>(terms.Count);
        for (var i = 0; i < terms.Count; i++) {
            string? sourceText = null;
            string? originalSourceText = null;
            MarkdownOriginalSourceSliceFailureReason? originalFailureReason = null;
            var failureReason = MarkdownOriginalSourceSliceFailureReason.None;

            if (document != null && document.TryCreateSourceSlice(terms[i], out var sourceSlice)) {
                sourceText = sourceSlice.Text;
            }

            if (document != null && document.TryCreateOriginalSourceSlice(terms[i], out var originalSlice, out failureReason)) {
                originalSourceText = originalSlice.Text;
            } else if (document != null) {
                originalFailureReason = failureReason;
            }

            snapshots.Add(new MarkdownNativeDefinitionListTermSnapshot(
                terms[i].Text,
                terms[i].Markdown,
                ToSpanSnapshot(terms[i].SourceSpan),
                sourceText,
                originalSourceText,
                originalFailureReason,
                FromInlines(document, terms[i].InlineRuns)));
        }

        return snapshots;
    }

    private static IReadOnlyList<MarkdownNativeDefinitionListDefinitionSnapshot> FromDefinitions(
        MarkdownNativeDocument? document,
        IReadOnlyList<MarkdownNativeDefinitionListDefinition> definitions) {
        if (definitions == null || definitions.Count == 0) {
            return Array.Empty<MarkdownNativeDefinitionListDefinitionSnapshot>();
        }

        var snapshots = new List<MarkdownNativeDefinitionListDefinitionSnapshot>(definitions.Count);
        for (var i = 0; i < definitions.Count; i++) {
            string? sourceText = null;
            string? originalSourceText = null;
            MarkdownOriginalSourceSliceFailureReason? originalFailureReason = null;
            var failureReason = MarkdownOriginalSourceSliceFailureReason.None;

            if (document != null && document.TryCreateSourceSlice(definitions[i], out var sourceSlice)) {
                sourceText = sourceSlice.Text;
            }

            if (document != null && document.TryCreateOriginalSourceSlice(definitions[i], out var originalSlice, out failureReason)) {
                originalSourceText = originalSlice.Text;
            } else if (document != null) {
                originalFailureReason = failureReason;
            }

            snapshots.Add(new MarkdownNativeDefinitionListDefinitionSnapshot(
                definitions[i].Markdown,
                ToSpanSnapshot(definitions[i].SourceSpan),
                sourceText,
                originalSourceText,
                originalFailureReason,
                FromBlocks(document, definitions[i].Children)));
        }

        return snapshots;
    }

    private static IReadOnlyList<MarkdownNativeTableCellSnapshot> FromCells(
        MarkdownNativeDocument? document,
        IReadOnlyList<MarkdownNativeTableCell> cells) {
        if (cells == null || cells.Count == 0) {
            return Array.Empty<MarkdownNativeTableCellSnapshot>();
        }

        var snapshots = new List<MarkdownNativeTableCellSnapshot>(cells.Count);
        for (var i = 0; i < cells.Count; i++) {
            snapshots.Add(FromCell(document, cells[i]));
        }

        return snapshots;
    }

    private static IReadOnlyList<IReadOnlyList<MarkdownNativeTableCellSnapshot>> FromRows(
        MarkdownNativeDocument? document,
        IReadOnlyList<IReadOnlyList<MarkdownNativeTableCell>> rows) {
        if (rows == null || rows.Count == 0) {
            return Array.Empty<IReadOnlyList<MarkdownNativeTableCellSnapshot>>();
        }

        var snapshots = new List<IReadOnlyList<MarkdownNativeTableCellSnapshot>>(rows.Count);
        for (var i = 0; i < rows.Count; i++) {
            snapshots.Add(FromCells(document, rows[i]));
        }

        return snapshots;
    }

    private static MarkdownNativeTableCellSnapshot FromCell(MarkdownNativeDocument? document, MarkdownNativeTableCell cell) {
        string? sourceText = null;
        string? originalSourceText = null;
        MarkdownOriginalSourceSliceFailureReason? originalFailureReason = null;
        var failureReason = MarkdownOriginalSourceSliceFailureReason.None;

        if (document != null && document.TryCreateSourceSlice(cell, out var sourceSlice)) {
            sourceText = sourceSlice.Text;
        }

        if (document != null && document.TryCreateOriginalSourceSlice(cell, out var originalSlice, out failureReason)) {
            originalSourceText = originalSlice.Text;
        } else if (document != null) {
            originalFailureReason = failureReason;
        }

        return new MarkdownNativeTableCellSnapshot(
            cell.Text,
            cell.Markdown,
            cell.IsHeader,
            cell.RowIndex,
            cell.ColumnIndex,
            cell.Alignment,
            ToSpanSnapshot(cell.SourceSpan),
            sourceText,
            originalSourceText,
            originalFailureReason,
            FromInlines(document, cell.InlineRuns),
            FromBlocks(document, cell.Children));
    }

    private static IReadOnlyList<MarkdownNativeTableRowSnapshot> FromTableRows(
        MarkdownNativeDocument? document,
        IReadOnlyList<MarkdownNativeTableRow> rows) {
        if (rows == null || rows.Count == 0) {
            return Array.Empty<MarkdownNativeTableRowSnapshot>();
        }

        var snapshots = new List<MarkdownNativeTableRowSnapshot>(rows.Count);
        for (var i = 0; i < rows.Count; i++) {
            snapshots.Add(FromRow(document, rows[i]));
        }

        return snapshots;
    }

    private static MarkdownNativeTableRowSnapshot FromRow(MarkdownNativeDocument? document, MarkdownNativeTableRow row) {
        string? sourceText = null;
        string? originalSourceText = null;
        MarkdownOriginalSourceSliceFailureReason? originalFailureReason = null;
        var failureReason = MarkdownOriginalSourceSliceFailureReason.None;

        if (document != null && document.TryCreateSourceSlice(row, out var sourceSlice)) {
            sourceText = sourceSlice.Text;
        }

        if (document != null && document.TryCreateOriginalSourceSlice(row, out var originalSlice, out failureReason)) {
            originalSourceText = originalSlice.Text;
        } else if (document != null) {
            originalFailureReason = failureReason;
        }

        return new MarkdownNativeTableRowSnapshot(
            row.Markdown,
            row.IsHeader,
            row.RowIndex,
            ToSpanSnapshot(row.SourceSpan),
            sourceText,
            originalSourceText,
            originalFailureReason,
            FromCells(document, row.Cells));
    }

    private static MarkdownNativeSourceSpanSnapshot? ToSpanSnapshot(MarkdownSourceSpan? span) =>
        span.HasValue ? new MarkdownNativeSourceSpanSnapshot(span.Value) : null;

    private static IReadOnlyList<MarkdownNativeSourceSpanSnapshot> ToSpanSnapshots(IReadOnlyList<MarkdownSourceSpan> spans) {
        if (spans == null || spans.Count == 0) {
            return Array.Empty<MarkdownNativeSourceSpanSnapshot>();
        }

        var snapshots = new List<MarkdownNativeSourceSpanSnapshot>(spans.Count);
        for (var i = 0; i < spans.Count; i++) {
            snapshots.Add(new MarkdownNativeSourceSpanSnapshot(spans[i]));
        }

        return snapshots;
    }

    private static IReadOnlyList<MarkdownNativeBlockSourceFieldSnapshot> FromSourceFields(
        MarkdownNativeDocument? document,
        MarkdownNativeBlock block) {
        var fields = MarkdownNativeDocument.EnumerateBlockSourceFields(block).ToArray();
        if (fields.Length == 0) {
            return Array.Empty<MarkdownNativeBlockSourceFieldSnapshot>();
        }

        var snapshots = new List<MarkdownNativeBlockSourceFieldSnapshot>(fields.Length);
        for (var i = 0; i < fields.Length; i++) {
            if (document == null) {
                snapshots.Add(new MarkdownNativeBlockSourceFieldSnapshot(fields[i]));
                continue;
            }

            string? sourceText = null;
            string? originalSourceText = null;
            MarkdownOriginalSourceSliceFailureReason? originalFailureReason = null;

            if (document.TryCreateSourceSlice(fields[i], out var sourceSlice)) {
                sourceText = sourceSlice.Text;
            }

            if (document.TryCreateOriginalSourceSlice(fields[i], out var originalSlice, out var failureReason)) {
                originalSourceText = originalSlice.Text;
            } else {
                originalFailureReason = failureReason;
            }

            snapshots.Add(new MarkdownNativeBlockSourceFieldSnapshot(
                fields[i],
                sourceText,
                originalSourceText,
                originalFailureReason));
        }

        return snapshots;
    }

    private static IReadOnlyDictionary<string, MarkdownNativeSourceSpanSnapshot?> FieldSpans(params (string Key, MarkdownSourceSpan? Value)[] values) {
        var fields = new Dictionary<string, MarkdownNativeSourceSpanSnapshot?>(StringComparer.OrdinalIgnoreCase);
        if (values == null) {
            return fields;
        }

        for (var i = 0; i < values.Length; i++) {
            if (!string.IsNullOrWhiteSpace(values[i].Key)) {
                fields[values[i].Key] = ToSpanSnapshot(values[i].Value);
            }
        }

        return fields;
    }

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
