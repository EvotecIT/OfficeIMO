namespace OfficeIMO.Markdown;

/// <summary>
/// UI-safe snapshot of a native markdown document without parser object references.
/// </summary>
public sealed class MarkdownNativeDocumentSnapshot {
    internal MarkdownNativeDocumentSnapshot(
        MarkdownNativeDocumentSourceKind sourceKind,
        IReadOnlyList<MarkdownNativeReferenceLinkDefinitionSnapshot> referenceLinkDefinitions,
        IReadOnlyList<MarkdownNativeSourceTriviaSnapshot> sourceTrivia,
        IReadOnlyList<MarkdownNativeBlockSnapshot> blocks,
        IReadOnlyList<MarkdownNativeDiagnosticSnapshot> diagnostics) {
        SourceKind = sourceKind;
        ReferenceLinkDefinitions = referenceLinkDefinitions ?? Array.Empty<MarkdownNativeReferenceLinkDefinitionSnapshot>();
        SourceTrivia = sourceTrivia ?? Array.Empty<MarkdownNativeSourceTriviaSnapshot>();
        Blocks = blocks ?? Array.Empty<MarkdownNativeBlockSnapshot>();
        Diagnostics = diagnostics ?? Array.Empty<MarkdownNativeDiagnosticSnapshot>();
    }

    /// <summary>Identifies the markdown source backing this snapshot.</summary>
    public MarkdownNativeDocumentSourceKind SourceKind { get; }

    /// <summary>Effective reference-style link definitions collected during parsing.</summary>
    public IReadOnlyList<MarkdownNativeReferenceLinkDefinitionSnapshot> ReferenceLinkDefinitions { get; }

    /// <summary>Document-level source trivia such as blank lines, in source order.</summary>
    public IReadOnlyList<MarkdownNativeSourceTriviaSnapshot> SourceTrivia { get; }

    /// <summary>Top-level block snapshots.</summary>
    public IReadOnlyList<MarkdownNativeBlockSnapshot> Blocks { get; }

    /// <summary>Projection diagnostics.</summary>
    public IReadOnlyList<MarkdownNativeDiagnosticSnapshot> Diagnostics { get; }
}

/// <summary>
/// UI-safe snapshot of document-level source trivia.
/// </summary>
public sealed class MarkdownNativeSourceTriviaSnapshot {
    internal MarkdownNativeSourceTriviaSnapshot(MarkdownNativeSourceTrivia trivia) {
        Kind = trivia.Kind;
        Text = trivia.Text;
        SourceSpan = new MarkdownNativeSourceSpanSnapshot(trivia.SourceSpan);
    }

    /// <summary>Trivia kind.</summary>
    public MarkdownNativeSourceTriviaKind Kind { get; }

    /// <summary>Exact normalized line content represented by this trivia, excluding the line ending.</summary>
    public string Text { get; }

    /// <summary>Source span for the trivia content.</summary>
    public MarkdownNativeSourceSpanSnapshot SourceSpan { get; }
}

/// <summary>
/// UI-safe snapshot of a reference-style link definition.
/// </summary>
public sealed class MarkdownNativeReferenceLinkDefinitionSnapshot {
    internal MarkdownNativeReferenceLinkDefinitionSnapshot(MarkdownReferenceLinkDefinition definition) {
        Label = definition.Label;
        Url = definition.Url;
        Title = definition.Title;
        SourceSpan = definition.SourceSpan.HasValue ? new MarkdownNativeSourceSpanSnapshot(definition.SourceSpan.Value) : null;
        LabelSourceSpan = definition.LabelSourceSpan.HasValue ? new MarkdownNativeSourceSpanSnapshot(definition.LabelSourceSpan.Value) : null;
        OpeningMarkerSourceSpan = definition.OpeningMarkerSourceSpan.HasValue ? new MarkdownNativeSourceSpanSnapshot(definition.OpeningMarkerSourceSpan.Value) : null;
        SeparatorMarkerSourceSpan = definition.SeparatorMarkerSourceSpan.HasValue ? new MarkdownNativeSourceSpanSnapshot(definition.SeparatorMarkerSourceSpan.Value) : null;
        UrlSourceSpan = definition.UrlSourceSpan.HasValue ? new MarkdownNativeSourceSpanSnapshot(definition.UrlSourceSpan.Value) : null;
        TitleSourceSpan = definition.TitleSourceSpan.HasValue ? new MarkdownNativeSourceSpanSnapshot(definition.TitleSourceSpan.Value) : null;
        SourceFields = FromReferenceDefinitionFields(definition);
    }

    /// <summary>Normalized reference label.</summary>
    public string Label { get; }

    /// <summary>Resolved destination URL.</summary>
    public string Url { get; }

    /// <summary>Optional definition title.</summary>
    public string? Title { get; }

    /// <summary>Source span for the entire reference-style definition, when available.</summary>
    public MarkdownNativeSourceSpanSnapshot? SourceSpan { get; }

    /// <summary>Source span for the definition label token, when available.</summary>
    public MarkdownNativeSourceSpanSnapshot? LabelSourceSpan { get; }

    /// <summary>Source span for the opening <c>[</c> marker before the definition label, when available.</summary>
    public MarkdownNativeSourceSpanSnapshot? OpeningMarkerSourceSpan { get; }

    /// <summary>Source span for the <c>]:</c> marker after the definition label, when available.</summary>
    public MarkdownNativeSourceSpanSnapshot? SeparatorMarkerSourceSpan { get; }

    /// <summary>Source span for the destination token, when available.</summary>
    public MarkdownNativeSourceSpanSnapshot? UrlSourceSpan { get; }

    /// <summary>Source span for the optional title token, when available.</summary>
    public MarkdownNativeSourceSpanSnapshot? TitleSourceSpan { get; }

    /// <summary>Source-backed token and payload fields in source order.</summary>
    public IReadOnlyList<MarkdownNativeReferenceLinkDefinitionFieldSnapshot> SourceFields { get; }

    private static IReadOnlyList<MarkdownNativeReferenceLinkDefinitionFieldSnapshot> FromReferenceDefinitionFields(MarkdownReferenceLinkDefinition definition) {
        var fields = MarkdownNativeDocument.EnumerateReferenceLinkDefinitionFields(definition).ToArray();
        if (fields.Length == 0) {
            return Array.Empty<MarkdownNativeReferenceLinkDefinitionFieldSnapshot>();
        }

        var snapshots = new List<MarkdownNativeReferenceLinkDefinitionFieldSnapshot>(fields.Length);
        for (var i = 0; i < fields.Length; i++) {
            snapshots.Add(new MarkdownNativeReferenceLinkDefinitionFieldSnapshot(fields[i]));
        }

        return snapshots;
    }
}

/// <summary>
/// UI-safe snapshot of a source-backed token or payload field owned by a reference-style link definition.
/// </summary>
public sealed class MarkdownNativeReferenceLinkDefinitionFieldSnapshot {
    internal MarkdownNativeReferenceLinkDefinitionFieldSnapshot(MarkdownNativeReferenceLinkDefinitionField field) {
        Name = field.Name;
        Value = field.Value;
        SourceSpan = new MarkdownNativeSourceSpanSnapshot(field.SourceSpan);
    }

    /// <summary>Stable field name such as <c>openingMarker</c>, <c>label</c>, <c>separatorMarker</c>, <c>url</c>, or <c>title</c>.</summary>
    public string Name { get; }

    /// <summary>Semantic value represented by the field when one is available.</summary>
    public string? Value { get; }

    /// <summary>Source span for this field.</summary>
    public MarkdownNativeSourceSpanSnapshot SourceSpan { get; }
}

/// <summary>
/// UI-safe snapshot of a native block.
/// </summary>
public sealed class MarkdownNativeBlockSnapshot {
    internal MarkdownNativeBlockSnapshot() {
        Fields = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);
        FieldSourceSpans = new Dictionary<string, MarkdownNativeSourceSpanSnapshot?>(StringComparer.OrdinalIgnoreCase);
        SourceFields = Array.Empty<MarkdownNativeBlockSourceFieldSnapshot>();
        MarkerSourceSpans = Array.Empty<MarkdownNativeSourceSpanSnapshot>();
        Inlines = Array.Empty<MarkdownNativeInlineSnapshot>();
        Children = Array.Empty<MarkdownNativeBlockSnapshot>();
        Items = Array.Empty<MarkdownNativeListItemSnapshot>();
        DefinitionGroups = Array.Empty<MarkdownNativeDefinitionListGroupSnapshot>();
        HeaderCells = Array.Empty<MarkdownNativeTableCellSnapshot>();
        Rows = Array.Empty<IReadOnlyList<MarkdownNativeTableCellSnapshot>>();
    }

    /// <summary>Stable block id.</summary>
    public string Id { get; internal set; } = string.Empty;

    /// <summary>Native block kind.</summary>
    public MarkdownNativeBlockKind Kind { get; internal set; }

    /// <summary>Source span snapshot when available.</summary>
    public MarkdownNativeSourceSpanSnapshot? SourceSpan { get; internal set; }

    /// <summary>Common text payload when the block exposes one.</summary>
    public string? Text { get; internal set; }

    /// <summary>Common markdown payload when the block exposes one.</summary>
    public string? Markdown { get; internal set; }

    /// <summary>String fields for block-specific metadata.</summary>
    public IReadOnlyDictionary<string, string?> Fields { get; internal set; }

    /// <summary>Source spans for block-specific metadata fields when available.</summary>
    public IReadOnlyDictionary<string, MarkdownNativeSourceSpanSnapshot?> FieldSourceSpans { get; internal set; }

    /// <summary>Source-backed token and payload fields in source order, including repeated fields.</summary>
    public IReadOnlyList<MarkdownNativeBlockSourceFieldSnapshot> SourceFields { get; internal set; }

    /// <summary>Source spans for repeated marker tokens owned by this block, such as blockquote markers.</summary>
    public IReadOnlyList<MarkdownNativeSourceSpanSnapshot> MarkerSourceSpans { get; internal set; }

    /// <summary>Inline snapshots owned directly by this block.</summary>
    public IReadOnlyList<MarkdownNativeInlineSnapshot> Inlines { get; internal set; }

    /// <summary>Nested child block snapshots.</summary>
    public IReadOnlyList<MarkdownNativeBlockSnapshot> Children { get; internal set; }

    /// <summary>List item snapshots for native list blocks.</summary>
    public IReadOnlyList<MarkdownNativeListItemSnapshot> Items { get; internal set; }

    /// <summary>Definition-list group snapshots for native definition list blocks.</summary>
    public IReadOnlyList<MarkdownNativeDefinitionListGroupSnapshot> DefinitionGroups { get; internal set; }

    /// <summary>Table header cell snapshots.</summary>
    public IReadOnlyList<MarkdownNativeTableCellSnapshot> HeaderCells { get; internal set; }

    /// <summary>Table body row snapshots.</summary>
    public IReadOnlyList<IReadOnlyList<MarkdownNativeTableCellSnapshot>> Rows { get; internal set; }

    /// <summary>Enumerates source-backed fields with the supplied field name in source order.</summary>
    public IEnumerable<MarkdownNativeBlockSourceFieldSnapshot> EnumerateSourceFields(string name) {
        if (string.IsNullOrWhiteSpace(name) || SourceFields.Count == 0) {
            yield break;
        }

        for (var i = 0; i < SourceFields.Count; i++) {
            var field = SourceFields[i];
            if (string.Equals(field.Name, name, StringComparison.OrdinalIgnoreCase)) {
                yield return field;
            }
        }
    }

    /// <summary>
    /// Finds the first source-backed field with the supplied name, optionally constrained to a repeated-field occurrence index.
    /// </summary>
    public MarkdownNativeBlockSourceFieldSnapshot? FindSourceField(string name, int index = -1) {
        if (string.IsNullOrWhiteSpace(name) || SourceFields.Count == 0) {
            return null;
        }

        for (var i = 0; i < SourceFields.Count; i++) {
            var field = SourceFields[i];
            if (!string.Equals(field.Name, name, StringComparison.OrdinalIgnoreCase)) {
                continue;
            }

            if (index < 0 || field.Index == index) {
                return field;
            }
        }

        return null;
    }
}

/// <summary>
/// UI-safe snapshot of a source-backed token or payload field owned by a native block.
/// </summary>
public sealed class MarkdownNativeBlockSourceFieldSnapshot {
    internal MarkdownNativeBlockSourceFieldSnapshot(MarkdownNativeBlockSourceField field) {
        Name = field.Name;
        Value = field.Value;
        SourceSpan = new MarkdownNativeSourceSpanSnapshot(field.SourceSpan);
        Index = field.Index;
    }

    /// <summary>Stable field name such as <c>level</c>, <c>infoString</c>, or <c>quoteMarker</c>.</summary>
    public string Name { get; }

    /// <summary>Semantic value represented by the field when one is available.</summary>
    public string? Value { get; }

    /// <summary>Source span for this field.</summary>
    public MarkdownNativeSourceSpanSnapshot SourceSpan { get; }

    /// <summary>Zero-based occurrence index for repeated fields, or <c>-1</c> for singular fields.</summary>
    public int Index { get; }
}

/// <summary>
/// UI-safe snapshot of a definition-list group.
/// </summary>
public sealed class MarkdownNativeDefinitionListGroupSnapshot {
    internal MarkdownNativeDefinitionListGroupSnapshot(
        MarkdownNativeSourceSpanSnapshot? sourceSpan,
        IReadOnlyList<MarkdownNativeDefinitionListTermSnapshot> terms,
        IReadOnlyList<MarkdownNativeDefinitionListDefinitionSnapshot> definitions) {
        SourceSpan = sourceSpan;
        Terms = terms ?? Array.Empty<MarkdownNativeDefinitionListTermSnapshot>();
        Definitions = definitions ?? Array.Empty<MarkdownNativeDefinitionListDefinitionSnapshot>();
    }

    /// <summary>Source span snapshot when available.</summary>
    public MarkdownNativeSourceSpanSnapshot? SourceSpan { get; }

    /// <summary>Terms in this group.</summary>
    public IReadOnlyList<MarkdownNativeDefinitionListTermSnapshot> Terms { get; }

    /// <summary>Definitions in this group.</summary>
    public IReadOnlyList<MarkdownNativeDefinitionListDefinitionSnapshot> Definitions { get; }
}

/// <summary>
/// UI-safe snapshot of a definition-list term.
/// </summary>
public sealed class MarkdownNativeDefinitionListTermSnapshot {
    internal MarkdownNativeDefinitionListTermSnapshot(
        string text,
        string markdown,
        MarkdownNativeSourceSpanSnapshot? sourceSpan,
        IReadOnlyList<MarkdownNativeInlineSnapshot> inlines) {
        Text = text ?? string.Empty;
        Markdown = markdown ?? string.Empty;
        SourceSpan = sourceSpan;
        Inlines = inlines ?? Array.Empty<MarkdownNativeInlineSnapshot>();
    }

    /// <summary>Plain text term content.</summary>
    public string Text { get; }

    /// <summary>Markdown term content.</summary>
    public string Markdown { get; }

    /// <summary>Source span snapshot when available.</summary>
    public MarkdownNativeSourceSpanSnapshot? SourceSpan { get; }

    /// <summary>Inline snapshots for the term.</summary>
    public IReadOnlyList<MarkdownNativeInlineSnapshot> Inlines { get; }
}

/// <summary>
/// UI-safe snapshot of a definition-list definition body.
/// </summary>
public sealed class MarkdownNativeDefinitionListDefinitionSnapshot {
    internal MarkdownNativeDefinitionListDefinitionSnapshot(
        string markdown,
        MarkdownNativeSourceSpanSnapshot? sourceSpan,
        IReadOnlyList<MarkdownNativeBlockSnapshot> children) {
        Markdown = markdown ?? string.Empty;
        SourceSpan = sourceSpan;
        Children = children ?? Array.Empty<MarkdownNativeBlockSnapshot>();
    }

    /// <summary>Markdown definition body.</summary>
    public string Markdown { get; }

    /// <summary>Source span snapshot when available.</summary>
    public MarkdownNativeSourceSpanSnapshot? SourceSpan { get; }

    /// <summary>Nested child block snapshots.</summary>
    public IReadOnlyList<MarkdownNativeBlockSnapshot> Children { get; }
}

/// <summary>
/// UI-safe snapshot of a native inline.
/// </summary>
public sealed class MarkdownNativeInlineSnapshot {
    internal MarkdownNativeInlineSnapshot(
        string id,
        MarkdownNativeInlineKind kind,
        MarkdownSyntaxKind syntaxKind,
        string text,
        string markdown,
        string literal,
        MarkdownNativeSourceSpanSnapshot? sourceSpan,
        IReadOnlyDictionary<string, string> metadata,
        IReadOnlyDictionary<string, MarkdownNativeSourceSpanSnapshot?> metadataSourceSpans,
        IReadOnlyList<MarkdownNativeInlineSnapshot> children) {
        Id = id ?? string.Empty;
        Kind = kind;
        SyntaxKind = syntaxKind;
        Text = text ?? string.Empty;
        Markdown = markdown ?? string.Empty;
        Literal = literal ?? string.Empty;
        SourceSpan = sourceSpan;
        Metadata = metadata ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        MetadataSourceSpans = metadataSourceSpans ?? new Dictionary<string, MarkdownNativeSourceSpanSnapshot?>(StringComparer.OrdinalIgnoreCase);
        Children = children ?? Array.Empty<MarkdownNativeInlineSnapshot>();
    }

    /// <summary>Stable inline id.</summary>
    public string Id { get; }

    /// <summary>Native inline kind.</summary>
    public MarkdownNativeInlineKind Kind { get; }

    /// <summary>Syntax kind that produced this inline.</summary>
    public MarkdownSyntaxKind SyntaxKind { get; }

    /// <summary>Plain text represented by this inline.</summary>
    public string Text { get; }

    /// <summary>Markdown represented by this inline.</summary>
    public string Markdown { get; }

    /// <summary>Literal syntax payload.</summary>
    public string Literal { get; }

    /// <summary>Source span snapshot when available.</summary>
    public MarkdownNativeSourceSpanSnapshot? SourceSpan { get; }

    /// <summary>Metadata values such as target/title/source/alt.</summary>
    public IReadOnlyDictionary<string, string> Metadata { get; }

    /// <summary>Source spans for metadata values such as target/title/source/alt.</summary>
    public IReadOnlyDictionary<string, MarkdownNativeSourceSpanSnapshot?> MetadataSourceSpans { get; }

    /// <summary>Nested inline snapshots.</summary>
    public IReadOnlyList<MarkdownNativeInlineSnapshot> Children { get; }
}

/// <summary>
/// UI-safe snapshot of a native list item.
/// </summary>
public sealed class MarkdownNativeListItemSnapshot {
    internal MarkdownNativeListItemSnapshot(
        string id,
        string text,
        bool isTask,
        bool isChecked,
        int level,
        MarkdownNativeSourceSpanSnapshot? sourceSpan,
        MarkdownNativeSourceSpanSnapshot? markerSourceSpan,
        string? markerText,
        MarkdownNativeSourceSpanSnapshot? taskMarkerSourceSpan,
        string? taskMarkerText,
        IReadOnlyList<MarkdownNativeInlineSnapshot> inlines,
        IReadOnlyList<MarkdownNativeListItemParagraphSnapshot> paragraphs,
        IReadOnlyList<MarkdownNativeBlockSnapshot> children) {
        Id = id ?? string.Empty;
        Text = text ?? string.Empty;
        IsTask = isTask;
        IsChecked = isChecked;
        Level = level;
        SourceSpan = sourceSpan;
        MarkerSourceSpan = markerSourceSpan;
        MarkerText = markerText;
        TaskMarkerSourceSpan = taskMarkerSourceSpan;
        TaskMarkerText = taskMarkerText;
        Inlines = inlines ?? Array.Empty<MarkdownNativeInlineSnapshot>();
        Paragraphs = paragraphs ?? Array.Empty<MarkdownNativeListItemParagraphSnapshot>();
        Children = children ?? Array.Empty<MarkdownNativeBlockSnapshot>();
    }

    /// <summary>Stable list item id.</summary>
    public string Id { get; }

    /// <summary>Plain text lead content.</summary>
    public string Text { get; }

    /// <summary>Whether the item is a task item.</summary>
    public bool IsTask { get; }

    /// <summary>Whether the task item is checked.</summary>
    public bool IsChecked { get; }

    /// <summary>Indentation level from the source item.</summary>
    public int Level { get; }

    /// <summary>Source span snapshot when available.</summary>
    public MarkdownNativeSourceSpanSnapshot? SourceSpan { get; }

    /// <summary>List marker token source span snapshot when available.</summary>
    public MarkdownNativeSourceSpanSnapshot? MarkerSourceSpan { get; }

    /// <summary>Exact list marker token when available.</summary>
    public string? MarkerText { get; }

    /// <summary>Task marker token source span snapshot when available.</summary>
    public MarkdownNativeSourceSpanSnapshot? TaskMarkerSourceSpan { get; }

    /// <summary>Exact task marker token when available.</summary>
    public string? TaskMarkerText { get; }

    /// <summary>Lead inline snapshots.</summary>
    public IReadOnlyList<MarkdownNativeInlineSnapshot> Inlines { get; }

    /// <summary>Paragraph-level snapshots owned by this list item.</summary>
    public IReadOnlyList<MarkdownNativeListItemParagraphSnapshot> Paragraphs { get; }

    /// <summary>Nested child block snapshots.</summary>
    public IReadOnlyList<MarkdownNativeBlockSnapshot> Children { get; }
}

/// <summary>
/// UI-safe snapshot of one paragraph owned by a native list item.
/// </summary>
public sealed class MarkdownNativeListItemParagraphSnapshot {
    internal MarkdownNativeListItemParagraphSnapshot(
        int index,
        string text,
        MarkdownNativeSourceSpanSnapshot? sourceSpan,
        IReadOnlyList<MarkdownNativeInlineSnapshot> inlines) {
        Index = index;
        Text = text ?? string.Empty;
        SourceSpan = sourceSpan;
        Inlines = inlines ?? Array.Empty<MarkdownNativeInlineSnapshot>();
    }

    /// <summary>Zero-based paragraph index within the list item.</summary>
    public int Index { get; }

    /// <summary>Plain-text paragraph content.</summary>
    public string Text { get; }

    /// <summary>Source span snapshot when available.</summary>
    public MarkdownNativeSourceSpanSnapshot? SourceSpan { get; }

    /// <summary>Inline snapshots for this paragraph.</summary>
    public IReadOnlyList<MarkdownNativeInlineSnapshot> Inlines { get; }
}

/// <summary>
/// UI-safe snapshot of a native table cell.
/// </summary>
public sealed class MarkdownNativeTableCellSnapshot {
    internal MarkdownNativeTableCellSnapshot(
        string text,
        string markdown,
        bool isHeader,
        int rowIndex,
        int columnIndex,
        ColumnAlignment alignment,
        MarkdownNativeSourceSpanSnapshot? sourceSpan,
        IReadOnlyList<MarkdownNativeInlineSnapshot> inlines,
        IReadOnlyList<MarkdownNativeBlockSnapshot> children) {
        Text = text ?? string.Empty;
        Markdown = markdown ?? string.Empty;
        IsHeader = isHeader;
        RowIndex = rowIndex;
        ColumnIndex = columnIndex;
        Alignment = alignment;
        SourceSpan = sourceSpan;
        Inlines = inlines ?? Array.Empty<MarkdownNativeInlineSnapshot>();
        Children = children ?? Array.Empty<MarkdownNativeBlockSnapshot>();
    }

    /// <summary>Plain text cell content.</summary>
    public string Text { get; }

    /// <summary>Markdown cell content.</summary>
    public string Markdown { get; }

    /// <summary>Whether this is a header cell.</summary>
    public bool IsHeader { get; }

    /// <summary>Zero-based row index, or -1 for headers.</summary>
    public int RowIndex { get; }

    /// <summary>Zero-based column index.</summary>
    public int ColumnIndex { get; }

    /// <summary>Projected alignment.</summary>
    public ColumnAlignment Alignment { get; }

    /// <summary>Source span snapshot when available.</summary>
    public MarkdownNativeSourceSpanSnapshot? SourceSpan { get; }

    /// <summary>Inline snapshots for cell content when available.</summary>
    public IReadOnlyList<MarkdownNativeInlineSnapshot> Inlines { get; }

    /// <summary>Native child block snapshots projected from structured cell content.</summary>
    public IReadOnlyList<MarkdownNativeBlockSnapshot> Children { get; }
}

/// <summary>
/// UI-safe source span snapshot.
/// </summary>
public sealed class MarkdownNativeSourceSpanSnapshot {
    internal MarkdownNativeSourceSpanSnapshot(MarkdownSourceSpan span) {
        StartLine = span.StartLine;
        StartColumn = span.StartColumn;
        EndLine = span.EndLine;
        EndColumn = span.EndColumn;
        StartOffset = span.StartOffset;
        EndOffset = span.EndOffset;
        Display = span.ToString();
    }

    /// <summary>1-based start line.</summary>
    public int StartLine { get; }

    /// <summary>1-based start column when available.</summary>
    public int? StartColumn { get; }

    /// <summary>1-based end line.</summary>
    public int EndLine { get; }

    /// <summary>1-based end column when available.</summary>
    public int? EndColumn { get; }

    /// <summary>0-based start offset when available.</summary>
    public int? StartOffset { get; }

    /// <summary>0-based end offset when available.</summary>
    public int? EndOffset { get; }

    /// <summary>Human-readable span display.</summary>
    public string Display { get; }
}

/// <summary>
/// UI-safe diagnostic snapshot.
/// </summary>
public sealed class MarkdownNativeDiagnosticSnapshot {
    internal MarkdownNativeDiagnosticSnapshot(MarkdownNativeDiagnostic diagnostic) {
        Id = diagnostic.Id;
        Message = diagnostic.Message;
        Severity = diagnostic.Severity;
        SourceSpan = diagnostic.SourceSpan.HasValue ? new MarkdownNativeSourceSpanSnapshot(diagnostic.SourceSpan.Value) : null;
        BlockId = diagnostic.Block?.Id;
        RelatedSourceSpans = ToSourceSpanSnapshots(diagnostic.RelatedSourceSpans);
    }

    /// <summary>Diagnostic id.</summary>
    public string Id { get; }

    /// <summary>Diagnostic message.</summary>
    public string Message { get; }

    /// <summary>Diagnostic severity.</summary>
    public MarkdownNativeDiagnosticSeverity Severity { get; }

    /// <summary>Source span snapshot when available.</summary>
    public MarkdownNativeSourceSpanSnapshot? SourceSpan { get; }

    /// <summary>Additional related source spans, such as individual transform input blocks.</summary>
    public IReadOnlyList<MarkdownNativeSourceSpanSnapshot> RelatedSourceSpans { get; }

    /// <summary>Associated block id when available.</summary>
    public string? BlockId { get; }

    private static IReadOnlyList<MarkdownNativeSourceSpanSnapshot> ToSourceSpanSnapshots(IReadOnlyList<MarkdownSourceSpan> spans) {
        if (spans == null || spans.Count == 0) {
            return Array.Empty<MarkdownNativeSourceSpanSnapshot>();
        }

        var snapshots = new List<MarkdownNativeSourceSpanSnapshot>(spans.Count);
        for (var i = 0; i < spans.Count; i++) {
            snapshots.Add(new MarkdownNativeSourceSpanSnapshot(spans[i]));
        }

        return snapshots;
    }
}
