namespace OfficeIMO.Markdown;

/// <summary>
/// UI-safe snapshot of a native markdown document without parser object references.
/// </summary>
public sealed class MarkdownNativeDocumentInfo {
    internal MarkdownNativeDocumentInfo(
        MarkdownNativeDocumentSourceKind sourceKind,
        IReadOnlyList<MarkdownNativeReferenceLinkDefinitionInfo> referenceLinkDefinitions,
        IReadOnlyList<MarkdownNativeAbbreviationDefinitionInfo> abbreviationDefinitions,
        IReadOnlyList<MarkdownNativeSourceTriviaInfo> sourceTrivia,
        IReadOnlyList<MarkdownNativeBlockInfo> blocks,
        IReadOnlyList<MarkdownNativeDiagnosticInfo> diagnostics) {
        SourceKind = sourceKind;
        ReferenceLinkDefinitions = referenceLinkDefinitions ?? Array.Empty<MarkdownNativeReferenceLinkDefinitionInfo>();
        AbbreviationDefinitions = abbreviationDefinitions ?? Array.Empty<MarkdownNativeAbbreviationDefinitionInfo>();
        SourceTrivia = sourceTrivia ?? Array.Empty<MarkdownNativeSourceTriviaInfo>();
        Blocks = blocks ?? Array.Empty<MarkdownNativeBlockInfo>();
        Diagnostics = diagnostics ?? Array.Empty<MarkdownNativeDiagnosticInfo>();
    }

    /// <summary>Identifies the markdown source backing this snapshot.</summary>
    public MarkdownNativeDocumentSourceKind SourceKind { get; }

    /// <summary>Effective reference-style link definitions collected during parsing.</summary>
    public IReadOnlyList<MarkdownNativeReferenceLinkDefinitionInfo> ReferenceLinkDefinitions { get; }

    /// <summary>Effective abbreviation definitions collected during parsing.</summary>
    public IReadOnlyList<MarkdownNativeAbbreviationDefinitionInfo> AbbreviationDefinitions { get; }

    /// <summary>Document-level source trivia such as blank lines, in source order.</summary>
    public IReadOnlyList<MarkdownNativeSourceTriviaInfo> SourceTrivia { get; }

    /// <summary>Top-level block snapshots.</summary>
    public IReadOnlyList<MarkdownNativeBlockInfo> Blocks { get; }

    /// <summary>Projection diagnostics.</summary>
    public IReadOnlyList<MarkdownNativeDiagnosticInfo> Diagnostics { get; }
}

/// <summary>
/// UI-safe snapshot of document-level source trivia.
/// </summary>
public sealed class MarkdownNativeSourceTriviaInfo {
    internal MarkdownNativeSourceTriviaInfo(MarkdownNativeSourceTrivia trivia) {
        Kind = trivia.Kind;
        Text = trivia.Text;
        SourceSpan = new MarkdownNativeSourceSpanInfo(trivia.SourceSpan);
    }

    internal MarkdownNativeSourceTriviaInfo(
        MarkdownNativeSourceTrivia trivia,
        string? sourceText,
        string? originalSourceText,
        MarkdownOriginalSourceSliceFailureReason? originalSourceFailureReason) {
        Kind = trivia.Kind;
        Text = trivia.Text;
        SourceSpan = new MarkdownNativeSourceSpanInfo(trivia.SourceSpan);
        SourceText = sourceText;
        OriginalSourceText = originalSourceText;
        OriginalSourceFailureReason = originalSourceFailureReason;
    }

    /// <summary>Trivia kind.</summary>
    public MarkdownNativeSourceTriviaKind Kind { get; }

    /// <summary>Exact normalized line content represented by this trivia, excluding the line ending.</summary>
    public string Text { get; }

    /// <summary>Source span for the trivia content.</summary>
    public MarkdownNativeSourceSpanInfo SourceSpan { get; }

    /// <summary>Exact normalized source text represented by this trivia when it could be materialized.</summary>
    public string? SourceText { get; }

    /// <summary>Exact original reader input represented by this trivia when trivia was preserved and mapping succeeded.</summary>
    public string? OriginalSourceText { get; }

    /// <summary>Reason original reader input could not be materialized for this trivia, when applicable.</summary>
    public MarkdownOriginalSourceSliceFailureReason? OriginalSourceFailureReason { get; }
}

/// <summary>
/// UI-safe snapshot of a reference-style link definition.
/// </summary>
public sealed class MarkdownNativeReferenceLinkDefinitionInfo {
    internal MarkdownNativeReferenceLinkDefinitionInfo(MarkdownReferenceLinkDefinition definition)
        : this(null, definition) {
    }

    internal MarkdownNativeReferenceLinkDefinitionInfo(MarkdownNativeDocument? document, MarkdownReferenceLinkDefinition definition) {
        Label = definition.Label;
        Url = definition.Url;
        Title = definition.Title;
        SourceSpan = definition.SourceSpan.HasValue ? new MarkdownNativeSourceSpanInfo(definition.SourceSpan.Value) : null;
        LabelSourceSpan = definition.LabelSourceSpan.HasValue ? new MarkdownNativeSourceSpanInfo(definition.LabelSourceSpan.Value) : null;
        OpeningMarkerSourceSpan = definition.OpeningMarkerSourceSpan.HasValue ? new MarkdownNativeSourceSpanInfo(definition.OpeningMarkerSourceSpan.Value) : null;
        SeparatorMarkerSourceSpan = definition.SeparatorMarkerSourceSpan.HasValue ? new MarkdownNativeSourceSpanInfo(definition.SeparatorMarkerSourceSpan.Value) : null;
        UrlSourceSpan = definition.UrlSourceSpan.HasValue ? new MarkdownNativeSourceSpanInfo(definition.UrlSourceSpan.Value) : null;
        TitleSourceSpan = definition.TitleSourceSpan.HasValue ? new MarkdownNativeSourceSpanInfo(definition.TitleSourceSpan.Value) : null;
        SourceFields = FromReferenceDefinitionFields(document, definition);
    }

    /// <summary>Normalized reference label.</summary>
    public string Label { get; }

    /// <summary>Resolved destination URL.</summary>
    public string Url { get; }

    /// <summary>Optional definition title.</summary>
    public string? Title { get; }

    /// <summary>Source span for the entire reference-style definition, when available.</summary>
    public MarkdownNativeSourceSpanInfo? SourceSpan { get; }

    /// <summary>Source span for the definition label token, when available.</summary>
    public MarkdownNativeSourceSpanInfo? LabelSourceSpan { get; }

    /// <summary>Source span for the opening <c>[</c> marker before the definition label, when available.</summary>
    public MarkdownNativeSourceSpanInfo? OpeningMarkerSourceSpan { get; }

    /// <summary>Source span for the <c>]:</c> marker after the definition label, when available.</summary>
    public MarkdownNativeSourceSpanInfo? SeparatorMarkerSourceSpan { get; }

    /// <summary>Source span for the destination token, when available.</summary>
    public MarkdownNativeSourceSpanInfo? UrlSourceSpan { get; }

    /// <summary>Source span for the optional title token, when available.</summary>
    public MarkdownNativeSourceSpanInfo? TitleSourceSpan { get; }

    /// <summary>Source-backed token and payload fields in source order.</summary>
    public IReadOnlyList<MarkdownNativeReferenceLinkDefinitionFieldInfo> SourceFields { get; }

    private static IReadOnlyList<MarkdownNativeReferenceLinkDefinitionFieldInfo> FromReferenceDefinitionFields(
        MarkdownNativeDocument? document,
        MarkdownReferenceLinkDefinition definition) {
        var fields = MarkdownNativeDocument.EnumerateReferenceLinkDefinitionFields(definition).ToArray();
        if (fields.Length == 0) {
            return Array.Empty<MarkdownNativeReferenceLinkDefinitionFieldInfo>();
        }

        var snapshots = new List<MarkdownNativeReferenceLinkDefinitionFieldInfo>(fields.Length);
        for (var i = 0; i < fields.Length; i++) {
            snapshots.Add(new MarkdownNativeReferenceLinkDefinitionFieldInfo(document, fields[i]));
        }

        return snapshots;
    }
}

/// <summary>
/// UI-safe snapshot of a source-backed token or payload field owned by a reference-style link definition.
/// </summary>
public sealed class MarkdownNativeReferenceLinkDefinitionFieldInfo {
    internal MarkdownNativeReferenceLinkDefinitionFieldInfo(MarkdownNativeReferenceLinkDefinitionField field)
        : this(null, field) {
    }

    internal MarkdownNativeReferenceLinkDefinitionFieldInfo(MarkdownNativeDocument? document, MarkdownNativeReferenceLinkDefinitionField field) {
        Name = field.Name;
        Value = field.Value;
        SourceSpan = new MarkdownNativeSourceSpanInfo(field.SourceSpan);

        if (document == null) {
            return;
        }

        if (document.TryCreateSourceSlice(field, out var sourceSlice)) {
            SourceText = sourceSlice.Text;
        }

        if (document.TryCreateOriginalSourceSlice(field, out var originalSlice, out var failureReason)) {
            OriginalSourceText = originalSlice.Text;
        } else {
            OriginalSourceFailureReason = failureReason;
        }
    }

    /// <summary>Stable field name such as <c>openingMarker</c>, <c>label</c>, <c>separatorMarker</c>, <c>url</c>, or <c>title</c>.</summary>
    public string Name { get; }

    /// <summary>Semantic value represented by the field when one is available.</summary>
    public string? Value { get; }

    /// <summary>Source span for this field.</summary>
    public MarkdownNativeSourceSpanInfo SourceSpan { get; }

    /// <summary>Exact normalized source text represented by this field when it could be materialized.</summary>
    public string? SourceText { get; }

    /// <summary>Exact original reader input represented by this field when trivia was preserved and mapping succeeded.</summary>
    public string? OriginalSourceText { get; }

    /// <summary>Reason original reader input could not be materialized for this field, when applicable.</summary>
    public MarkdownOriginalSourceSliceFailureReason? OriginalSourceFailureReason { get; }
}

/// <summary>
/// UI-safe snapshot of a Markdig-style abbreviation definition.
/// </summary>
public sealed class MarkdownNativeAbbreviationDefinitionInfo {
    internal MarkdownNativeAbbreviationDefinitionInfo(MarkdownAbbreviationDefinition definition)
        : this(null, definition) {
    }

    internal MarkdownNativeAbbreviationDefinitionInfo(MarkdownNativeDocument? document, MarkdownAbbreviationDefinition definition) {
        Label = definition.Label;
        Title = definition.Title;
        IsListItemDefinition = definition.IsListItemDefinition;
        SourceSpan = definition.SourceSpan.HasValue ? new MarkdownNativeSourceSpanInfo(definition.SourceSpan.Value) : null;
        LabelSourceSpan = definition.LabelSourceSpan.HasValue ? new MarkdownNativeSourceSpanInfo(definition.LabelSourceSpan.Value) : null;
        OpeningMarkerSourceSpan = definition.OpeningMarkerSourceSpan.HasValue ? new MarkdownNativeSourceSpanInfo(definition.OpeningMarkerSourceSpan.Value) : null;
        SeparatorMarkerSourceSpan = definition.SeparatorMarkerSourceSpan.HasValue ? new MarkdownNativeSourceSpanInfo(definition.SeparatorMarkerSourceSpan.Value) : null;
        TitleSourceSpan = definition.TitleSourceSpan.HasValue ? new MarkdownNativeSourceSpanInfo(definition.TitleSourceSpan.Value) : null;
        SourceFields = FromAbbreviationDefinitionFields(document, definition);
    }

    /// <summary>Abbreviation label.</summary>
    public string Label { get; }

    /// <summary>Abbreviation title.</summary>
    public string Title { get; }

    /// <summary>Whether the definition was authored as leading content inside a list item.</summary>
    public bool IsListItemDefinition { get; }

    /// <summary>Source span for the entire abbreviation definition, when available.</summary>
    public MarkdownNativeSourceSpanInfo? SourceSpan { get; }

    /// <summary>Source span for the abbreviation label token, when available.</summary>
    public MarkdownNativeSourceSpanInfo? LabelSourceSpan { get; }

    /// <summary>Source span for the opening <c>*[</c> marker, when available.</summary>
    public MarkdownNativeSourceSpanInfo? OpeningMarkerSourceSpan { get; }

    /// <summary>Source span for the closing <c>]:</c> marker, when available.</summary>
    public MarkdownNativeSourceSpanInfo? SeparatorMarkerSourceSpan { get; }

    /// <summary>Source span for the title token, when available.</summary>
    public MarkdownNativeSourceSpanInfo? TitleSourceSpan { get; }

    /// <summary>Source-backed token and payload fields in source order.</summary>
    public IReadOnlyList<MarkdownNativeAbbreviationDefinitionFieldInfo> SourceFields { get; }

    private static IReadOnlyList<MarkdownNativeAbbreviationDefinitionFieldInfo> FromAbbreviationDefinitionFields(
        MarkdownNativeDocument? document,
        MarkdownAbbreviationDefinition definition) {
        var fields = MarkdownNativeDocument.EnumerateAbbreviationDefinitionFields(definition).ToArray();
        if (fields.Length == 0) {
            return Array.Empty<MarkdownNativeAbbreviationDefinitionFieldInfo>();
        }

        var snapshots = new List<MarkdownNativeAbbreviationDefinitionFieldInfo>(fields.Length);
        for (var i = 0; i < fields.Length; i++) {
            snapshots.Add(new MarkdownNativeAbbreviationDefinitionFieldInfo(document, fields[i]));
        }

        return snapshots;
    }
}

/// <summary>
/// UI-safe snapshot of a source-backed token or payload field owned by an abbreviation definition.
/// </summary>
public sealed class MarkdownNativeAbbreviationDefinitionFieldInfo {
    internal MarkdownNativeAbbreviationDefinitionFieldInfo(MarkdownNativeAbbreviationDefinitionField field)
        : this(null, field) {
    }

    internal MarkdownNativeAbbreviationDefinitionFieldInfo(MarkdownNativeDocument? document, MarkdownNativeAbbreviationDefinitionField field) {
        Name = field.Name;
        Value = field.Value;
        SourceSpan = new MarkdownNativeSourceSpanInfo(field.SourceSpan);

        if (document == null) {
            return;
        }

        if (document.TryCreateSourceSlice(field, out var sourceSlice)) {
            SourceText = sourceSlice.Text;
        }

        if (document.TryCreateOriginalSourceSlice(field, out var originalSlice, out var failureReason)) {
            OriginalSourceText = originalSlice.Text;
        } else {
            OriginalSourceFailureReason = failureReason;
        }
    }

    /// <summary>Stable field name such as <c>openingMarker</c>, <c>label</c>, <c>separatorMarker</c>, or <c>title</c>.</summary>
    public string Name { get; }

    /// <summary>Semantic value represented by the field when one is available.</summary>
    public string? Value { get; }

    /// <summary>Source span for this field.</summary>
    public MarkdownNativeSourceSpanInfo SourceSpan { get; }

    /// <summary>Exact normalized source text represented by this field when it could be materialized.</summary>
    public string? SourceText { get; }

    /// <summary>Exact original reader input represented by this field when trivia was preserved and mapping succeeded.</summary>
    public string? OriginalSourceText { get; }

    /// <summary>Reason original reader input could not be materialized for this field, when applicable.</summary>
    public MarkdownOriginalSourceSliceFailureReason? OriginalSourceFailureReason { get; }
}

/// <summary>
/// UI-safe snapshot of a native block.
/// </summary>
public sealed class MarkdownNativeBlockInfo {
    internal MarkdownNativeBlockInfo() {
        Fields = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);
        FieldSourceSpans = new Dictionary<string, MarkdownNativeSourceSpanInfo?>(StringComparer.OrdinalIgnoreCase);
        SourceFields = Array.Empty<MarkdownNativeBlockSourceFieldInfo>();
        MarkerSourceSpans = Array.Empty<MarkdownNativeSourceSpanInfo>();
        Inlines = Array.Empty<MarkdownNativeInlineInfo>();
        Children = Array.Empty<MarkdownNativeBlockInfo>();
        Items = Array.Empty<MarkdownNativeListItemInfo>();
        DefinitionGroups = Array.Empty<MarkdownNativeDefinitionListGroupInfo>();
        HeaderCells = Array.Empty<MarkdownNativeTableCellInfo>();
        Rows = Array.Empty<IReadOnlyList<MarkdownNativeTableCellInfo>>();
        BodyRows = Array.Empty<MarkdownNativeTableRowInfo>();
    }

    /// <summary>Stable block id.</summary>
    public string Id { get; internal set; } = string.Empty;

    /// <summary>Native block kind.</summary>
    public MarkdownNativeBlockKind Kind { get; internal set; }

    /// <summary>Source span snapshot when available.</summary>
    public MarkdownNativeSourceSpanInfo? SourceSpan { get; internal set; }

    /// <summary>Common text payload when the block exposes one.</summary>
    public string? Text { get; internal set; }

    /// <summary>Common markdown payload when the block exposes one.</summary>
    public string? Markdown { get; internal set; }

    /// <summary>String fields for block-specific metadata.</summary>
    public IReadOnlyDictionary<string, string?> Fields { get; internal set; }

    /// <summary>Source spans for block-specific metadata fields when available.</summary>
    public IReadOnlyDictionary<string, MarkdownNativeSourceSpanInfo?> FieldSourceSpans { get; internal set; }

    /// <summary>Source-backed token and payload fields in source order, including repeated fields.</summary>
    public IReadOnlyList<MarkdownNativeBlockSourceFieldInfo> SourceFields { get; internal set; }

    /// <summary>Source spans for repeated marker tokens owned by this block, such as blockquote markers.</summary>
    public IReadOnlyList<MarkdownNativeSourceSpanInfo> MarkerSourceSpans { get; internal set; }

    /// <summary>Inline snapshots owned directly by this block.</summary>
    public IReadOnlyList<MarkdownNativeInlineInfo> Inlines { get; internal set; }

    /// <summary>Nested child block snapshots.</summary>
    public IReadOnlyList<MarkdownNativeBlockInfo> Children { get; internal set; }

    /// <summary>List item snapshots for native list blocks.</summary>
    public IReadOnlyList<MarkdownNativeListItemInfo> Items { get; internal set; }

    /// <summary>Definition-list group snapshots for native definition list blocks.</summary>
    public IReadOnlyList<MarkdownNativeDefinitionListGroupInfo> DefinitionGroups { get; internal set; }

    /// <summary>MarkdownTable header cell snapshots.</summary>
    public IReadOnlyList<MarkdownNativeTableCellInfo> HeaderCells { get; internal set; }

    /// <summary>MarkdownTable header row snapshot when available.</summary>
    public MarkdownNativeTableRowInfo? HeaderRow { get; internal set; }

    /// <summary>MarkdownTable body row snapshots.</summary>
    public IReadOnlyList<IReadOnlyList<MarkdownNativeTableCellInfo>> Rows { get; internal set; }

    /// <summary>MarkdownTable body row object snapshots.</summary>
    public IReadOnlyList<MarkdownNativeTableRowInfo> BodyRows { get; internal set; }

    /// <summary>Enumerates source-backed fields with the supplied field name in source order.</summary>
    public IEnumerable<MarkdownNativeBlockSourceFieldInfo> EnumerateSourceFields(string name) {
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
    public MarkdownNativeBlockSourceFieldInfo? FindSourceField(string name, int index = -1) {
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
public sealed class MarkdownNativeBlockSourceFieldInfo {
    internal MarkdownNativeBlockSourceFieldInfo(MarkdownNativeBlockSourceField field) {
        Name = field.Name;
        Value = field.Value;
        SourceSpan = new MarkdownNativeSourceSpanInfo(field.SourceSpan);
        Index = field.Index;
    }

    internal MarkdownNativeBlockSourceFieldInfo(
        MarkdownNativeBlockSourceField field,
        string? sourceText,
        string? originalSourceText,
        MarkdownOriginalSourceSliceFailureReason? originalSourceFailureReason) {
        Name = field.Name;
        Value = field.Value;
        SourceSpan = new MarkdownNativeSourceSpanInfo(field.SourceSpan);
        Index = field.Index;
        SourceText = sourceText;
        OriginalSourceText = originalSourceText;
        OriginalSourceFailureReason = originalSourceFailureReason;
    }

    /// <summary>Stable field name such as <c>level</c>, <c>infoString</c>, or <c>quoteMarker</c>.</summary>
    public string Name { get; }

    /// <summary>Semantic value represented by the field when one is available.</summary>
    public string? Value { get; }

    /// <summary>Source span for this field.</summary>
    public MarkdownNativeSourceSpanInfo SourceSpan { get; }

    /// <summary>Zero-based occurrence index for repeated fields, or <c>-1</c> for singular fields.</summary>
    public int Index { get; }

    /// <summary>Exact normalized source text represented by the field when it could be materialized.</summary>
    public string? SourceText { get; }

    /// <summary>Exact original reader input represented by the field when trivia was preserved and mapping succeeded.</summary>
    public string? OriginalSourceText { get; }

    /// <summary>Reason original reader input could not be materialized for this field, when applicable.</summary>
    public MarkdownOriginalSourceSliceFailureReason? OriginalSourceFailureReason { get; }
}

/// <summary>
/// UI-safe snapshot of a definition-list group.
/// </summary>
public sealed class MarkdownNativeDefinitionListGroupInfo {
    internal MarkdownNativeDefinitionListGroupInfo(
        MarkdownNativeSourceSpanInfo? sourceSpan,
        string? sourceText,
        string? originalSourceText,
        MarkdownOriginalSourceSliceFailureReason? originalSourceFailureReason,
        IReadOnlyList<MarkdownNativeDefinitionListTermInfo> terms,
        IReadOnlyList<MarkdownNativeDefinitionListDefinitionInfo> definitions) {
        SourceSpan = sourceSpan;
        SourceText = sourceText;
        OriginalSourceText = originalSourceText;
        OriginalSourceFailureReason = originalSourceFailureReason;
        Terms = terms ?? Array.Empty<MarkdownNativeDefinitionListTermInfo>();
        Definitions = definitions ?? Array.Empty<MarkdownNativeDefinitionListDefinitionInfo>();
    }

    /// <summary>Source span snapshot when available.</summary>
    public MarkdownNativeSourceSpanInfo? SourceSpan { get; }

    /// <summary>Normalized markdown text backing this definition-list group when available.</summary>
    public string? SourceText { get; }

    /// <summary>Original reader input backing this definition-list group when available.</summary>
    public string? OriginalSourceText { get; }

    /// <summary>Reason original reader input could not be materialized for this group, when applicable.</summary>
    public MarkdownOriginalSourceSliceFailureReason? OriginalSourceFailureReason { get; }

    /// <summary>Terms in this group.</summary>
    public IReadOnlyList<MarkdownNativeDefinitionListTermInfo> Terms { get; }

    /// <summary>Definitions in this group.</summary>
    public IReadOnlyList<MarkdownNativeDefinitionListDefinitionInfo> Definitions { get; }
}

/// <summary>
/// UI-safe snapshot of a definition-list term.
/// </summary>
public sealed class MarkdownNativeDefinitionListTermInfo {
    internal MarkdownNativeDefinitionListTermInfo(
        string text,
        string markdown,
        MarkdownNativeSourceSpanInfo? sourceSpan,
        string? sourceText,
        string? originalSourceText,
        MarkdownOriginalSourceSliceFailureReason? originalSourceFailureReason,
        IReadOnlyList<MarkdownNativeInlineInfo> inlines) {
        Text = text ?? string.Empty;
        Markdown = markdown ?? string.Empty;
        SourceSpan = sourceSpan;
        SourceText = sourceText;
        OriginalSourceText = originalSourceText;
        OriginalSourceFailureReason = originalSourceFailureReason;
        Inlines = inlines ?? Array.Empty<MarkdownNativeInlineInfo>();
    }

    /// <summary>Plain text term content.</summary>
    public string Text { get; }

    /// <summary>Markdown term content.</summary>
    public string Markdown { get; }

    /// <summary>Source span snapshot when available.</summary>
    public MarkdownNativeSourceSpanInfo? SourceSpan { get; }

    /// <summary>Normalized markdown text backing this definition-list term when available.</summary>
    public string? SourceText { get; }

    /// <summary>Original reader input backing this definition-list term when available.</summary>
    public string? OriginalSourceText { get; }

    /// <summary>Reason original reader input could not be materialized for this term, when applicable.</summary>
    public MarkdownOriginalSourceSliceFailureReason? OriginalSourceFailureReason { get; }

    /// <summary>Inline snapshots for the term.</summary>
    public IReadOnlyList<MarkdownNativeInlineInfo> Inlines { get; }
}

/// <summary>
/// UI-safe snapshot of a definition-list definition body.
/// </summary>
public sealed class MarkdownNativeDefinitionListDefinitionInfo {
    internal MarkdownNativeDefinitionListDefinitionInfo(
        string markdown,
        MarkdownNativeSourceSpanInfo? sourceSpan,
        string? sourceText,
        string? originalSourceText,
        MarkdownOriginalSourceSliceFailureReason? originalSourceFailureReason,
        IReadOnlyList<MarkdownNativeBlockInfo> children) {
        Markdown = markdown ?? string.Empty;
        SourceSpan = sourceSpan;
        SourceText = sourceText;
        OriginalSourceText = originalSourceText;
        OriginalSourceFailureReason = originalSourceFailureReason;
        Children = children ?? Array.Empty<MarkdownNativeBlockInfo>();
    }

    /// <summary>Markdown definition body.</summary>
    public string Markdown { get; }

    /// <summary>Source span snapshot when available.</summary>
    public MarkdownNativeSourceSpanInfo? SourceSpan { get; }

    /// <summary>Normalized markdown text backing this definition body when available.</summary>
    public string? SourceText { get; }

    /// <summary>Original reader input backing this definition body when available.</summary>
    public string? OriginalSourceText { get; }

    /// <summary>Reason original reader input could not be materialized for this definition body, when applicable.</summary>
    public MarkdownOriginalSourceSliceFailureReason? OriginalSourceFailureReason { get; }

    /// <summary>Nested child block snapshots.</summary>
    public IReadOnlyList<MarkdownNativeBlockInfo> Children { get; }
}

/// <summary>
/// UI-safe snapshot of a native inline.
/// </summary>
public sealed class MarkdownNativeInlineInfo {
    internal MarkdownNativeInlineInfo(
        string id,
        MarkdownNativeInlineKind kind,
        MarkdownSyntaxKind syntaxKind,
        string text,
        string markdown,
        string literal,
        MarkdownNativeSourceSpanInfo? sourceSpan,
        IReadOnlyDictionary<string, string> metadata,
        IReadOnlyDictionary<string, MarkdownNativeSourceSpanInfo?> metadataSourceSpans,
        IReadOnlyList<MarkdownNativeInlineMetadataInfo> metadataFields,
        IReadOnlyList<MarkdownNativeInlineInfo> children) {
        Id = id ?? string.Empty;
        Kind = kind;
        SyntaxKind = syntaxKind;
        Text = text ?? string.Empty;
        Markdown = markdown ?? string.Empty;
        Literal = literal ?? string.Empty;
        SourceSpan = sourceSpan;
        Metadata = metadata ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        MetadataSourceSpans = metadataSourceSpans ?? new Dictionary<string, MarkdownNativeSourceSpanInfo?>(StringComparer.OrdinalIgnoreCase);
        MetadataFields = metadataFields ?? Array.Empty<MarkdownNativeInlineMetadataInfo>();
        Children = children ?? Array.Empty<MarkdownNativeInlineInfo>();
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
    public MarkdownNativeSourceSpanInfo? SourceSpan { get; }

    /// <summary>Metadata values such as target/title/source/alt.</summary>
    public IReadOnlyDictionary<string, string> Metadata { get; }

    /// <summary>Source spans for metadata values such as target/title/source/alt.</summary>
    public IReadOnlyDictionary<string, MarkdownNativeSourceSpanInfo?> MetadataSourceSpans { get; }

    /// <summary>Source-backed metadata fields in source order, including repeated metadata names.</summary>
    public IReadOnlyList<MarkdownNativeInlineMetadataInfo> MetadataFields { get; }

    /// <summary>Nested inline snapshots.</summary>
    public IReadOnlyList<MarkdownNativeInlineInfo> Children { get; }

    /// <summary>Enumerates source-backed metadata fields with the supplied field name in source order.</summary>
    public IEnumerable<MarkdownNativeInlineMetadataInfo> EnumerateMetadataFields(string name) {
        if (string.IsNullOrWhiteSpace(name) || MetadataFields.Count == 0) {
            yield break;
        }

        for (var i = 0; i < MetadataFields.Count; i++) {
            var field = MetadataFields[i];
            if (string.Equals(field.Name, name, StringComparison.OrdinalIgnoreCase)) {
                yield return field;
            }
        }
    }

    /// <summary>
    /// Finds the first source-backed metadata field with the supplied name, optionally constrained to an occurrence index.
    /// </summary>
    public MarkdownNativeInlineMetadataInfo? FindMetadataField(string name, int index = -1) {
        if (string.IsNullOrWhiteSpace(name) || MetadataFields.Count == 0) {
            return null;
        }

        for (var i = 0; i < MetadataFields.Count; i++) {
            var field = MetadataFields[i];
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
/// UI-safe snapshot of a source-backed inline metadata field.
/// </summary>
public sealed class MarkdownNativeInlineMetadataInfo {
    internal MarkdownNativeInlineMetadataInfo(MarkdownNativeInlineMetadata metadata, int index) {
        Name = metadata?.Name ?? string.Empty;
        Value = metadata?.Value ?? string.Empty;
        SourceSpan = metadata?.SourceSpan.HasValue == true
            ? new MarkdownNativeSourceSpanInfo(metadata.SourceSpan.Value)
            : null;
        Index = index;
    }

    internal MarkdownNativeInlineMetadataInfo(
        MarkdownNativeInlineMetadata metadata,
        int index,
        string? sourceText,
        string? originalSourceText,
        MarkdownOriginalSourceSliceFailureReason? originalSourceFailureReason) {
        Name = metadata?.Name ?? string.Empty;
        Value = metadata?.Value ?? string.Empty;
        SourceSpan = metadata?.SourceSpan.HasValue == true
            ? new MarkdownNativeSourceSpanInfo(metadata.SourceSpan.Value)
            : null;
        Index = index;
        SourceText = sourceText;
        OriginalSourceText = originalSourceText;
        OriginalSourceFailureReason = originalSourceFailureReason;
    }

    /// <summary>Stable metadata field name such as <c>target</c>, <c>title</c>, or <c>openingMarker</c>.</summary>
    public string Name { get; }

    /// <summary>Semantic value represented by the metadata field.</summary>
    public string Value { get; }

    /// <summary>Source span for this metadata field when available.</summary>
    public MarkdownNativeSourceSpanInfo? SourceSpan { get; }

    /// <summary>Zero-based occurrence index in the source-order metadata list.</summary>
    public int Index { get; }

    /// <summary>Exact normalized source text represented by the metadata field when it could be materialized.</summary>
    public string? SourceText { get; }

    /// <summary>Exact original reader input represented by the metadata field when trivia was preserved and mapping succeeded.</summary>
    public string? OriginalSourceText { get; }

    /// <summary>Reason original reader input could not be materialized for this metadata field, when applicable.</summary>
    public MarkdownOriginalSourceSliceFailureReason? OriginalSourceFailureReason { get; }
}

/// <summary>
/// UI-safe snapshot of a native list item.
/// </summary>
public sealed class MarkdownNativeListItemInfo {
    internal MarkdownNativeListItemInfo(
        string id,
        string text,
        bool isTask,
        bool isChecked,
        int level,
        MarkdownNativeSourceSpanInfo? sourceSpan,
        MarkdownNativeSourceSpanInfo? markerSourceSpan,
        string? markerText,
        MarkdownNativeSourceSpanInfo? taskMarkerSourceSpan,
        string? taskMarkerText,
        string? sourceText,
        string? originalSourceText,
        MarkdownOriginalSourceSliceFailureReason? originalSourceFailureReason,
        IReadOnlyList<MarkdownNativeInlineInfo> inlines,
        IReadOnlyList<MarkdownNativeListItemParagraphInfo> paragraphs,
        IReadOnlyList<MarkdownNativeBlockInfo> children) {
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
        SourceText = sourceText;
        OriginalSourceText = originalSourceText;
        OriginalSourceFailureReason = originalSourceFailureReason;
        Inlines = inlines ?? Array.Empty<MarkdownNativeInlineInfo>();
        Paragraphs = paragraphs ?? Array.Empty<MarkdownNativeListItemParagraphInfo>();
        Children = children ?? Array.Empty<MarkdownNativeBlockInfo>();
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
    public MarkdownNativeSourceSpanInfo? SourceSpan { get; }

    /// <summary>List marker token source span snapshot when available.</summary>
    public MarkdownNativeSourceSpanInfo? MarkerSourceSpan { get; }

    /// <summary>Exact list marker token when available.</summary>
    public string? MarkerText { get; }

    /// <summary>Task marker token source span snapshot when available.</summary>
    public MarkdownNativeSourceSpanInfo? TaskMarkerSourceSpan { get; }

    /// <summary>Exact task marker token when available.</summary>
    public string? TaskMarkerText { get; }

    /// <summary>Normalized markdown text backing the list item content span when available.</summary>
    public string? SourceText { get; }

    /// <summary>Original reader input backing the list item content span when available.</summary>
    public string? OriginalSourceText { get; }

    /// <summary>Reason original reader input could not be materialized for this list item, when applicable.</summary>
    public MarkdownOriginalSourceSliceFailureReason? OriginalSourceFailureReason { get; }

    /// <summary>Lead inline snapshots.</summary>
    public IReadOnlyList<MarkdownNativeInlineInfo> Inlines { get; }

    /// <summary>MarkdownParagraph-level snapshots owned by this list item.</summary>
    public IReadOnlyList<MarkdownNativeListItemParagraphInfo> Paragraphs { get; }

    /// <summary>Nested child block snapshots.</summary>
    public IReadOnlyList<MarkdownNativeBlockInfo> Children { get; }
}

/// <summary>
/// UI-safe snapshot of one paragraph owned by a native list item.
/// </summary>
public sealed class MarkdownNativeListItemParagraphInfo {
    internal MarkdownNativeListItemParagraphInfo(
        int index,
        string text,
        MarkdownNativeSourceSpanInfo? sourceSpan,
        string? sourceText,
        string? originalSourceText,
        MarkdownOriginalSourceSliceFailureReason? originalSourceFailureReason,
        IReadOnlyList<MarkdownNativeInlineInfo> inlines) {
        Index = index;
        Text = text ?? string.Empty;
        SourceSpan = sourceSpan;
        SourceText = sourceText;
        OriginalSourceText = originalSourceText;
        OriginalSourceFailureReason = originalSourceFailureReason;
        Inlines = inlines ?? Array.Empty<MarkdownNativeInlineInfo>();
    }

    /// <summary>Zero-based paragraph index within the list item.</summary>
    public int Index { get; }

    /// <summary>Plain-text paragraph content.</summary>
    public string Text { get; }

    /// <summary>Source span snapshot when available.</summary>
    public MarkdownNativeSourceSpanInfo? SourceSpan { get; }

    /// <summary>Normalized markdown text backing this list-item paragraph when available.</summary>
    public string? SourceText { get; }

    /// <summary>Original reader input backing this list-item paragraph when available.</summary>
    public string? OriginalSourceText { get; }

    /// <summary>Reason original reader input could not be materialized for this paragraph, when applicable.</summary>
    public MarkdownOriginalSourceSliceFailureReason? OriginalSourceFailureReason { get; }

    /// <summary>Inline snapshots for this paragraph.</summary>
    public IReadOnlyList<MarkdownNativeInlineInfo> Inlines { get; }
}

/// <summary>
/// UI-safe snapshot of a native table row.
/// </summary>
public sealed class MarkdownNativeTableRowInfo {
    internal MarkdownNativeTableRowInfo(
        string markdown,
        bool isHeader,
        int rowIndex,
        MarkdownNativeSourceSpanInfo? sourceSpan,
        string? sourceText,
        string? originalSourceText,
        MarkdownOriginalSourceSliceFailureReason? originalSourceFailureReason,
        IReadOnlyList<MarkdownNativeTableCellInfo> cells) {
        Markdown = markdown ?? string.Empty;
        IsHeader = isHeader;
        RowIndex = rowIndex;
        SourceSpan = sourceSpan;
        SourceText = sourceText;
        OriginalSourceText = originalSourceText;
        OriginalSourceFailureReason = originalSourceFailureReason;
        Cells = cells ?? Array.Empty<MarkdownNativeTableCellInfo>();
    }

    /// <summary>Markdown representation of the row payload.</summary>
    public string Markdown { get; }

    /// <summary>Whether this is the table header row.</summary>
    public bool IsHeader { get; }

    /// <summary>Zero-based row index, or -1 for headers.</summary>
    public int RowIndex { get; }

    /// <summary>Source span snapshot when available.</summary>
    public MarkdownNativeSourceSpanInfo? SourceSpan { get; }

    /// <summary>Normalized markdown text backing this table row when available.</summary>
    public string? SourceText { get; }

    /// <summary>Original reader input backing this table row when available.</summary>
    public string? OriginalSourceText { get; }

    /// <summary>Reason original reader input could not be materialized for this table row, when applicable.</summary>
    public MarkdownOriginalSourceSliceFailureReason? OriginalSourceFailureReason { get; }

    /// <summary>Cell snapshots in document column order.</summary>
    public IReadOnlyList<MarkdownNativeTableCellInfo> Cells { get; }
}

/// <summary>
/// UI-safe snapshot of a native table cell.
/// </summary>
public sealed class MarkdownNativeTableCellInfo {
    internal MarkdownNativeTableCellInfo(
        string text,
        string markdown,
        bool isHeader,
        int rowIndex,
        int columnIndex,
        MarkdownColumnAlignment alignment,
        MarkdownNativeSourceSpanInfo? sourceSpan,
        string? sourceText,
        string? originalSourceText,
        MarkdownOriginalSourceSliceFailureReason? originalSourceFailureReason,
        IReadOnlyList<MarkdownNativeInlineInfo> inlines,
        IReadOnlyList<MarkdownNativeBlockInfo> children) {
        Text = text ?? string.Empty;
        Markdown = markdown ?? string.Empty;
        IsHeader = isHeader;
        RowIndex = rowIndex;
        ColumnIndex = columnIndex;
        Alignment = alignment;
        SourceSpan = sourceSpan;
        SourceText = sourceText;
        OriginalSourceText = originalSourceText;
        OriginalSourceFailureReason = originalSourceFailureReason;
        Inlines = inlines ?? Array.Empty<MarkdownNativeInlineInfo>();
        Children = children ?? Array.Empty<MarkdownNativeBlockInfo>();
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
    public MarkdownColumnAlignment Alignment { get; }

    /// <summary>Source span snapshot when available.</summary>
    public MarkdownNativeSourceSpanInfo? SourceSpan { get; }

    /// <summary>Normalized markdown text backing this table cell when available.</summary>
    public string? SourceText { get; }

    /// <summary>Original reader input backing this table cell when available.</summary>
    public string? OriginalSourceText { get; }

    /// <summary>Reason original reader input could not be materialized for this table cell, when applicable.</summary>
    public MarkdownOriginalSourceSliceFailureReason? OriginalSourceFailureReason { get; }

    /// <summary>Inline snapshots for cell content when available.</summary>
    public IReadOnlyList<MarkdownNativeInlineInfo> Inlines { get; }

    /// <summary>Native child block snapshots projected from structured cell content.</summary>
    public IReadOnlyList<MarkdownNativeBlockInfo> Children { get; }
}

/// <summary>
/// UI-safe source span snapshot.
/// </summary>
public sealed class MarkdownNativeSourceSpanInfo {
    internal MarkdownNativeSourceSpanInfo(MarkdownSourceSpan span) {
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
public sealed class MarkdownNativeDiagnosticInfo {
    internal MarkdownNativeDiagnosticInfo(MarkdownNativeDiagnostic diagnostic) {
        Id = diagnostic.Id;
        Message = diagnostic.Message;
        Severity = diagnostic.Severity;
        SourceSpan = diagnostic.SourceSpan.HasValue ? new MarkdownNativeSourceSpanInfo(diagnostic.SourceSpan.Value) : null;
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
    public MarkdownNativeSourceSpanInfo? SourceSpan { get; }

    /// <summary>Additional related source spans, such as individual transform input blocks.</summary>
    public IReadOnlyList<MarkdownNativeSourceSpanInfo> RelatedSourceSpans { get; }

    /// <summary>Associated block id when available.</summary>
    public string? BlockId { get; }

    private static IReadOnlyList<MarkdownNativeSourceSpanInfo> ToSourceSpanSnapshots(IReadOnlyList<MarkdownSourceSpan> spans) {
        if (spans == null || spans.Count == 0) {
            return Array.Empty<MarkdownNativeSourceSpanInfo>();
        }

        var snapshots = new List<MarkdownNativeSourceSpanInfo>(spans.Count);
        for (var i = 0; i < spans.Count; i++) {
            snapshots.Add(new MarkdownNativeSourceSpanInfo(spans[i]));
        }

        return snapshots;
    }
}
