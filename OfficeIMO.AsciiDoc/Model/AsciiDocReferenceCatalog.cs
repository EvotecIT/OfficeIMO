using System.Collections.ObjectModel;

namespace OfficeIMO.AsciiDoc;

/// <summary>A source-ordered snapshot of explicit anchors, generated section IDs, and footnotes, including compound blocks.</summary>
public sealed class AsciiDocReferenceCatalog {
    private readonly Dictionary<string, AsciiDocReferenceTarget> _targets = new Dictionary<string, AsciiDocReferenceTarget>(StringComparer.Ordinal);
    private readonly Dictionary<AsciiDocBlock, string> _blockIds = new Dictionary<AsciiDocBlock, string>();
    private readonly Dictionary<string, int> _nextSectionNumbers = new Dictionary<string, int>(StringComparer.Ordinal);
    private readonly Dictionary<string, AsciiDocFootnoteInline> _definitions = new Dictionary<string, AsciiDocFootnoteInline>(StringComparer.Ordinal);
    private readonly Dictionary<AsciiDocFootnoteInline, string> _footnotes = new Dictionary<AsciiDocFootnoteInline, string>();
    private readonly Dictionary<AsciiDocFootnoteInline, (AsciiDocBlock Owner, AsciiDocDocumentAttributes Attributes)> _contexts = new Dictionary<AsciiDocFootnoteInline, (AsciiDocBlock, AsciiDocDocumentAttributes)>();
    private readonly List<AsciiDocReferenceDiagnostic> _diagnostics = new List<AsciiDocReferenceDiagnostic>();
    private readonly IReadOnlyDictionary<string, AsciiDocReferenceTarget> _targetView;
    private readonly IReadOnlyDictionary<string, AsciiDocFootnoteInline> _definitionView;
    private readonly IReadOnlyList<AsciiDocReferenceDiagnostic> _diagnosticView;
    private int _anonymous;

    private AsciiDocReferenceCatalog() {
        _targetView = new ReadOnlyDictionary<string, AsciiDocReferenceTarget>(_targets);
        _definitionView = new ReadOnlyDictionary<string, AsciiDocFootnoteInline>(_definitions);
        _diagnosticView = _diagnostics.AsReadOnly();
    }
    /// <summary>Explicit and generated anchor targets, keyed with ordinal comparison.</summary>
    public IReadOnlyDictionary<string, AsciiDocReferenceTarget> Targets => _targetView;
    /// <summary>Footnote definitions in source order. Repeated references share one definition.</summary>
    public IReadOnlyDictionary<string, AsciiDocFootnoteInline> Footnotes => _definitionView;
    /// <summary>Duplicate targets, conflicting definitions, and dangling references.</summary>
    public IReadOnlyList<AsciiDocReferenceDiagnostic> Diagnostics => _diagnosticView;
    /// <summary>Finds the stable catalog label for a source footnote or returns null.</summary>
    public string? GetFootnoteLabel(AsciiDocFootnoteInline footnote) => _footnotes.TryGetValue(footnote, out string? label) ? label : null;
    /// <summary>Gets the catalog's explicit or generated ID for a block, or null when it has none.</summary>
    /// <remarks>Rebuild the catalog after editing titles, IDs, attributes, or document order.</remarks>
    public string? GetBlockId(AsciiDocBlock block) {
        if (block == null) throw new ArgumentNullException(nameof(block));
        return _blockIds.TryGetValue(block, out string? id) ? id : null;
    }
    internal (AsciiDocBlock Owner, AsciiDocDocumentAttributes Attributes) GetFootnoteContext(AsciiDocFootnoteInline footnote) => _contexts[footnote];

    /// <summary>Builds a bounded snapshot without preprocessing includes or fetching resources.</summary>
    public static AsciiDocReferenceCatalog Create(AsciiDocDocument document, int maximumNestingDepth = 64, System.Threading.CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (maximumNestingDepth < 1) throw new ArgumentOutOfRangeException(nameof(maximumNestingDepth));
        cancellationToken.ThrowIfCancellationRequested();
        var result = new AsciiDocReferenceCatalog();
        var references = new List<(string Target, AsciiDocSourceSpan Span)>();
        result.Visit(document.Blocks, references, AsciiDocDocumentAttributes.Create(), true, 0, maximumNestingDepth, cancellationToken);
        result.CheckReferences(references, cancellationToken);
        return result;
    }

    private void CheckReferences(IEnumerable<(string Target, AsciiDocSourceSpan Span)> references, System.Threading.CancellationToken token) {
        foreach (var reference in references) {
            token.ThrowIfCancellationRequested();
            if (reference.Target.IndexOfAny(new[] { '#', '/', '.' }) < 0 && !_targets.ContainsKey(reference.Target))
                Report("ADOCREF003", "Unknown cross-reference target '" + reference.Target + "'.", reference.Span);
        }
        foreach (var footnote in _footnotes) {
            token.ThrowIfCancellationRequested();
            if (!_definitions.ContainsKey(footnote.Value)) Report("ADOCREF004", "Unknown footnote '" + footnote.Key.Target + "'.", footnote.Key.Span);
        }
    }

    internal static AsciiDocReferenceCatalog CreateForBlock(AsciiDocBlock block, AsciiDocDocumentAttributes attributes, int maximumDepth) {
        var result = new AsciiDocReferenceCatalog();
        var references = new List<(string Target, AsciiDocSourceSpan Span)>();
        result.Visit(new[] { block }, references, attributes, true, 0, maximumDepth, default);
        result.CheckReferences(references, default);
        return result;
    }
    private (AsciiDocDocumentAttributes Attributes, bool SectionIds) Visit(IEnumerable<AsciiDocBlock> blocks, IList<(string Target, AsciiDocSourceSpan Span)> references, AsciiDocDocumentAttributes attributes, bool sectionIds, int depth, int maximumDepth, System.Threading.CancellationToken token) {
        if (depth >= maximumDepth) throw new InvalidDataException("AsciiDoc reference traversal exceeds MaximumNestingDepth.");
        foreach (AsciiDocBlock block in blocks) {
            token.ThrowIfCancellationRequested();
            if (block is AsciiDocAttributeEntry entry) {
                attributes = attributes.Apply(entry, true);
                if (string.Equals(entry.Name, "sectids", StringComparison.OrdinalIgnoreCase)) sectionIds = !entry.IsUnset;
            }
            if (block is AsciiDocBlockAnchor anchor) Add(anchor.Id, anchor.ReferenceText ?? anchor.Target?.BlockTitle?.Title ?? (anchor.Target as AsciiDocHeading)?.Title ?? anchor.Id, anchor.Target, false, anchor.Span, bindBlock: true);
            foreach (AsciiDocBlockAttributeList elementAttributes in block.AttributeLists)
                if (elementAttributes.Attributes.Id is string id) Add(id, block.BlockTitle?.Title ?? (block as AsciiDocHeading)?.Title ?? id, block, false, elementAttributes.Span, bindBlock: true);
            if (block is AsciiDocParagraph paragraph) VisitInlines(paragraph.Inlines, block, attributes, references, token);
            else if (block is AsciiDocHeading heading) {
                VisitInlines(heading.Inlines, block, attributes, references, token, sectionAnchor: true);
                if (sectionIds && !heading.IsDocumentTitle && !_blockIds.ContainsKey(heading)) AddGeneratedHeading(heading, attributes, token);
            }
            else if (block is AsciiDocAdmonitionBlock admonition) VisitInlines(admonition.Inlines, block, attributes, references, token);
            else if (block is AsciiDocListBlock list) foreach (AsciiDocListItem item in list.Items) VisitInlines(item.Inlines, block, attributes, references, token);
            else if (block is AsciiDocDescriptionListBlock descriptions) foreach (AsciiDocDescriptionListItem item in descriptions.Items) {
                VisitInlines(item.TermInlines, block, attributes, references, token); VisitInlines(item.DescriptionInlines, block, attributes, references, token);
            }
            if (block is AsciiDocDelimitedBlock compound && compound.GetBody(token) is AsciiDocDocument body)
                (attributes, sectionIds) = Visit(body.Blocks, references, attributes, sectionIds, depth + 1, maximumDepth, token);
            if (block is AsciiDocTableBlock table) foreach (AsciiDocTableCell cell in table.Table.Cells) {
                if (cell.GetBody(token) is AsciiDocDocument cellBody) Visit(cellBody.Blocks, references, attributes, sectionIds, depth + 1, maximumDepth, token);
                else if (cell.GetInlines(token) is AsciiDocInlineSequence cellInlines) VisitInlines(cellInlines, block, attributes, references, token);
            }
        }
        return (attributes, sectionIds);
    }

    private void VisitInlines(AsciiDocInlineSequence sequence, AsciiDocBlock owner, AsciiDocDocumentAttributes attributes, IList<(string Target, AsciiDocSourceSpan Span)> references, System.Threading.CancellationToken token, bool sectionAnchor = false) {
        foreach (AsciiDocInline inline in sequence.Items) {
            token.ThrowIfCancellationRequested();
            if (inline is AsciiDocAnchorInline anchor) Add(anchor.Id, anchor.ReferenceText ?? (sectionAnchor ?
                AsciiDocSectionIdentifier.Title(((AsciiDocHeading)owner).Inlines, attributes, token, out _) : anchor.Id), owner, anchor.IsBibliography, anchor.Span,
                bindBlock: sectionAnchor && ReferenceEquals(sequence.Items.LastOrDefault(), anchor));
            else if (inline is AsciiDocCrossReferenceInline reference) references.Add((AsciiDocAttributeSubstitutor.Substitute(reference.Target, attributes).Value, reference.Span));
            else if (inline is AsciiDocMacroInline macro && macro.Name == "xref") references.Add((AsciiDocAttributeSubstitutor.Substitute(macro.Target, attributes).Value, macro.Span));
            else if (inline is AsciiDocFormattedInline formatted) VisitInlines(formatted.Content, owner, attributes, references, token);
            else if (inline is AsciiDocFootnoteInline footnote) {
                // Named labels are encoded so arbitrary native identifiers cannot
                // activate Markdown label syntax or collide with anonymous labels.
                string label = footnote.Target.Length == 0 ? "adoc-anon-" + (++_anonymous).ToString(System.Globalization.CultureInfo.InvariantCulture) :
                    "adoc-named-" + BitConverter.ToString(Encoding.UTF8.GetBytes(footnote.Target)).Replace("-", string.Empty).ToLowerInvariant();
                if (_footnotes.ContainsKey(footnote)) continue;
                _footnotes.Add(footnote, label);
                _contexts.Add(footnote, (owner, attributes));
                if (footnote.AttributeList.Length == 0) continue;
                if (_definitions.TryGetValue(label, out AsciiDocFootnoteInline? previous)) {
                    if (previous.AttributeList != footnote.AttributeList) Report("ADOCREF002", "Conflicting footnote definition '" + footnote.Target + "'.", footnote.Span);
                } else _definitions.Add(label, footnote);
            }
        }
    }
    private void Add(string id, string label, AsciiDocBlock? block, bool bibliography, AsciiDocSourceSpan span, bool bindBlock = false, bool generated = false) {
        if (_targets.ContainsKey(id)) Report("ADOCREF001", "Duplicate anchor '" + id + "'.", span);
        else _targets.Add(id, new AsciiDocReferenceTarget(id, label, block, bibliography, span, generated));
        if (bindBlock && block != null) _blockIds[block] = id;
    }

    private void AddGeneratedHeading(AsciiDocHeading heading, AsciiDocDocumentAttributes attributes, System.Threading.CancellationToken token) {
        string title = AsciiDocSectionIdentifier.Title(heading.Inlines, attributes, token, out bool approximate);
        approximate |= heading.AttributeLists.Any(list => list.Attributes.Entries.Any(entry =>
            entry.Kind == AsciiDocElementAttributeKind.Named && string.Equals(entry.Name, "subs", StringComparison.OrdinalIgnoreCase)));
        string separator = attributes.GetValueOrDefault("idseparator") ?? "_";
        if (separator.Length > 1) separator = separator.Substring(0, char.IsSurrogatePair(separator, 0) ? 2 : 1);
        string stem = AsciiDocSectionIdentifier.Create(title, attributes.GetValueOrDefault("idprefix") ?? "_", separator, token);
        if (stem.Length == 0) {
            Report("ADOCREF006", "The section title produces an empty ID. Supply an explicit ID to create a reference target.", heading.Span);
            return;
        }
        string id = stem;
        if (_targets.ContainsKey(id)) {
            int next = _nextSectionNumbers.TryGetValue(stem, out int number) ? number : 2;
            do { token.ThrowIfCancellationRequested(); id = stem + separator + (next++).ToString(System.Globalization.CultureInfo.InvariantCulture); } while (_targets.ContainsKey(id));
            _nextSectionNumbers[stem] = next;
        }
        Add(id, AsciiDocSectionIdentifier.Label(title, token), heading, false, heading.Span, bindBlock: true, generated: true);
        if (approximate) Report("ADOCREF005", "The generated section ID uses a simplified title macro or substitution. Supply an explicit ID when interoperability requires the original processor's substitution behavior.", heading.Span);
    }
    private void Report(string code, string message, AsciiDocSourceSpan span) => _diagnostics.Add(new AsciiDocReferenceDiagnostic(code, message, span));
}

/// <summary>An explicit block, inline, bibliography anchor, or generated section target.</summary>
public sealed class AsciiDocReferenceTarget {
    internal AsciiDocReferenceTarget(string id, string label, AsciiDocBlock? block, bool bibliography, AsciiDocSourceSpan span, bool generated = false) { Id = id; Label = label; Block = block; IsBibliography = bibliography; Span = span; IsGenerated = generated; }
    /// <summary>Native anchor identifier.</summary>
    public string Id { get; }
    /// <summary>Reference label from source metadata, or the identifier.</summary>
    public string Label { get; }
    /// <summary>Target block, when bound.</summary>
    public AsciiDocBlock? Block { get; }
    /// <summary>Whether the anchor identifies a bibliography entry.</summary>
    public bool IsBibliography { get; }
    /// <summary>Whether this target's ID was generated from its section title.</summary>
    public bool IsGenerated { get; }
    /// <summary>Original anchor source span.</summary>
    public AsciiDocSourceSpan Span { get; }
}

/// <summary>A reference-integrity diagnostic tied to source syntax.</summary>
public sealed class AsciiDocReferenceDiagnostic {
    internal AsciiDocReferenceDiagnostic(string code, string message, AsciiDocSourceSpan span) { Code = code; Message = message; Span = span; }
    /// <summary>Stable diagnostic code.</summary>
    public string Code { get; }
    /// <summary>Description of the ambiguous or missing reference.</summary>
    public string Message { get; }
    /// <summary>Source span of the reference.</summary>
    public AsciiDocSourceSpan Span { get; }
}
