namespace OfficeIMO.Bibliography;

/// <summary>Limits and precedence for an explicit bibliography reference-resolution operation.</summary>
public sealed class BibliographyReferenceOptions {
    /// <summary>Resolves BibTeX/BibLaTeX <c>crossref</c> fields. Existing child values take precedence.</summary>
    public bool ResolveCrossref { get; set; } = true;
    /// <summary>Resolves whole-entry BibLaTeX <c>xdata</c> fields in their listed order.</summary>
    public bool ResolveXData { get; set; } = true;
    /// <summary>Allows xdata to replace existing fields, including child fields. Later containers win, as in Biber.</summary>
    public bool XDataOverridesExistingFields { get; set; } = true;
    /// <summary>Maximum number of source records. The default is 250,000.</summary>
    public int MaximumItems { get; set; } = 250_000;
    /// <summary>Maximum number of reference edges, including unsuccessful references. The default is 1,000,000.</summary>
    public int MaximumReferences { get; set; } = 1_000_000;
    /// <summary>Maximum number of reference edges in a chain. The default is 64.</summary>
    public int MaximumDepth { get; set; } = 64;
    /// <summary>Maximum copied values across the snapshot and inherited fields. The default is 2,000,000.</summary>
    public int MaximumValues { get; set; } = 2_000_000;
    /// <summary>Maximum aggregate UTF-16 characters in copied values. The default is 64 MiB.</summary>
    public long MaximumExpandedCharacters { get; set; } = 64L * 1024 * 1024;
    /// <summary>Maximum diagnostics before the operation fails. The default is 10,000.</summary>
    public int MaximumDiagnostics { get; set; } = 10_000;

    internal BibliographyReferenceOptions Snapshot() {
        var copy = (BibliographyReferenceOptions)MemberwiseClone();
        if (copy.MaximumItems <= 0) throw new ArgumentOutOfRangeException(nameof(MaximumItems));
        if (copy.MaximumReferences <= 0) throw new ArgumentOutOfRangeException(nameof(MaximumReferences));
        if (copy.MaximumDepth < 0) throw new ArgumentOutOfRangeException(nameof(MaximumDepth));
        if (copy.MaximumValues <= 0) throw new ArgumentOutOfRangeException(nameof(MaximumValues));
        if (copy.MaximumExpandedCharacters <= 0) throw new ArgumentOutOfRangeException(nameof(MaximumExpandedCharacters));
        if (copy.MaximumDiagnostics <= 0) throw new ArgumentOutOfRangeException(nameof(MaximumDiagnostics));
        return copy;
    }
}

/// <summary>The final origin of one inherited field or role group in a resolved snapshot.</summary>
public sealed class BibliographyFieldProvenance {
    internal BibliographyFieldProvenance(int itemIndex, string itemKey, string field, int sourceItemIndex,
        string sourceItemKey, string sourceField, string relation, IReadOnlyList<string> referencePath) {
        ItemIndex = itemIndex; ItemKey = itemKey; Field = field; SourceItemIndex = sourceItemIndex;
        SourceItemKey = sourceItemKey; SourceField = sourceField; Relation = relation;
        ReferencePath = referencePath;
    }
    /// <summary>Zero-based destination record index in <see cref="BibliographyReferenceResult.Document"/>.</summary>
    public int ItemIndex { get; }
    /// <summary>Destination record key.</summary>
    public string ItemKey { get; }
    /// <summary>Canonical model field, such as <c>container-title</c>, <c>contributors.author</c>, or <c>dates.issued</c>.</summary>
    public string Field { get; }
    /// <summary>Zero-based original source record index.</summary>
    public int SourceItemIndex { get; }
    /// <summary>Key of the record that originally supplied the field.</summary>
    public string SourceItemKey { get; }
    /// <summary>Canonical field in the original supplying record.</summary>
    public string SourceField { get; }
    /// <summary>Immediate inheritance mechanism: <c>crossref</c> or <c>xdata</c>.</summary>
    public string Relation { get; }
    /// <summary>Read-only keys from the destination through its parents to the original supplying record.</summary>
    public IReadOnlyList<string> ReferencePath { get; }
}

/// <summary>An independently editable resolved snapshot and immutable operation evidence.</summary>
public sealed class BibliographyReferenceResult {
    internal BibliographyReferenceResult(BibliographyDocument document, IEnumerable<BibliographyDiagnostic> diagnostics,
        IEnumerable<BibliographyFieldProvenance> provenance) {
        Document = document;
        Diagnostics = Array.AsReadOnly(diagnostics.ToArray());
        Provenance = Array.AsReadOnly(provenance.ToArray());
        CitationItems = Array.AsReadOnly(document.Items.Where(item => !BibliographyReferenceResolver.IsDataContainer(item)).ToArray());
    }
    /// <summary>Deep-copied document in source order, including retained data containers and reference fields.</summary>
    public BibliographyDocument Document { get; }
    /// <summary>Records eligible for citation, excluding <c>@xdata</c> containers. These are the snapshot's editable items.</summary>
    public IReadOnlyList<BibliographyItem> CitationItems { get; }
    /// <summary>Read-only key, reference, cycle, and depth diagnostics.</summary>
    public IReadOnlyList<BibliographyDiagnostic> Diagnostics { get; }
    /// <summary>Read-only final inherited-field origins. Subsequent snapshot edits do not update this operation evidence.</summary>
    public IReadOnlyList<BibliographyFieldProvenance> Provenance { get; }
    /// <summary>True if resolution diagnosed an error.</summary>
    public bool HasErrors => Diagnostics.Any(diagnostic => diagnostic.Severity == BibliographyDiagnosticSeverity.Error);
    /// <summary>True if no reference or key could require caller attention.</summary>
    public bool IsComplete => Diagnostics.All(diagnostic => diagnostic.Severity == BibliographyDiagnosticSeverity.Information);
}
