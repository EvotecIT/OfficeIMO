namespace OfficeIMO.Pdf;

/// <summary>Optional accessibility attributes for a canvas structure container.</summary>
public sealed class PdfCanvasStructureOptions {
    private string? _alternativeText;
    private PdfCanvasTableHeaderScope? _headerScope;
    private int _columnSpan = 1;
    private int _rowSpan = 1;
    private List<PdfEmbeddedFile>? _associatedFiles;

    internal string? StructureElementKey { get; set; }

    /// <summary>Alternative text associated with the structure container.</summary>
    public string? AlternativeText {
        get => _alternativeText;
        set {
            if (value != null) Guard.NotNullOrWhiteSpace(value, nameof(AlternativeText));
            _alternativeText = value?.Trim();
        }
    }

    /// <summary>Row, column, or combined scope for a table-header cell.</summary>
    public PdfCanvasTableHeaderScope? HeaderScope {
        get => _headerScope;
        set {
            if (value.HasValue && ((int)value.Value < (int)PdfCanvasTableHeaderScope.Row || (int)value.Value > (int)PdfCanvasTableHeaderScope.Both)) {
                throw new ArgumentOutOfRangeException(nameof(HeaderScope));
            }
            _headerScope = value;
        }
    }

    /// <summary>Number of table columns occupied by a tagged cell.</summary>
    public int ColumnSpan {
        get => _columnSpan;
        set => _columnSpan = ValidateSpan(value, nameof(ColumnSpan));
    }

    /// <summary>Number of table rows occupied by a tagged cell.</summary>
    public int RowSpan {
        get => _rowSpan;
        set => _rowSpan = ValidateSpan(value, nameof(RowSpan));
    }

    /// <summary>Files associated with this structure element. Returned descriptions and payloads are independent copies.</summary>
    public IReadOnlyList<PdfEmbeddedFile> AssociatedFiles => _associatedFiles == null
        ? Array.Empty<PdfEmbeddedFile>()
        : _associatedFiles.Select(file => file.Clone()).ToList().AsReadOnly();

    /// <summary>Associates a snapshot of an embedded file with this structure element.</summary>
    /// <remarks>Requires <see cref="PdfTaggedStructureMode.CatalogMarkers"/>. A MIME type is required. File names must be unique within the generated document,
    /// except that identical structure attachments may share one file specification. Plain PDF output
    /// uses PDF 2.0 for these associations; PDF/A-3 groundwork uses PDF 1.7. This does not request or prove conformance.</remarks>
    public PdfCanvasStructureOptions AddAssociatedFile(PdfEmbeddedFile file) {
        Guard.NotNull(file, nameof(file));
        if (string.IsNullOrWhiteSpace(file.MimeType))
            throw new ArgumentException("A structure-associated file requires a MIME type.", nameof(file));
        if (_associatedFiles?.Any(existing => string.Equals(existing.FileName, file.FileName, StringComparison.Ordinal)) == true)
            throw new ArgumentException("Structure-associated file names must be unique.", nameof(file));
        (_associatedFiles ??= new List<PdfEmbeddedFile>()).Add(file.Clone());
        return this;
    }

    // Owned snapshots are never exposed or mutated. Clones can share payload descriptions without
    // making an additional byte-array copy for every page fragment of the same structure element.
    internal IReadOnlyList<PdfEmbeddedFile> AssociatedFileSnapshots => _associatedFiles ?? (IReadOnlyList<PdfEmbeddedFile>)Array.Empty<PdfEmbeddedFile>();

    internal PdfCanvasStructureOptions Clone() => new PdfCanvasStructureOptions {
        AlternativeText = AlternativeText,
        HeaderScope = HeaderScope,
        ColumnSpan = ColumnSpan,
        RowSpan = RowSpan,
        StructureElementKey = StructureElementKey,
        _associatedFiles = _associatedFiles == null ? null : new List<PdfEmbeddedFile>(_associatedFiles)
    };

    private static int ValidateSpan(int value, string parameterName) {
        if (value < 1 || value > 1000) throw new ArgumentOutOfRangeException(parameterName, "Canvas table spans must be between 1 and 1000.");
        return value;
    }
}
