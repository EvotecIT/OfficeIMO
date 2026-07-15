namespace OfficeIMO.Pdf;

/// <summary>Configures the collection dictionary used to present generated embedded files as a portfolio.</summary>
public sealed class PdfPortfolioOptions {
    private readonly List<PdfPortfolioField> _fields = new();
    private string? _initialDocumentFileName;
    private PdfPortfolioFieldKind? _sortBy;

    /// <summary>Creates portfolio options with a visible file-name field.</summary>
    public PdfPortfolioOptions() {
        _fields.Add(new PdfPortfolioField(PdfPortfolioFieldKind.FileName, "File name"));
    }

    /// <summary>Initial viewer presentation.</summary>
    public PdfPortfolioView View { get; set; } = PdfPortfolioView.Details;

    /// <summary>Optional embedded file to open initially. The name must match a generated embedded file.</summary>
    public string? InitialDocumentFileName {
        get => _initialDocumentFileName;
        set {
            if (value != null) Guard.NotNullOrWhiteSpace(value, nameof(InitialDocumentFileName));
            _initialDocumentFileName = value;
        }
    }

    /// <summary>Optional standard property used to sort portfolio entries.</summary>
    public PdfPortfolioFieldKind? SortBy {
        get => _sortBy;
        set {
            if (value.HasValue) PdfPortfolioField.ValidateKind(value.Value, nameof(SortBy));
            _sortBy = value;
        }
    }

    /// <summary>Whether the configured sort is ascending.</summary>
    public bool SortAscending { get; set; } = true;

    /// <summary>Configured portfolio fields in display order.</summary>
    public IReadOnlyList<PdfPortfolioField> Fields => _fields
        .OrderBy(item => item.Order)
        .Select(item => item.Clone())
        .ToList()
        .AsReadOnly();

    /// <summary>Adds or replaces a standard portfolio field.</summary>
    public PdfPortfolioOptions SetField(PdfPortfolioField field) {
        Guard.NotNull(field, nameof(field));
        _fields.RemoveAll(existing => existing.Kind == field.Kind);
        _fields.Add(field.Clone());
        return this;
    }

    /// <summary>Removes a standard portfolio field.</summary>
    public PdfPortfolioOptions RemoveField(PdfPortfolioFieldKind kind) {
        PdfPortfolioField.ValidateKind(kind, nameof(kind));
        _fields.RemoveAll(field => field.Kind == kind);
        return this;
    }

    /// <summary>Removes every portfolio field.</summary>
    public PdfPortfolioOptions ClearFields() {
        _fields.Clear();
        return this;
    }

    internal PdfPortfolioOptions Clone() {
        var clone = new PdfPortfolioOptions {
            View = View,
            InitialDocumentFileName = InitialDocumentFileName,
            SortBy = SortBy,
            SortAscending = SortAscending
        };
        clone._fields.Clear();
        clone._fields.AddRange(_fields.Select(field => field.Clone()));
        return clone;
    }
}
