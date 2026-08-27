namespace OfficeIMO.Pdf;

/// <summary>
/// Basic PDF document metadata extracted from the Info dictionary.
/// </summary>
public sealed class PdfMetadata {
    private DateTimeOffset? _creationDate;
    private DateTimeOffset? _modificationDate;
    /// <summary>Document title.</summary>
    public string? Title { get; set; }
    /// <summary>Document author.</summary>
    public string? Author { get; set; }
    /// <summary>Document subject.</summary>
    public string? Subject { get; set; }
    /// <summary>Document keywords.</summary>
    public string? Keywords { get; set; }
    /// <summary>Print trapping status from the Info dictionary.</summary>
    public PdfTrappingStatus? TrappingStatus { get; set; }
    /// <summary>Document creation date from the Info dictionary.</summary>
    public DateTimeOffset? CreationDate {
        get => _creationDate;
        set {
            _creationDate = value;
            CreationDateRaw = null;
            CreationDateIsProductionPrecise = false;
        }
    }
    /// <summary>Document modification date from the Info dictionary.</summary>
    public DateTimeOffset? ModificationDate {
        get => _modificationDate;
        set {
            _modificationDate = value;
            ModificationDateRaw = null;
            ModificationDateIsProductionPrecise = false;
        }
    }
    internal string? CreationDateRaw { get; private set; }
    internal string? ModificationDateRaw { get; private set; }
    internal bool CreationDateIsProductionPrecise { get; set; }
    internal bool ModificationDateIsProductionPrecise { get; set; }
    /// <summary>PDF/X version from <c>GTS_PDFXVersion</c> in the Info dictionary.</summary>
    public string? PdfXVersion { get; set; }
    /// <summary>PDF/X conformance from <c>GTS_PDFXConformance</c> in the Info dictionary.</summary>
    public string? PdfXConformance { get; set; }

    internal void SetCreationDateFromSource(DateTimeOffset? value, string? raw, bool productionPrecise) {
        _creationDate = value;
        CreationDateRaw = value.HasValue ? raw : null;
        CreationDateIsProductionPrecise = productionPrecise;
    }

    internal void SetModificationDateFromSource(DateTimeOffset? value, string? raw, bool productionPrecise) {
        _modificationDate = value;
        ModificationDateRaw = value.HasValue ? raw : null;
        ModificationDateIsProductionPrecise = productionPrecise;
    }

    internal void CopySourceDatesFrom(PdfMetadata source) {
        Guard.NotNull(source, nameof(source));
        SetCreationDateFromSource(source.CreationDate, source.CreationDateRaw, source.CreationDateIsProductionPrecise);
        SetModificationDateFromSource(source.ModificationDate, source.ModificationDateRaw, source.ModificationDateIsProductionPrecise);
    }
}
