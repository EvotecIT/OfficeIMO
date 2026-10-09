using OfficeIMO.Pdf;

namespace OfficeIMO.OpenDocument.Odg.Pdf;

/// <summary>Controls page-preserving Draw-to-PDF conversion through the shared Drawing engine.</summary>
public sealed class OdgToPdfOptions {
    /// <summary>PDF generation settings. Page sizes and margins come from the drawing pages.</summary>
    public PdfOptions? PdfOptions { get; set; }

    /// <summary>Selects print-visible layers. Set false to use screen-visible layers.</summary>
    public bool ForPrint { get; set; } = true;

    /// <summary>Source projection losses are reported by default. Strict policies reject output before saving.</summary>
    public OdfConversionLossPolicy LossPolicy { get; set; } = OdfConversionLossPolicy.ReportOnly;

    /// <summary>Maximum number of source pages converted in one operation. Default: 1,000.</summary>
    public int MaximumPages { get; set; } = 1000;

    /// <summary>Optional deterministic date/time field evaluation. Null preserves cached display text.</summary>
    public OdfDateTimeFieldProjectionOptions? DateTimeFields { get; set; }

    /// <summary>Creates independent settings for an operation.</summary>
    public OdgToPdfOptions Clone() => new OdgToPdfOptions {
        PdfOptions = PdfOptions?.Clone(), ForPrint = ForPrint, LossPolicy = LossPolicy, MaximumPages = MaximumPages,
        DateTimeFields = DateTimeFields?.Clone()
    };

    internal OdgToPdfOptions Snapshot() {
        if (MaximumPages < 1) throw new ArgumentOutOfRangeException(nameof(MaximumPages));
        if (LossPolicy < OdfConversionLossPolicy.ReportOnly || LossPolicy > OdfConversionLossPolicy.ThrowOnAnyLoss)
            throw new ArgumentOutOfRangeException(nameof(LossPolicy));
        return Clone();
    }
}
