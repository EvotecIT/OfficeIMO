using System.Collections.Generic;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf;

/// <summary>Editable Word output and diagnostics from one semantic PDF import.</summary>
public sealed class PdfWordConversionResult {
    internal PdfWordConversionResult(WordDocument value, PdfCore.PdfConversionReport report) {
        Value = value ?? throw new ArgumentNullException(nameof(value));
        Report = new PdfCore.PdfConversionReport();
        Report.AddRange((report ?? throw new ArgumentNullException(nameof(report))).Warnings);
    }

    /// <summary>The imported Word document.</summary>
    public WordDocument Value { get; }

    /// <summary>Snapshot of accepted degradations reported by the import.</summary>
    public PdfCore.PdfConversionReport Report { get; }

    /// <summary>Warnings captured while reconstructing editable Word content.</summary>
    public IReadOnlyList<PdfCore.PdfConversionWarning> Warnings => Report.Warnings;

    /// <summary>True when import reported at least one accepted degradation.</summary>
    public bool HasWarnings => Report.HasWarnings;

    /// <summary>Returns the imported Word document.</summary>
    public WordDocument RequireValue() => Value;
}
