using System.Globalization;

namespace OfficeIMO.Pdf;

/// <summary>Controls culture-sensitive parsing during logical table value analysis.</summary>
public sealed class PdfLogicalTableValueAnalysisOptions {
    /// <summary>
    /// Culture used for numbers and percentages. The default is invariant culture.
    /// </summary>
    public CultureInfo NumericCulture { get; set; } = CultureInfo.InvariantCulture;

    /// <summary>
    /// Optional culture used for dates, date-times, and clock times. When null, date inference accepts only
    /// unambiguous invariant year-first forms and clock times use invariant culture.
    /// </summary>
    public CultureInfo? DateTimeCulture { get; set; }
}
