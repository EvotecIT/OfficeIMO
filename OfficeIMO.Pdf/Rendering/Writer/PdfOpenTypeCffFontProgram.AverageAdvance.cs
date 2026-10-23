using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal sealed partial class PdfOpenTypeCffFontProgram {
    /// <summary>Optional immutable source average used by Word's preferred-character policy.</summary>
    internal OfficeOpenTypeAverageAdvanceMetrics? AverageAdvanceMetrics => _fontUnitLayoutMetrics.AverageAdvanceMetrics;
}
