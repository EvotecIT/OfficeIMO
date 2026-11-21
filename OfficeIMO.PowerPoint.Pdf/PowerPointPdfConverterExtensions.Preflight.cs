using PdfCore = OfficeIMO.Pdf;
using PptCore = OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.Pdf;

public static partial class PowerPointPdfConverterExtensions {
    /// <summary>
    /// Returns independent preflight options that measure text with the fonts PDF export uses with <paramref name="pdfOptions"/>.
    /// </summary>
    /// <remarks>
    /// Text overflow and auto-fit checks then lay out each text box with the same font registrations, substitutions,
    /// <see cref="PowerPointPdfExportOptions.FontFamily"/>, text fallbacks and <see cref="PowerPointPdfExportOptions.ResourcePolicy"/>
    /// as PDF export, so a finding means the text clips in that PDF. The options are copied when this method is called.
    /// Without this selection, preflight measures with installed system fonts.
    /// </remarks>
    /// <param name="options">Preflight settings to copy; the caller's instance is unchanged.</param>
    /// <param name="pdfOptions">PDF export options whose text measurement preflight uses.</param>
    /// <returns>Preflight options that measure text as PDF export does.</returns>
    public static PptCore.PowerPointDeckPreflightOptions WithPdfTextMeasurement(this PptCore.PowerPointDeckPreflightOptions options,
        PowerPointPdfExportOptions pdfOptions) {
        ArgumentNullException.ThrowIfNull(options);
        ArgumentNullException.ThrowIfNull(pdfOptions);
        PowerPointPdfExportOptions export = pdfOptions.Clone();
        PptCore.PowerPointDeckPreflightOptions selected = options.Clone();
        selected.TextMeasurementProvider = presentation => {
            PowerPointPdfExportOptions operation = export.CloneForConversion();
            PdfCore.PdfCreateOptions created = CreatePdfOptions(presentation, operation);
            return new PptCore.PowerPointTextMeasurement(PdfCore.PdfWriter.CreateDrawingTextMeasure(created),
                ResolveNativeTextDefaultFontFamily(operation, created));
        };
        return selected;
    }
}
