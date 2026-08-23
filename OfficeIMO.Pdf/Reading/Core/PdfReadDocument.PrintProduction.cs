namespace OfficeIMO.Pdf;

public sealed partial class PdfReadDocument {
    /// <summary>
    /// Inspects readable page, form, pattern, image, shading, graphics-state, and transparency-group
    /// objects for device-color and transparency evidence used by print-production workflows.
    /// </summary>
    /// <returns>Exact-artifact color and transparency evidence for the loaded PDF.</returns>
    public PdfPrintProductionColorEvidence InspectPrintProductionColors() {
        DemandContentExtraction("print-production content");
        return PdfPrintProductionColorInspector.Inspect(this);
    }
}
