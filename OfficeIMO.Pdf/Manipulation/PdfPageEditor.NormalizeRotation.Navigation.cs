namespace OfficeIMO.Pdf;

internal static partial class PdfPageEditor {
    // Destination coordinates use default page user space. Quarter-turns exchange the axes;
    // null coordinates still mean retain, and fit destinations must remain fit destinations.
    private static bool TransformNormalizationDestination(PdfArray destination, string mode, PageResizeTransform transform) {
        if (mode == "XYZ" && destination.Items.Count >= 5) {
            PdfObject x = destination.Items[2], y = destination.Items[3];
            destination.Items[2] = transform.HasAxisSwap
                ? TransformNormalizationCoordinate(y, transform.C, transform.E)
                : TransformNormalizationCoordinate(x, transform.A, transform.E);
            destination.Items[3] = transform.HasAxisSwap
                ? TransformNormalizationCoordinate(x, transform.B, transform.F)
                : TransformNormalizationCoordinate(y, transform.D, transform.F);
            return true;
        }
        bool horizontal = mode == "FitH" || mode == "FitBH";
        bool vertical = mode == "FitV" || mode == "FitBV";
        if ((!horizontal && !vertical) || destination.Items.Count < 3) return false;
        if (transform.HasAxisSwap) {
            destination.Items[1] = new PdfName(mode == "FitH" ? "FitV" : mode == "FitV" ? "FitH" : mode == "FitBH" ? "FitBV" : "FitBH");
            destination.Items[2] = horizontal
                ? TransformNormalizationCoordinate(destination.Items[2], transform.C, transform.E)
                : TransformNormalizationCoordinate(destination.Items[2], transform.B, transform.F);
        } else {
            destination.Items[2] = horizontal
                ? TransformNormalizationCoordinate(destination.Items[2], transform.D, transform.F)
                : TransformNormalizationCoordinate(destination.Items[2], transform.A, transform.E);
        }
        return true;
    }

    private static PdfObject TransformNormalizationCoordinate(PdfObject coordinate, double scale, double offset) =>
        coordinate is PdfNumber number ? new PdfNumber(number.Value * scale + offset) : ClonePdfObject(coordinate);
}
