using OfficeIMO.Drawing;

namespace OfficeIMO.OpenDocument;

public sealed partial class OdgPage {
    private static OfficeShape? ProjectEnhancedGeometry(OdgShape shape, double width, double height, OdfConversionReport report, string feature) {
        try { return shape.CreateDrawingGeometry(width, height); }
        catch (Exception exception) when (exception is FormatException || exception is ArgumentException || exception is NotSupportedException || exception is InvalidDataException || exception is OverflowException) {
            report.Add(feature, OdfConversionMappingStatus.Skipped, message: "Native custom-shape geometry is preserved but not projected: " + exception.Message);
            return null;
        }
    }

    private static void ReportEnhancedTextArea(OdgShape shape, OdfConversionReport report) {
        if (shape.ElementName != "custom-shape") return;
        XElement? geometry = shape.Element.Element(OdfNamespaces.Draw + "enhanced-geometry");
        if (geometry?.Attributes().Any(attribute => attribute.Name == OdfNamespaces.Draw + "text-areas" ||
            attribute.Name == OdfNamespaces.Draw + "text-rotate-angle" || attribute.Name == OdfNamespaces.Draw + "text-path" && attribute.Value is not ("false" or "0")) == true)
            report.Add("shape:" + shape.Name + ":enhanced-text-area", OdfConversionMappingStatus.Unsupported,
                message: "Enhanced text areas, rotation and text-on-path declarations are preserved; supported text uses the ordinary shape text frame.");
    }
}
