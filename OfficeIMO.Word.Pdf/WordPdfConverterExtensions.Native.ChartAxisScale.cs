using static OfficeIMO.Word.OpenXml.WordXml;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        // The shared reader preserves authored bounds and tick units. The legacy
        // simple-chart projection retains formats but does not carry those scales.
        private static bool HasNativeExplicitValueAxisScale(OfficeXmlNode plotArea) =>
            plotArea.Elements(C, "valAx").Any(axis =>
                axis.Child(C, "majorUnit").Exists || axis.Child(C, "minorUnit").Exists ||
                axis.Child(C, "scaling") is { IsNone: false } scaling && (scaling.Child(C, "min").Exists || scaling.Child(C, "max").Exists));
    }
}
