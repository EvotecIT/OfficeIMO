using System.Linq;
using DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        // The shared reader preserves authored bounds and tick units. The legacy
        // simple-chart projection retains formats but does not carry those scales.
        private static bool HasNativeExplicitValueAxisScale(PlotArea plotArea) =>
            plotArea.Elements<ValueAxis>().Any(axis =>
                axis.GetFirstChild<MajorUnit>() != null || axis.GetFirstChild<MinorUnit>() != null ||
                axis.Scaling?.MinAxisValue != null || axis.Scaling?.MaxAxisValue != null);
    }
}
