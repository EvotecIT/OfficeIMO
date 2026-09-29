using System.Linq;
using OfficeIMO.Drawing;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private static bool TryCreateSharedNativeWordChartSnapshot(WordChart chart, out OfficeChartSnapshot? snapshot, out string? warning) {
            snapshot = null;
            warning = null;
            if (!chart.TryGetOfficeSnapshot(out var shared)) {
                warning = "Word charts are not partially exported when complete cached data and supported appearance cannot be projected through the shared chart reader.";
                return false;
            }
            if (shared.Data.Series.Count > MaxNativeWordChartSeries || shared.Data.Categories.Count > MaxNativeWordChartPoints ||
                shared.Data.Series.Any(series => series.Values.Count > MaxNativeWordChartPoints)) {
                warning = "Word chart cache exceeds the maximum supported PDF series or point count.";
                return false;
            }
            var size = GetNativeWordChartSizePoints(chart);
            snapshot = shared.WithSize(size.Width, size.Height);
            return true;
        }
    }
}
