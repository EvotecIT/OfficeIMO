using System.Globalization;

namespace OfficeIMO.Word.Html {
    internal static class HtmlWordRoundTripMarkers {
        private const string MarkerNamespace = "_OfficeIMO7D728D6B_";
        private const string FigureBeforePrefix = MarkerNamespace + "FB_";
        private const string FigureAfterPrefix = MarkerNamespace + "FA_";

        internal const string ListItemTableTag = "OfficeIMO:7D728D6B:Html:ListItemTable:v1";

        internal static string CreateFigureCaptionBookmark(bool before, int sequence) =>
            (before ? FigureBeforePrefix : FigureAfterPrefix) +
            sequence.ToString("D8", CultureInfo.InvariantCulture);

        internal static bool IsFigureCaptionBookmark(string? name, bool before) {
            string prefix = before ? FigureBeforePrefix : FigureAfterPrefix;
            if (name == null || name.Length != prefix.Length + 8 ||
                !name.StartsWith(prefix, StringComparison.Ordinal)) {
                return false;
            }

            for (int index = prefix.Length; index < name.Length; index++) {
                if (name[index] < '0' || name[index] > '9') return false;
            }
            return true;
        }
    }
}
