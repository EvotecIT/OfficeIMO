using System;
using System.Globalization;

namespace OfficeIMO.Rtf.Markdown;

internal static class RtfMarkdownBridgeMarkers {
    private const string ListContinuationBookmarkPrefix = "_OfficeIMO_MLC_";

    internal static string CreateListContinuationBookmarkName(int ordinal) =>
        ListContinuationBookmarkPrefix + Math.Max(0, ordinal).ToString(CultureInfo.InvariantCulture);

    internal static bool IsListContinuationBookmark(RtfBookmarkMarker marker) =>
        marker.Name.StartsWith(ListContinuationBookmarkPrefix, StringComparison.Ordinal);
}
