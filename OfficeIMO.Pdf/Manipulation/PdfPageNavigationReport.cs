using System.Globalization;

namespace OfficeIMO.Pdf;

/// <summary>
/// Sentences recorded in the operation's pipeline step when page extraction removes or changes navigation that
/// targets pages outside the output.
/// </summary>
internal static class PdfPageNavigationReport {
    internal const string OpenActionRemoved = "The document open action was removed because it goes to a page that is not in the output.";

    internal static string BookmarkRemoved(string title) =>
        "Bookmark " + Quote(title) + " was removed because its destination page is not in the output.";

    internal static string BookmarkDestinationRemoved(string title) =>
        "Bookmark " + Quote(title) + " was kept without a destination because its destination page is not in the output; its child bookmarks are.";

    internal static string NamedDestinationsRemoved(int count) =>
        Count(count, "named destination", "named destinations") + " that go to pages outside the output " +
        (count == 1 ? "was" : "were") + " removed.";

    internal static string LinksRemoved(int count) =>
        Count(count, "link annotation", "link annotations") + " that go to pages outside the output " +
        (count == 1 ? "was" : "were") + " removed.";

    internal static string ActionsRemoved(int count) =>
        Count(count, "go-to action", "go-to actions") + " that go to pages outside the output " +
        (count == 1 ? "was" : "were") + " removed.";

    private static string Count(int count, string singular, string plural) =>
        count.ToString(CultureInfo.InvariantCulture) + " " + (count == 1 ? singular : plural);

    private static string Quote(string title) => title.Length == 0 ? "(untitled)" : "\"" + title + "\"";
}
