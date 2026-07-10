using OfficeIMO.Rtf.Syntax;

namespace OfficeIMO.Rtf;

/// <summary>
/// Registry of RTF destination categories used by readers, editors, and diagnostics.
/// </summary>
public static class RtfDestinationRegistry {
    private static readonly Dictionary<string, RtfDestinationType> Destinations = new Dictionary<string, RtfDestinationType>(StringComparer.Ordinal) {
        ["rtf"] = RtfDestinationType.Header,
        ["ansi"] = RtfDestinationType.Header,
        ["mac"] = RtfDestinationType.Header,
        ["pc"] = RtfDestinationType.Header,
        ["pca"] = RtfDestinationType.Header,
        ["ansicpg"] = RtfDestinationType.Header,
        ["deff"] = RtfDestinationType.Header,
        ["deflang"] = RtfDestinationType.Header,
        ["deflangfe"] = RtfDestinationType.Header,
        ["adeflang"] = RtfDestinationType.Header,
        ["generator"] = RtfDestinationType.Metadata,
        ["info"] = RtfDestinationType.Metadata,
        ["title"] = RtfDestinationType.Metadata,
        ["subject"] = RtfDestinationType.Metadata,
        ["author"] = RtfDestinationType.Metadata,
        ["manager"] = RtfDestinationType.Metadata,
        ["company"] = RtfDestinationType.Metadata,
        ["operator"] = RtfDestinationType.Metadata,
        ["category"] = RtfDestinationType.Metadata,
        ["keywords"] = RtfDestinationType.Metadata,
        ["comment"] = RtfDestinationType.Metadata,
        ["userprops"] = RtfDestinationType.Metadata,
        ["propname"] = RtfDestinationType.Metadata,
        ["staticval"] = RtfDestinationType.Metadata,
        ["linkval"] = RtfDestinationType.Metadata,
        ["docvar"] = RtfDestinationType.Metadata,
        ["revtbl"] = RtfDestinationType.Metadata,
        ["rsidtbl"] = RtfDestinationType.Metadata,
        ["atnid"] = RtfDestinationType.Metadata,
        ["atnauthor"] = RtfDestinationType.Metadata,
        ["atntime"] = RtfDestinationType.Metadata,
        ["filetbl"] = RtfDestinationType.Metadata,
        ["file"] = RtfDestinationType.Metadata,
        ["xmlnstbl"] = RtfDestinationType.Metadata,
        ["xmlns"] = RtfDestinationType.Metadata,
        ["fonttbl"] = RtfDestinationType.FontTable,
        ["colortbl"] = RtfDestinationType.ColorTable,
        ["stylesheet"] = RtfDestinationType.StyleSheet,
        ["listtable"] = RtfDestinationType.ListTable,
        ["listoverridetable"] = RtfDestinationType.ListTable,
        ["pntext"] = RtfDestinationType.ListTable,
        ["listtext"] = RtfDestinationType.ListTable,
        ["pict"] = RtfDestinationType.Picture,
        ["object"] = RtfDestinationType.Object,
        ["objdata"] = RtfDestinationType.Object,
        ["shp"] = RtfDestinationType.Drawing,
        ["shpinst"] = RtfDestinationType.Drawing,
        ["shptxt"] = RtfDestinationType.Drawing,
        ["sp"] = RtfDestinationType.Drawing,
        ["sn"] = RtfDestinationType.Drawing,
        ["sv"] = RtfDestinationType.Drawing,
        ["field"] = RtfDestinationType.Field,
        ["fldinst"] = RtfDestinationType.Field,
        ["fldrslt"] = RtfDestinationType.Field,
        ["ffdata"] = RtfDestinationType.Field,
        ["ffname"] = RtfDestinationType.Field,
        ["ffdeftext"] = RtfDestinationType.Field,
        ["ffformat"] = RtfDestinationType.Field,
        ["ffhelptext"] = RtfDestinationType.Field,
        ["ffstattext"] = RtfDestinationType.Field,
        ["ffentrymcr"] = RtfDestinationType.Field,
        ["ffexitmcr"] = RtfDestinationType.Field,
        ["ffl"] = RtfDestinationType.Field,
        ["upr"] = RtfDestinationType.BodyText,
        ["ud"] = RtfDestinationType.BodyText,
        ["header"] = RtfDestinationType.HeaderFooter,
        ["footer"] = RtfDestinationType.HeaderFooter,
        ["headerl"] = RtfDestinationType.HeaderFooter,
        ["headerr"] = RtfDestinationType.HeaderFooter,
        ["headerf"] = RtfDestinationType.HeaderFooter,
        ["footerl"] = RtfDestinationType.HeaderFooter,
        ["footerr"] = RtfDestinationType.HeaderFooter,
        ["footerf"] = RtfDestinationType.HeaderFooter,
        ["footnote"] = RtfDestinationType.Footnote,
        ["endnote"] = RtfDestinationType.Endnote,
        ["annotation"] = RtfDestinationType.Annotation,
        ["bkmkstart"] = RtfDestinationType.Bookmark,
        ["bkmkend"] = RtfDestinationType.Bookmark,
        ["htmltag"] = RtfDestinationType.Metadata,
        ["mhtmltag"] = RtfDestinationType.Metadata
    };

    private static readonly HashSet<string> SemanticSkipDestinations = new HashSet<string>(StringComparer.Ordinal) {
        "fonttbl", "colortbl", "stylesheet", "generator", "info", "userprops", "docvar", "revtbl", "rsidtbl", "atnid", "atnauthor", "atntime", "filetbl", "file", "xmlnstbl", "xmlns", "listtable", "listoverridetable", "pntext", "listtext", "shpinst", "sp", "sn", "sv", "ffdata", "ffname", "ffdeftext", "ffformat", "ffhelptext", "ffstattext", "ffentrymcr", "ffexitmcr", "ffl", "htmltag", "mhtmltag"
    };

    private static readonly HashSet<string> TextReplacementSkipDestinations = new HashSet<string>(StringComparer.Ordinal) {
        "fonttbl", "colortbl", "stylesheet", "generator", "info", "userprops", "docvar", "revtbl", "rsidtbl", "atnid", "atnauthor", "atntime", "filetbl", "file", "xmlnstbl", "xmlns", "pict", "object", "objdata", "listtable", "listoverridetable", "pntext", "listtext", "fldinst", "shpinst", "sp", "sn", "sv", "ffdata", "ffname", "ffdeftext", "ffformat", "ffhelptext", "ffstattext", "ffentrymcr", "ffexitmcr", "ffl"
    };

    private static readonly HashSet<string> HeaderControlsBeforeInfo = new HashSet<string>(StringComparer.Ordinal) {
        "rtf", "ansi", "mac", "pc", "pca", "ansicpg", "deff", "deflang", "deflangfe", "adeflang"
    };

    /// <summary>
    /// Gets the registered destination type.
    /// </summary>
    public static RtfDestinationType GetDestinationType(string? destination) {
        if (destination == null) return RtfDestinationType.Unknown;
        return Destinations.TryGetValue(destination, out RtfDestinationType type) ? type : RtfDestinationType.Unknown;
    }

    /// <summary>
    /// Returns whether OfficeIMO.Rtf has a category entry for this destination.
    /// </summary>
    public static bool IsKnown(string? destination) => GetDestinationType(destination) != RtfDestinationType.Unknown;

    /// <summary>
    /// Returns whether semantic binding should skip the destination while preserving syntax.
    /// </summary>
    public static bool ShouldSkipSemanticBinding(string? destination) =>
        destination != null && SemanticSkipDestinations.Contains(destination);

    /// <summary>
    /// Returns whether visible text replacement should skip the destination.
    /// </summary>
    public static bool ShouldSkipTextReplacement(string? destination) =>
        destination != null && TextReplacementSkipDestinations.Contains(destination);

    /// <summary>
    /// Returns whether the destination is currently preserved but not semantically modeled.
    /// </summary>
    public static bool IsUnsupportedSemanticDestination(string? destination) =>
        destination != null && destination != "object" && GetDestinationType(destination) is RtfDestinationType.Object;

    /// <summary>
    /// Returns whether the group is marked with the ignorable destination control symbol.
    /// </summary>
    public static bool IsIgnorableDestinationGroup(RtfGroup group) =>
        group != null && group.Children.OfType<RtfControlSymbol>().Any(symbol => symbol.Symbol == '*');

    internal static bool IsHeaderControlBeforeInfo(string controlWord) => HeaderControlsBeforeInfo.Contains(controlWord);
}
