namespace OfficeIMO.Rtf;

public sealed partial class RtfDocument {
    private void ImportMergeSettings(RtfDocument source, RtfConversionReport report, bool hasDestinationContent) {
        RtfDocumentSettings from = source.Settings;
        var flattened = new List<string>();
        // Encoding, default font and primary language are handled by semantic decoding/resource remapping.
        Settings.DefaultTabWidthTwips = RetainMergeSetting(Settings.DefaultTabWidthTwips, from.DefaultTabWidthTwips, hasDestinationContent, nameof(Settings.DefaultTabWidthTwips), flattened);
        Settings.DefaultFarEastLanguageId = RetainMergeSetting(Settings.DefaultFarEastLanguageId, from.DefaultFarEastLanguageId, hasDestinationContent, nameof(Settings.DefaultFarEastLanguageId), flattened);
        Settings.DefaultAlternateLanguageId = RetainMergeSetting(Settings.DefaultAlternateLanguageId, from.DefaultAlternateLanguageId, hasDestinationContent, nameof(Settings.DefaultAlternateLanguageId), flattened);
        Settings.ViewKind = RetainMergeSetting(Settings.ViewKind, from.ViewKind, hasDestinationContent, nameof(Settings.ViewKind), flattened);
        Settings.ViewScale = RetainMergeSetting(Settings.ViewScale, from.ViewScale, hasDestinationContent, nameof(Settings.ViewScale), flattened);
        Settings.ZoomKind = RetainMergeSetting(Settings.ZoomKind, from.ZoomKind, hasDestinationContent, nameof(Settings.ZoomKind), flattened);
        Settings.ViewBackspaceBehavior = RetainMergeSetting(Settings.ViewBackspaceBehavior, from.ViewBackspaceBehavior, hasDestinationContent, nameof(Settings.ViewBackspaceBehavior), flattened);
        Settings.WidowOrphanControl = RetainMergeSetting(Settings.WidowOrphanControl, from.WidowOrphanControl, hasDestinationContent, nameof(Settings.WidowOrphanControl), flattened);
        Settings.AutoHyphenation = RetainMergeSetting(Settings.AutoHyphenation, from.AutoHyphenation, hasDestinationContent, nameof(Settings.AutoHyphenation), flattened);
        Settings.HyphenateCaps = RetainMergeSetting(Settings.HyphenateCaps, from.HyphenateCaps, hasDestinationContent, nameof(Settings.HyphenateCaps), flattened);
        Settings.ConsecutiveHyphenLimit = RetainMergeSetting(Settings.ConsecutiveHyphenLimit, from.ConsecutiveHyphenLimit, hasDestinationContent, nameof(Settings.ConsecutiveHyphenLimit), flattened);
        Settings.HyphenationZoneTwips = RetainMergeSetting(Settings.HyphenationZoneTwips, from.HyphenationZoneTwips, hasDestinationContent, nameof(Settings.HyphenationZoneTwips), flattened);
        Settings.FacingPages = RetainMergeSetting(Settings.FacingPages, from.FacingPages, hasDestinationContent, nameof(Settings.FacingPages), flattened);
        Settings.MirrorMargins = RetainMergeSetting(Settings.MirrorMargins, from.MirrorMargins, hasDestinationContent, nameof(Settings.MirrorMargins), flattened);
        Settings.FormProtection = RetainMergeSetting(Settings.FormProtection, from.FormProtection, hasDestinationContent, nameof(Settings.FormProtection), flattened);
        Settings.RevisionProtection = RetainMergeSetting(Settings.RevisionProtection, from.RevisionProtection, hasDestinationContent, nameof(Settings.RevisionProtection), flattened);
        Settings.AnnotationProtection = RetainMergeSetting(Settings.AnnotationProtection, from.AnnotationProtection, hasDestinationContent, nameof(Settings.AnnotationProtection), flattened);
        Settings.ReadOnlyProtection = RetainMergeSetting(Settings.ReadOnlyProtection, from.ReadOnlyProtection, hasDestinationContent, nameof(Settings.ReadOnlyProtection), flattened);
        Settings.TrackRevisions = RetainMergeSetting(Settings.TrackRevisions, from.TrackRevisions, hasDestinationContent, nameof(Settings.TrackRevisions), flattened);
        Settings.RevisionDisplayStyle = RetainMergeSetting(Settings.RevisionDisplayStyle, from.RevisionDisplayStyle, hasDestinationContent, nameof(Settings.RevisionDisplayStyle), flattened);
        Settings.RevisionBarPlacement = RetainMergeSetting(Settings.RevisionBarPlacement, from.RevisionBarPlacement, hasDestinationContent, nameof(Settings.RevisionBarPlacement), flattened);
        Settings.DrawingGridHorizontalSpacingTwips = RetainMergeSetting(Settings.DrawingGridHorizontalSpacingTwips, from.DrawingGridHorizontalSpacingTwips, hasDestinationContent, nameof(Settings.DrawingGridHorizontalSpacingTwips), flattened);
        Settings.DrawingGridVerticalSpacingTwips = RetainMergeSetting(Settings.DrawingGridVerticalSpacingTwips, from.DrawingGridVerticalSpacingTwips, hasDestinationContent, nameof(Settings.DrawingGridVerticalSpacingTwips), flattened);
        Settings.DrawingGridHorizontalOriginTwips = RetainMergeSetting(Settings.DrawingGridHorizontalOriginTwips, from.DrawingGridHorizontalOriginTwips, hasDestinationContent, nameof(Settings.DrawingGridHorizontalOriginTwips), flattened);
        Settings.DrawingGridVerticalOriginTwips = RetainMergeSetting(Settings.DrawingGridVerticalOriginTwips, from.DrawingGridVerticalOriginTwips, hasDestinationContent, nameof(Settings.DrawingGridVerticalOriginTwips), flattened);
        Settings.DrawingGridHorizontalShow = RetainMergeSetting(Settings.DrawingGridHorizontalShow, from.DrawingGridHorizontalShow, hasDestinationContent, nameof(Settings.DrawingGridHorizontalShow), flattened);
        Settings.DrawingGridVerticalShow = RetainMergeSetting(Settings.DrawingGridVerticalShow, from.DrawingGridVerticalShow, hasDestinationContent, nameof(Settings.DrawingGridVerticalShow), flattened);
        Settings.SnapToDrawingGrid = RetainMergeSetting(Settings.SnapToDrawingGrid, from.SnapToDrawingGrid, hasDestinationContent, nameof(Settings.SnapToDrawingGrid), flattened);
        Settings.DrawingGridUsesMargins = RetainMergeSetting(Settings.DrawingGridUsesMargins, from.DrawingGridUsesMargins, hasDestinationContent, nameof(Settings.DrawingGridUsesMargins), flattened);
        Settings.Direction = RetainMergeSetting(Settings.Direction, from.Direction, hasDestinationContent, nameof(Settings.Direction), flattened);
        if (flattened.Count > 0) report.Add(RtfConversionSeverity.Warning, "RtfMergeDocumentSettingsFlattened",
            "Independent source document settings cannot remain attached to imported content. The destination's document-wide settings are retained.",
            RtfConversionAction.Flattened, sourcePath: "Document/Settings", feature: "DocumentSettings", count: flattened.Count,
            detail: "Properties=" + string.Join(",", flattened));
    }

    private static T? RetainMergeSetting<T>(T? destination, T? source, bool hasDestinationContent, string name, ICollection<string> flattened) where T : struct {
        if (!hasDestinationContent && !destination.HasValue) return source;
        if (destination != null || source != null) {
            if (!Nullable.Equals(destination, source)) flattened.Add(name);
        }
        return destination;
    }
}
