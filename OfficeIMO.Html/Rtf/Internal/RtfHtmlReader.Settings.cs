namespace OfficeIMO.Html;

internal static partial class RtfHtmlReader {
    private sealed partial class ReadContext {
        private void ApplyDocumentSettings(Dictionary<string, string> values) {
            RtfDocumentSettings settings = _document.Settings;
            ApplyEnum<RtfDocumentCharacterSet>(values, "characterSet", value => settings.CharacterSet = value);
            ApplyInt(values, "ansiCodePage", value => settings.AnsiCodePage = value);
            ApplyInt(values, "unicodeSkipCount", value => settings.UnicodeSkipCount = value);
            ApplyInt(values, "defaultFontId", value => settings.DefaultFontId = value);
            ApplyInt(values, "defaultTabWidth", value => settings.DefaultTabWidthTwips = value);
            ApplyInt(values, "defaultLanguage", value => settings.DefaultLanguageId = value);
            ApplyInt(values, "defaultFarEastLanguage", value => settings.DefaultFarEastLanguageId = value);
            ApplyInt(values, "defaultAlternateLanguage", value => settings.DefaultAlternateLanguageId = value);
            ApplyInt(values, "viewKind", value => settings.ViewKind = value);
            ApplyInt(values, "viewScale", value => settings.ViewScale = value);
            ApplyInt(values, "zoomKind", value => settings.ZoomKind = value);
            ApplyInt(values, "viewBackspaceBehavior", value => settings.ViewBackspaceBehavior = value);
            ApplyBool(values, "widowOrphanControl", value => settings.WidowOrphanControl = value);
            ApplyBool(values, "autoHyphenation", value => settings.AutoHyphenation = value);
            ApplyBool(values, "hyphenateCaps", value => settings.HyphenateCaps = value);
            ApplyInt(values, "consecutiveHyphenLimit", value => settings.ConsecutiveHyphenLimit = value);
            ApplyInt(values, "hyphenationZone", value => settings.HyphenationZoneTwips = value);
            ApplyBool(values, "facingPages", value => settings.FacingPages = value);
            ApplyBool(values, "mirrorMargins", value => settings.MirrorMargins = value);
            ApplyBool(values, "formProtection", value => settings.FormProtection = value);
            ApplyBool(values, "revisionProtection", value => settings.RevisionProtection = value);
            ApplyBool(values, "annotationProtection", value => settings.AnnotationProtection = value);
            ApplyBool(values, "readOnlyProtection", value => settings.ReadOnlyProtection = value);
            ApplyBool(values, "trackRevisions", value => settings.TrackRevisions = value);
            ApplyInt(values, "revisionDisplayStyle", value => settings.RevisionDisplayStyle = value);
            ApplyInt(values, "revisionBarPlacement", value => settings.RevisionBarPlacement = value);
            ApplyInt(values, "drawingGrid.horizontalSpacing", value => settings.DrawingGridHorizontalSpacingTwips = value);
            ApplyInt(values, "drawingGrid.verticalSpacing", value => settings.DrawingGridVerticalSpacingTwips = value);
            ApplyInt(values, "drawingGrid.horizontalOrigin", value => settings.DrawingGridHorizontalOriginTwips = value);
            ApplyInt(values, "drawingGrid.verticalOrigin", value => settings.DrawingGridVerticalOriginTwips = value);
            ApplyInt(values, "drawingGrid.horizontalShow", value => settings.DrawingGridHorizontalShow = value);
            ApplyInt(values, "drawingGrid.verticalShow", value => settings.DrawingGridVerticalShow = value);
            ApplyBool(values, "drawingGrid.snapToGrid", value => settings.SnapToDrawingGrid = value);
            ApplyBool(values, "drawingGrid.useMargins", value => settings.DrawingGridUsesMargins = value);
            ApplyEnum<RtfTextDirection>(values, "direction", value => settings.Direction = value);
        }

        private static void ApplyInt(Dictionary<string, string> values, string key, Action<int?> assign) {
            if (values.ContainsKey(key)) {
                assign(ReadInt(values, key));
            }
        }

        private static void ApplyBool(Dictionary<string, string> values, string key, Action<bool?> assign) {
            if (values.ContainsKey(key)) {
                assign(ReadBool(values, key));
            }
        }

        private static void ApplyEnum<T>(Dictionary<string, string> values, string key, Action<T?> assign) where T : struct {
            if (values.ContainsKey(key)) {
                assign(ReadEnum<T>(values, key));
            }
        }
    }
}
