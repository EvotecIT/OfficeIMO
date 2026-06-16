namespace OfficeIMO.Html;

internal static partial class RtfHtmlWriter {
    private static void AppendDocumentSettingsMetadata(StringBuilder builder, RtfDocument document, string newline) {
        var values = new Dictionary<string, string>(StringComparer.Ordinal);
        AddDocumentSettings(values, document.Settings);
        if (values.Count == 0) {
            return;
        }

        builder.Append(newline);
        builder.Append("<meta name=\"officeimo-rtf-document-settings\" content=\"");
        builder.Append(EncodeAttribute(RtfHtmlMetadataCodec.Encode(values)));
        builder.Append("\">");
    }

    private static void AddDocumentSettings(Dictionary<string, string> values, RtfDocumentSettings settings) {
        AddEnum(values, "characterSet", settings.CharacterSet);
        AddNullableInt(values, "ansiCodePage", settings.AnsiCodePage);
        AddNullableInt(values, "unicodeSkipCount", settings.UnicodeSkipCount);
        AddNullableInt(values, "defaultFontId", settings.DefaultFontId);
        AddNullableInt(values, "defaultTabWidth", settings.DefaultTabWidthTwips);
        AddNullableInt(values, "defaultLanguage", settings.DefaultLanguageId);
        AddNullableInt(values, "defaultFarEastLanguage", settings.DefaultFarEastLanguageId);
        AddNullableInt(values, "defaultAlternateLanguage", settings.DefaultAlternateLanguageId);
        AddNullableInt(values, "viewKind", settings.ViewKind);
        AddNullableInt(values, "viewScale", settings.ViewScale);
        AddNullableInt(values, "zoomKind", settings.ZoomKind);
        AddNullableInt(values, "viewBackspaceBehavior", settings.ViewBackspaceBehavior);
        AddNullableBool(values, "widowOrphanControl", settings.WidowOrphanControl);
        AddNullableBool(values, "autoHyphenation", settings.AutoHyphenation);
        AddNullableBool(values, "hyphenateCaps", settings.HyphenateCaps);
        AddNullableInt(values, "consecutiveHyphenLimit", settings.ConsecutiveHyphenLimit);
        AddNullableInt(values, "hyphenationZone", settings.HyphenationZoneTwips);
        AddNullableBool(values, "facingPages", settings.FacingPages);
        AddNullableBool(values, "mirrorMargins", settings.MirrorMargins);
        AddNullableBool(values, "formProtection", settings.FormProtection);
        AddNullableBool(values, "revisionProtection", settings.RevisionProtection);
        AddNullableBool(values, "annotationProtection", settings.AnnotationProtection);
        AddNullableBool(values, "readOnlyProtection", settings.ReadOnlyProtection);
        AddNullableBool(values, "trackRevisions", settings.TrackRevisions);
        AddNullableInt(values, "revisionDisplayStyle", settings.RevisionDisplayStyle);
        AddNullableInt(values, "revisionBarPlacement", settings.RevisionBarPlacement);
        AddNullableInt(values, "drawingGrid.horizontalSpacing", settings.DrawingGridHorizontalSpacingTwips);
        AddNullableInt(values, "drawingGrid.verticalSpacing", settings.DrawingGridVerticalSpacingTwips);
        AddNullableInt(values, "drawingGrid.horizontalOrigin", settings.DrawingGridHorizontalOriginTwips);
        AddNullableInt(values, "drawingGrid.verticalOrigin", settings.DrawingGridVerticalOriginTwips);
        AddNullableInt(values, "drawingGrid.horizontalShow", settings.DrawingGridHorizontalShow);
        AddNullableInt(values, "drawingGrid.verticalShow", settings.DrawingGridVerticalShow);
        AddNullableBool(values, "drawingGrid.snapToGrid", settings.SnapToDrawingGrid);
        AddNullableBool(values, "drawingGrid.useMargins", settings.DrawingGridUsesMargins);
        AddEnum(values, "direction", settings.Direction);
    }
}
