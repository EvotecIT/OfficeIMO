namespace OfficeIMO.Visio;

/// <summary>Shared text-background resolution for the SVG, PNG and Drawing renderers.</summary>
internal static class VisioTextBackground {
    /// <summary>
    /// Returns the text style to paint with: the object's own style when it sets a background, otherwise a
    /// render-only copy carrying the background inherited from its text style sheet (or its master's).
    /// </summary>
    internal static VisioTextStyle? Effective(VisioShape shape) {
        OfficeColor? sheet = shape.StyleSheetTextBackground;
        if (sheet == null && string.IsNullOrEmpty(shape.NativeStyleReferences?.TextStyle))
            sheet = (shape.MasterShape ?? shape.Master?.Shape)?.StyleSheetTextBackground;
        return Effective(shape.TextStyle, sheet);
    }

    internal static VisioTextStyle? Effective(VisioConnector connector) =>
        Effective(connector.TextStyle, connector.StyleSheetTextBackground);

    private static VisioTextStyle? Effective(VisioTextStyle? style, OfficeColor? sheet) {
        if (style?.BackgroundColor != null || sheet == null) return style;
        VisioTextStyle copy = style?.Clone() ?? new VisioTextStyle();
        OfficeColor inherited = sheet.Value;
        if (inherited.A > 0 && style?.BackgroundTransparency != null) {
            // A local TextBkgndTrans overrides the sheet's transparency; the renderers apply it to the opaque color.
            copy.BackgroundColor = OfficeColor.FromRgb(inherited.R, inherited.G, inherited.B);
            copy.BackgroundTransparency = style.BackgroundTransparency;
        } else {
            // The sheet color already carries TextBkgndTrans in its alpha.
            copy.BackgroundColor = inherited;
            copy.BackgroundTransparency = null;
        }
        return copy;
    }
}
