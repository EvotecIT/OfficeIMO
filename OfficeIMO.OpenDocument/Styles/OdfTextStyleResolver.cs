namespace OfficeIMO.OpenDocument;

/// <summary>Resolves native inline, paragraph and enclosing graphic text defaults.</summary>
internal static class OdfTextStyleResolver {
    /// <summary>Resolves percentage font sizes against the first absolute base in the original style cascade.</summary>
    internal static double ResolveFontSize(IEnumerable<OdfStyle> styles, HashSet<string>? losses = null, XName? sizeAttribute = null) {
        var basis = ResolveFontSizeBasis(styles, losses, sizeAttribute);
        double size = (basis.Points ?? 12) * basis.Factor;
        if (size <= 0 || double.IsNaN(size) || double.IsInfinity(size)) throw new NotSupportedException("Text font size must be finite and positive.");
        return size;
    }

    /// <summary>Compiles a size cascade while retaining a relative-only result for contextual leader formatting.</summary>
    internal static (double? Points, double Factor) ResolveFontSizeBasis(IEnumerable<OdfStyle> styles, HashSet<string>? losses = null, XName? sizeAttribute = null) {
        XName attribute = sizeAttribute ?? OdfNamespaces.Fo + "font-size";
        double factor = 1; double? size = null;
        foreach (OdfStyle style in styles) {
            string? relative = (string?)style.TextProperties?.Attribute(OdfNamespaces.Style + "font-size-rel");
            if (relative != null && (!OdfLength.Parse(relative).TryToPoints(out double change) || change != 0)) losses?.Add("relative-font-size");
            string? lexical = (string?)style.TextProperties?.Attribute(attribute);
            if (lexical == null) continue;
            lexical = OdfLength.Parse(lexical).ToString();
            double value;
            if (lexical.EndsWith("%", StringComparison.Ordinal) && double.TryParse(lexical.Substring(0, lexical.Length - 1), NumberStyles.Float, CultureInfo.InvariantCulture, out double percent)) {
                value = percent / 100D; factor *= value;
            } else { value = OdfLength.Parse(lexical).ToPoints(); size = value; }
            if (value <= 0 || double.IsNaN(value) || double.IsInfinity(value)) throw new NotSupportedException("Text size and line spacing must be finite positive lengths.");
            if (!lexical.EndsWith("%", StringComparison.Ordinal)) break;
        }
        if (factor <= 0 || double.IsNaN(factor) || double.IsInfinity(factor)) throw new NotSupportedException("Text size multiplier must be finite and positive.");
        return (size, factor);
    }

    internal static IEnumerable<OdfStyle> Resolve(OdfStyleRepository styles, XElement element, XElement graphic, string part = "content.xml") {
        bool hasTextStyle = false;
        foreach (OdfStyle style in OdfInlineStyleResolver.InlineStyles(styles, element, part)) { hasTextStyle = true; yield return style; }
        XElement? paragraph = element.AncestorsAndSelf().TakeWhile(e => !ReferenceEquals(e, graphic)).FirstOrDefault(OdfTextTraversal.IsParagraph);
        if (paragraph != null) foreach (OdfStyle style in Referenced(styles, paragraph, OdfNamespaces.Text + "style-name", OdfStyleFamily.Paragraph, part)) yield return style;
        foreach (OdfStyle style in Referenced(styles, graphic, OdfNamespaces.Draw + "text-style-name", OdfStyleFamily.Paragraph, part)) yield return style;
        foreach (OdfStyle style in Referenced(styles, graphic, OdfNamespaces.Draw + "style-name", OdfStyleFamily.Graphic, part)) yield return style;
        OdfStyle? familyDefault = styles.FindDefault(hasTextStyle ? OdfStyleFamily.Text : OdfStyleFamily.Paragraph);
        if (familyDefault != null) yield return familyDefault;
        OdfStyle? graphicDefault = styles.FindDefault(OdfStyleFamily.Graphic);
        if (graphicDefault != null) yield return graphicDefault;
    }

    private static IEnumerable<OdfStyle> Referenced(OdfStyleRepository styles, XElement element, XName attribute, OdfStyleFamily family, string part) {
        string? name = (string?)element.Attribute(attribute);
        OdfStyle? style = name == null ? null : styles.FindInPart(family, name, part);
        return style == null ? Array.Empty<OdfStyle>() : styles.Resolve(style);
    }
}
