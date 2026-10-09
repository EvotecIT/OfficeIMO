namespace OfficeIMO.OpenDocument;

public abstract partial class OdfShape {
    /// <summary>Inherited native text fitting; null removes both local fitting declarations.</summary>
    /// <remarks>
    /// The nearest graphic style declaring either <c>draw:fit-to-size</c> or <c>style:shrink-to-fit</c> owns the complete mode.
    /// Setting writes both canonical boolean declarations, so an inherited opposite mode cannot remain active.
    /// Shared drawing projection approximates shrink-to-fit only for fixed ordinary text boxes and rectangle labels.
    /// Its uniform font scale stops when the largest declared font reaches six points (or its original size if smaller);
    /// absolute paragraph metrics remain fixed. Clipping is reported, and later font-provider changes can change fitting.
    /// Stretching, fitting combined with active auto-growth, line captions and other shape kinds remain outside that projection profile.
    /// </remarks>
    /// <exception cref="NotSupportedException">The effective native pair contains unknown, legacy or conflicting active values.</exception>
    /// <exception cref="ArgumentOutOfRangeException">The assigned mode is not defined.</exception>
    public OdfTextFitMode? TextFitMode {
        get {
            foreach (OdfStyle style in Document.Styles.ResolveWithDefault(GetGraphicStyle(), OdfStyleFamily.Graphic)) {
                XElement? properties = style.Element.Element(OdfNamespaces.Style + "graphic-properties");
                string? stretch = (string?)properties?.Attribute(OdfNamespaces.Draw + "fit-to-size");
                string? shrink = (string?)properties?.Attribute(OdfNamespaces.Style + "shrink-to-fit");
                if (stretch == null && shrink == null) continue;
                if (stretch is not (null or "false" or "true") || shrink is not (null or "false" or "true") ||
                    stretch == "true" && shrink == "true")
                    throw new NotSupportedException("Unsupported ODF text-fitting declarations.");
                return stretch == "true" ? OdfTextFitMode.Stretch : shrink == "true" ? OdfTextFitMode.ShrinkToFit : OdfTextFitMode.None;
            }
            return null;
        }
        set {
            if (value.HasValue && !Enum.IsDefined(typeof(OdfTextFitMode), value.Value)) throw new ArgumentOutOfRangeException(nameof(value));
            OdfStyle style = EnsureGraphicStyle();
            style.SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Draw + "fit-to-size",
                value.HasValue ? value == OdfTextFitMode.Stretch ? "true" : "false" : null);
            style.SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Style + "shrink-to-fit",
                value.HasValue ? value == OdfTextFitMode.ShrinkToFit ? "true" : "false" : null);
        }
    }
}
