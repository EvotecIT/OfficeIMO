namespace OfficeIMO.OpenDocument;

internal sealed partial class OdfStyleImportPlan {
    private static readonly XName[] EdgeShorthands = {
        OdfNamespaces.Fo + "padding", OdfNamespaces.Fo + "margin", OdfNamespaces.Fo + "border", OdfNamespaces.Style + "border-line-width"
    };
    private static readonly string[] EdgeSides = { "top", "right", "bottom", "left" };

    /// <summary>Captures only fallback sides not covered by an explicit side or shorthand in the original cascade.</summary>
    /// <remarks>
    /// A snapshot can sit above a named parent. Expanding a fallback shorthand into missing sides prevents it from
    /// overriding an explicit side inherited from that parent. Each fallback level resolves its side before its shorthand.
    /// </remarks>
    private static void CaptureEdgeFallbacks(XElement target, XElement fallback, IEnumerable<XElement> explicitCascade) {
        XElement[] explicitProperties = explicitCascade.ToArray();
        foreach (XName shorthand in EdgeShorthands) {
            foreach (string side in EdgeSides) {
                XName name = shorthand.Namespace + (shorthand.LocalName + "-" + side);
                if (target.Attribute(name) != null || target.Attribute(shorthand) != null ||
                    explicitProperties.Any(properties => properties.Attribute(name) != null || properties.Attribute(shorthand) != null)) continue;
                XAttribute? inherited = fallback.Attribute(name) ?? fallback.Attribute(shorthand);
                if (inherited != null) target.Add(new XAttribute(name, inherited.Value));
            }
        }
    }

    private static bool IsEdgeProperty(XName name) => EdgeShorthands.Any(shorthand => name == shorthand ||
        name.Namespace == shorthand.Namespace && EdgeSides.Any(side => name.LocalName == shorthand.LocalName + "-" + side));
}
