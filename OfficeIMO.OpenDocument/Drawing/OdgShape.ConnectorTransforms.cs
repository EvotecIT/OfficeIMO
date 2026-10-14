using OfficeIMO.Drawing;

namespace OfficeIMO.OpenDocument;

public sealed partial class OdgShape {
    /// <summary>
    /// Bakes a free connector's own affine transform into its endpoints and saved route, including curve controls.
    /// Clears the transform and native routing offsets; preserves the routing kind, declared stroke and other styles.
    /// Scaling changes geometry without scaling the declared stroke width. Attached ends, transformed parents
    /// and rotated, scaled or sheared connector labels are outside this profile. Translation retains labels.
    /// Unsupported input leaves the document unchanged.
    /// </summary>
    public void BakeConnectorTransform() {
        RequireConnector();
        if (Element.Ancestors(OdfNamespaces.Draw + "g").Any(group =>
            OdfDrawingTransform.Parse((string?)group.Attribute(OdfNamespaces.Draw + "transform")) != OfficeTransform.Identity))
            throw new NotSupportedException("Bake parent transforms with TransformChildren before baking the connector's own transform.");
        XAttribute[] attributes = PrepareFreeConnectorTransform(OdfDrawingTransform.Parse(Transform));
        ApplyBakedConnectorRoute(attributes);
        Dirty();
    }

    private XAttribute[] PrepareFreeConnectorTransform(OfficeTransform transform) {
        if (StartShapeId != null || EndShapeId != null)
            throw new NotSupportedException("Free connector transform baking requires both endpoints to be detached.");
        if (!IsTranslationTransform(transform) && HasConnectorLabel)
            throw new NotSupportedException("Baking connector labels supports translations; rotation, scale and shear require native label placement.");
        var path = ConnectorRouteCommands.Select(command => MapConnectorCommand(command, transform.TransformPoint)).ToArray();
        return EncodeConnectorRoute(path);
    }

    private void ApplyBakedConnectorRoute(IEnumerable<XAttribute> attributes) {
        foreach (XAttribute attribute in attributes) Element.SetAttributeValue(attribute.Name, attribute.Value);
        Element.Attribute(OdfNamespaces.Draw + "transform")?.Remove();
        Element.Attribute(OdfNamespaces.Draw + "line-skew")?.Remove();
    }

    // Native producers commonly emit an empty paragraph on an unlabeled connector. Keep that XML;
    // child elements may contain fields or explicit spaces whose displayed text is not in XElement.Value.
    private bool HasConnectorLabel => TextRoot.Elements().Where(IsTextContainer).Any(container =>
        !IsParagraph(container) || container.HasElements || !string.IsNullOrWhiteSpace(container.Value));

    internal static bool IsTranslationTransform(OfficeTransform transform) =>
        transform.M11 == 1 && transform.M12 == 0 && transform.M21 == 0 && transform.M22 == 1;
}
