using OfficeIMO.Drawing;

namespace OfficeIMO.OpenDocument;

public sealed partial class OdfTextParagraph {
    /// <summary>Replaces the paragraph's explicit tab set. An empty set overrides all inherited explicit stops.</summary>
    /// <remarks>At most 256 unique absolute positions are accepted. Existing leader definitions in this paragraph's overridden tab set are replaced. Default spacing belongs to the document's default paragraph style.</remarks>
    public void SetTabStops(IReadOnlyList<OdfTabStop> stops) {
        if (stops == null) throw new ArgumentNullException(nameof(stops));
        if (stops.Count > 256) throw new ArgumentException("At most 256 stops are supported.", nameof(stops));
        var positions = new HashSet<double>();
        var container = new XElement(OdfNamespaces.Style + "tab-stops");
        foreach (OdfTabStop stop in stops) {
            if (stop == null || !positions.Add(stop.Position.ToPoints())) throw new ArgumentException("Tab stops must be non-null and have unique positions.", nameof(stops));
            if (stop.LeaderTextStyle is { } leader && !ReferenceEquals(
                Document.Styles.FindInPart(OdfStyleFamily.Text, leader.Name, PartPath)?.Element, leader.Element))
                throw new ArgumentException("A leader text style must belong to this document and resolve in the paragraph's package part.", nameof(stops));
            container.Add(stop.ToElement());
        }
        OdfStyle style = EnsureStyle();
        XElement properties = style.GetProperties(OdfNamespaces.Style + "paragraph-properties");
        properties.Elements(OdfNamespaces.Style + "tab-stops").Remove();
        properties.Add(container); Document.MarkPartDirty(style.PartPath);
    }
}
