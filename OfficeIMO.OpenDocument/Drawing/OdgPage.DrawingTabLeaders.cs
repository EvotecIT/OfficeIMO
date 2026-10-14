using OfficeIMO.Drawing;

namespace OfficeIMO.OpenDocument;

public sealed partial class OdgPage {
    private static OfficeTextTabStop ProjectTabLeader(XElement element, OfficeTextTabStop stop, OdfTextParagraph paragraph, OdfStyle tabStyle, HashSet<string> losses) {
        string? text = (string?)element.Attribute(OdfNamespaces.Style + "leader-text");
        if (text != null) {
            // A declared text character takes precedence over line-pattern attributes.
            try {
                OfficeTextTabStop result = stop.WithLeader(text);
                string? styleName = (string?)element.Attribute(OdfNamespaces.Style + "leader-text-style");
                return styleName == null ? result : ProjectStyledTabLeader(result, paragraph, tabStyle, styleName, losses);
            } catch (Exception e) when (e is ArgumentException or FormatException or NotSupportedException or OverflowException or InvalidDataException) {
                losses.Add("tab-leaders"); return stop;
            }
        }
        string? style = (string?)element.Attribute(OdfNamespaces.Style + "leader-style");
        if (style is null or "none") return stop;
        try {
            string? nativeColor = (string?)element.Attribute(OdfNamespaces.Style + "leader-color");
            OdfColor? color = nativeColor is null or "font-color" ? null : OdfColor.Parse(nativeColor);
            // leader-text-style only applies to a declared text character, never a line.
            return stop.WithLineLeader(new OdfTabLineLeader(style,
                (string?)element.Attribute(OdfNamespaces.Style + "leader-type") ?? "single",
                (string?)element.Attribute(OdfNamespaces.Style + "leader-width") ?? "auto", color).ToDrawingLeader());
        } catch (Exception e) when (e is ArgumentException or FormatException or NotSupportedException or OverflowException) {
            losses.Add("tab-leaders"); return stop;
        }
    }
}
