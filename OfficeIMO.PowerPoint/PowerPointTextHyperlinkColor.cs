using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint;

/// <summary>Writes the native hyperlink color choice without changing targets, sounds or unrelated extensions.</summary>
internal static partial class PowerPointTextHyperlinkColor {
    private const string ExtensionUri = "{A12FA001-AC4F-418D-AE19-62706E023703}";
    private const string ColorNamespace = "http://schemas.microsoft.com/office/drawing/2018/hyperlinkcolor";

    // The a:hlinkClick and a:hlinkHover children, in sequence order.
    private static readonly OfficeXmlChildOrder HyperlinkOrder = new(
        new[] { OfficeXmlNamespaces.DrawingML, OfficeXmlNamespaces.DrawingML }, new[] { "snd", "extLst" });

    internal static bool? ReadChoice(PowerPointXmlNode hyperlink) {
        PowerPointXmlNode color = FindExtension(hyperlink).Child(ColorNamespace, "hlinkClr");
        return color.Attribute("val") switch {
            "tx" => true,
            "hlink" => false,
            _ => null
        };
    }

    internal static void SetChoice(PowerPointXmlNode hyperlink, bool useTextColor) {
        PowerPointXmlNode list = hyperlink.GetOrAddChild(OfficeXmlNamespaces.DrawingML, "extLst", HyperlinkOrder);
        PowerPointXmlNode extension = FindExtension(hyperlink);
        if (extension.IsNone) {
            extension = list.Append(list.CreateElement(OfficeXmlNamespaces.DrawingML, "ext"));
            extension.SetAttribute("uri", ExtensionUri);
        }
        foreach (PowerPointXmlNode existing in extension.Children().ToList()) {
            if (existing.Is(ColorNamespace, "hlinkClr")) existing.Remove();
        }
        PowerPointXmlNode color = extension.CreateElement(ColorNamespace, "hlinkClr", "ahyp");
        color.SetAttribute("val", useTextColor ? "tx" : "hlink");
        extension.Append(color);
    }

    internal static void UseExplicitTextColor(PowerPointXmlNode properties) {
        foreach (PowerPointXmlNode child in properties.Children().ToList()) {
            if (child.Is(OfficeXmlNamespaces.DrawingML, "hlinkClick") || child.Is(OfficeXmlNamespaces.DrawingML, "hlinkMouseOver")) SetChoice(child, true);
        }
    }

    private static PowerPointXmlNode FindExtension(PowerPointXmlNode hyperlink) {
        foreach (PowerPointXmlNode extension in hyperlink.Child(OfficeXmlNamespaces.DrawingML, "extLst").Children(OfficeXmlNamespaces.DrawingML, "ext")) {
            if (string.Equals(extension.Attribute("uri"), ExtensionUri, StringComparison.OrdinalIgnoreCase)) return extension;
        }
        return default;
    }
}
