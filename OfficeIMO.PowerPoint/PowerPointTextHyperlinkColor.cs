using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.PowerPoint.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint;

/// <summary>Writes the native hyperlink color choice without changing targets, sounds or unrelated extensions.</summary>
internal static partial class PowerPointTextHyperlinkColor {
    private const string ExtensionUri = "{A12FA001-AC4F-418D-AE19-62706E023703}";
    private const string ColorNamespace = "http://schemas.microsoft.com/office/drawing/2018/hyperlinkcolor";

    // The a:hlinkClick and a:hlinkHover children, in sequence order.
    private static readonly OfficeXmlChildOrder HyperlinkOrder = new(
        new[] { OfficeXmlNamespaces.DrawingML, OfficeXmlNamespaces.DrawingML }, new[] { "snd", "extLst" });

    internal static bool? ReadChoice(OfficeXmlNode hyperlink) {
        OfficeXmlNode color = FindExtension(hyperlink).Child(ColorNamespace, "hlinkClr");
        return color.Attribute("val") switch {
            "tx" => true,
            "hlink" => false,
            _ => null
        };
    }

    internal static void SetChoice(OfficeXmlNode hyperlink, bool useTextColor) {
        OfficeXmlNode list = hyperlink.GetOrAddChild(OfficeXmlNamespaces.DrawingML, "extLst", HyperlinkOrder);
        OfficeXmlNode extension = FindExtension(hyperlink);
        if (extension.IsNone) {
            extension = list.Append(list.CreateElement(OfficeXmlNamespaces.DrawingML, "ext"));
            extension.SetAttribute("uri", ExtensionUri);
        }
        foreach (OfficeXmlNode existing in extension.Children().ToList()) {
            if (existing.Is(ColorNamespace, "hlinkClr")) existing.Remove();
        }
        OfficeXmlNode color = extension.CreateElement(ColorNamespace, "hlinkClr", "ahyp");
        color.SetAttribute("val", useTextColor ? "tx" : "hlink");
        extension.Append(color);
    }

    internal static void UseExplicitTextColor(OfficeXmlNode properties) {
        foreach (OfficeXmlNode child in properties.Children().ToList()) {
            if (child.Is(OfficeXmlNamespaces.DrawingML, "hlinkClick") || child.Is(OfficeXmlNamespaces.DrawingML, "hlinkMouseOver")) SetChoice(child, true);
        }
    }

    private static OfficeXmlNode FindExtension(OfficeXmlNode hyperlink) {
        foreach (OfficeXmlNode extension in hyperlink.Child(OfficeXmlNamespaces.DrawingML, "extLst").Children(OfficeXmlNamespaces.DrawingML, "ext")) {
            if (string.Equals(extension.Attribute("uri"), ExtensionUri, StringComparison.OrdinalIgnoreCase)) return extension;
        }
        return default;
    }
}
