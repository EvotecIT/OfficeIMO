using A = DocumentFormat.OpenXml.Drawing;
using H = DocumentFormat.OpenXml.Office2019.Drawing.HyperLinkColor;

namespace OfficeIMO.PowerPoint;

/// <summary>Writes the native hyperlink color choice without changing targets, sounds or unrelated extensions.</summary>
internal static class PowerPointTextHyperlinkColor {
    private const string ExtensionUri = "{A12FA001-AC4F-418D-AE19-62706E023703}";

    internal static bool? ReadChoice(A.HyperlinkType hyperlink) {
        H.HyperlinkColor? color = hyperlink.HyperlinkExtensionList?
            .Elements<A.HyperlinkExtension>().FirstOrDefault(extension =>
                string.Equals(extension.Uri?.Value, ExtensionUri, StringComparison.OrdinalIgnoreCase))?
            .GetFirstChild<H.HyperlinkColor>();
        if (color?.Val?.Value == H.HyperlinkColorEnum.Tx) return true;
        if (color?.Val?.Value == H.HyperlinkColorEnum.HLink) return false;
        return null;
    }

    internal static void SetChoice(A.HyperlinkType hyperlink, bool useTextColor) {
        A.HyperlinkExtensionList list = hyperlink.HyperlinkExtensionList ?? new A.HyperlinkExtensionList();
        if (list.Parent == null) hyperlink.AddChild(list, true);
        A.HyperlinkExtension? extension = list.Elements<A.HyperlinkExtension>().FirstOrDefault(candidate =>
            string.Equals(candidate.Uri?.Value, ExtensionUri, StringComparison.OrdinalIgnoreCase));
        if (extension == null) {
            extension = new A.HyperlinkExtension { Uri = ExtensionUri };
            list.Append(extension);
        }
        extension.RemoveAllChildren<H.HyperlinkColor>();
        extension.Append(new H.HyperlinkColor {
            Val = useTextColor ? H.HyperlinkColorEnum.Tx : H.HyperlinkColorEnum.HLink
        });
    }

    internal static void UseExplicitTextColor(A.TextCharacterPropertiesType properties) {
        foreach (A.HyperlinkType hyperlink in properties.Elements<A.HyperlinkType>()) {
            SetChoice(hyperlink, true);
        }
    }
}
