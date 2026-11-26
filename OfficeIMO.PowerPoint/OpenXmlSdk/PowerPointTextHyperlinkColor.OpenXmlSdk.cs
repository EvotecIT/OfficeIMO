using OfficeIMO.PowerPoint.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint;

/// <summary>The SDK forms of the hyperlink color helpers, for code that still holds SDK elements.</summary>
internal static partial class PowerPointTextHyperlinkColor {
    internal static bool? ReadChoice(A.HyperlinkType hyperlink) => ReadChoice(PowerPointXmlNode.FromSdk(hyperlink));

    internal static void SetChoice(A.HyperlinkType hyperlink, bool useTextColor) => SetChoice(PowerPointXmlNode.FromSdk(hyperlink), useTextColor);
}
