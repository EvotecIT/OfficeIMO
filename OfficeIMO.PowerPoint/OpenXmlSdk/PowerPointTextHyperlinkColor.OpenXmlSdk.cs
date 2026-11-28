using OfficeIMO.PowerPoint.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint;

/// <summary>The SDK forms of the hyperlink color helpers, for code that still holds SDK elements.</summary>
internal static partial class PowerPointTextHyperlinkColor {
    internal static bool? ReadChoice(A.HyperlinkType hyperlink) => ReadChoice(OfficeXmlNode.FromSdk(hyperlink));

    internal static void SetChoice(A.HyperlinkType hyperlink, bool useTextColor) => SetChoice(OfficeXmlNode.FromSdk(hyperlink), useTextColor);
}
