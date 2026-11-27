using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    /// <summary>The SDK forms of the hyperlink helpers, for code that still holds SDK parts and elements.</summary>
    internal static partial class PowerPointHyperlinkResolver {
        internal static Uri? Resolve(OpenXmlPartContainer ownerPart, SlidePart? sourceSlidePart, A.HyperlinkOnClick? hyperlink) =>
            Resolve(PowerPointPartRef.FromSdk(ownerPart as OpenXmlPart), PowerPointPartRef.FromSdk(sourceSlidePart), OfficeXmlNode.FromSdk(hyperlink));

        internal static Uri CreatePartRelativeUri(OpenXmlPart sourcePart, OpenXmlPart targetPart) =>
            CreatePartRelativeUri(sourcePart.Uri.OriginalString, targetPart.Uri.OriginalString);
    }
}
