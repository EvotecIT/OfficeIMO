using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.PowerPoint.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint;

public partial class PowerPointSlide {
    /// <summary>The SDK form of <see cref="GetExportPlaceholderBounds(PowerPointPartRef, OfficeXmlNode)"/>, for the layout engine.</summary>
    internal static PowerPointLayoutBox? GetExportPlaceholderBounds(SlideLayoutPart? layoutPart, PlaceholderShape placeholder) =>
        GetExportPlaceholderBounds(PowerPointPartRef.FromSdk(layoutPart), OfficeXmlNode.FromSdk(placeholder));
}
