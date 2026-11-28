using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint;

public partial class PowerPointSlide {
    /// <summary>Gets or sets the authored slide name stored in common slide data.</summary>
    public string? Name {
        get => PartRef.Root.Child(OfficeXmlNamespaces.PresentationML, "cSld").Attribute("name");
        set {
            PowerPointPartRef part = PartRef;
            EnsureCommonSlideData(part).SetAttribute("name", string.IsNullOrEmpty(value) ? null : value);
            part.MarkEdited();
        }
    }

    private static OfficeXmlNode EnsureCommonSlideData(PowerPointPartRef part) {
        const int presentation = OfficeXmlNamespaces.PresentationML;
        OfficeXmlNode root = part.Root;
        OfficeXmlNode data = root.Child(presentation, "cSld");
        if (data.Exists) return data;
        data = root.AddOrdered(presentation, "cSld", PSlideView.Order);
        data.Append(data.CreateElement(presentation, "spTree"));
        part.MarkEdited();
        return data;
    }
}
