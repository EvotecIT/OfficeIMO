using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.OpenXml;
using Ns = OfficeIMO.Core.Internal.OpenXml.OfficeXmlNamespaces;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointChart {
        /// <summary>
        ///     The chart element (<c>c:chartSpace/c:chart</c>) of the chart part the frame references, wherever the part
        ///     is held, for the layout system's chart styling.
        /// </summary>
        internal OfficeXmlNode ChartNode {
            get {
                string? id = XmlElement.Child(Ns.DrawingML, "graphic").Child(Ns.DrawingML, "graphicData").Child(Ns.Chart, "chart")
                    .Attribute(Ns.Relationships, "id");
                if (id == null) throw new InvalidOperationException("Chart reference not found for the shape.");
                if (id.Length == 0) throw new InvalidOperationException("Chart relationship id is empty.");
                if (!OwnerPartRef.TryGetRelationship(id, out PowerPointRelationship relationship) || relationship.TargetPart.IsNone)
                    throw new InvalidOperationException("Chart part not found for the shape.");
                OfficeXmlNode chart = relationship.TargetPart.Root.Child(Ns.Chart, "chart");
                return chart.Exists ? chart : throw new InvalidOperationException("Chart element not found in chart part.");
            }
        }
    }
}
