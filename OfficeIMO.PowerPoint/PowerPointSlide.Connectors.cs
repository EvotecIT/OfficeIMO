using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        private bool? _hasGluedConnectors;

        /// <summary>
        ///     Adds a connector glued to the closest facing connection sites of two shapes. Moving or resizing either
        ///     shape through this API keeps the connector attached.
        /// </summary>
        public PowerPointConnectionShape AddConnector(PowerPointShape from, PowerPointShape to, PowerPointConnectorKind kind) {
            ValidateGlueTarget(from, nameof(from));
            ValidateGlueTarget(to, nameof(to));
            (int start, int end) = PowerPointConnectorGeometry.ChooseSites(from, to);
            return AddConnector(from, start, to, end, kind);
        }

        /// <summary>
        ///     Adds a connector glued to the given connection sites (see <see cref="PowerPointShape.ConnectionSites" />).
        ///     Moving or resizing either shape through this API keeps the connector attached.
        /// </summary>
        public PowerPointConnectionShape AddConnector(PowerPointShape from, int fromSite, PowerPointShape to, int toSite,
            PowerPointConnectorKind kind) {
            ValidateGlueTarget(from, nameof(from));
            ValidateGlueTarget(to, nameof(to));
            if (ReferenceEquals(from, to)) throw new ArgumentException("A connector needs two different shapes.", nameof(to));
            PowerPointConnectorGeometry.GetSite(from, fromSite, nameof(fromSite));
            PowerPointConnectorGeometry.GetSite(to, toSite, nameof(toSite));

            PowerPointConnectionShape connector = AddConnectorShape();
            connector.GlueEnds(from.Id!.Value, fromSite, to.Id!.Value, toSite);
            _hasGluedConnectors = true;
            connector.RouteGlued(kind);
            return connector;
        }

        /// <summary>Adds a free connector between two points.</summary>
        public PowerPointConnectionShape AddConnector(OfficeLength startX, OfficeLength startY, OfficeLength endX,
            OfficeLength endY, PowerPointConnectorKind kind) {
            PowerPointConnectionShape connector = AddConnectorShape();
            PowerPointConnectorGeometry.Route(connector, kind, startX.ToEmus(), startY.ToEmus(),
                endX.ToEmus(), endY.ToEmus(), null, null);
            return connector;
        }

        internal PowerPointShape? FindTopLevelShape(uint id) =>
            ShapeList.FirstOrDefault(shape => shape.Id == id && shape.XmlElement.Parent.Is(OfficeXmlNamespaces.PresentationML, "spTree"));

        internal void RerouteConnectorsAttachedTo(PowerPointShape shape) {
            // Read through the element handle, so a slide without glued connectors stays off the SDK.
            if (shape is PowerPointConnectionShape || !shape.XmlElement.Parent.Is(OfficeXmlNamespaces.PresentationML, "spTree")) return;
            _hasGluedConnectors ??= HasGluedConnectors();
            if (_hasGluedConnectors != true || shape.Id is not uint id) return;
            foreach (PowerPointConnectionShape connector in ShapeList.OfType<PowerPointConnectionShape>()) {
                if (connector.StartShapeId == id || connector.EndShapeId == id) connector.RerouteAfterShapeChange();
            }
        }

        private bool HasGluedConnectors() {
            OfficeXmlNode tree = PartRef.Root.Child(OfficeXmlNamespaces.PresentationML, "cSld").Child(OfficeXmlNamespaces.PresentationML, "spTree");
            // A grouped connector can become active again after ungrouping.
            foreach (OfficeXmlNode properties in tree.ContentDescendants(OfficeXmlNamespaces.PresentationML, "cNvCxnSpPr")) {
                if (!properties.Parent.Is(OfficeXmlNamespaces.PresentationML, "nvCxnSpPr")) continue;
                if (!properties.Child(OfficeXmlNamespaces.DrawingML, "stCxn").IsNone || !properties.Child(OfficeXmlNamespaces.DrawingML, "endCxn").IsNone) return true;
            }
            return false;
        }

        private void ValidateGlueTarget(PowerPointShape shape, string parameterName) {
            ArgumentNullException.ThrowIfNull(shape, parameterName);
            if (!ReferenceEquals(shape.OwnerSlide, this) || !shape.XmlElement.Parent.Is(OfficeXmlNamespaces.PresentationML, "spTree") || shape.Id == null) {
                throw new ArgumentException("Connectors glue to top-level shapes on the same slide.", parameterName);
            }
        }

        private PowerPointConnectionShape AddConnectorShape() =>
            AppendConnectionShape(OfficePresetShapeType.StraightConnector1.ToOpenXmlToken(), 0L, 0L, 0L, 0L, standardOutline: true);

        // Used by the binary PowerPoint reader while it still supplies a typed geometry token.
        internal PowerPointConnectionShape AddConnectionShape(A.ShapeTypeValues shapeType, long left,
            long top, long width, long height) =>
            AppendConnectionShape(((DocumentFormat.OpenXml.IEnumValue)shapeType).Value, left, top, width, height, standardOutline: false);

        private PowerPointConnectionShape AppendConnectionShape(string preset, long left, long top, long width, long height, bool standardOutline) {
            const int p = OfficeXmlNamespaces.PresentationML, a = OfficeXmlNamespaces.DrawingML;
            uint id = AllocateShapeId();
            string name = GenerateUniqueName("Connector");
            PowerPointPartRef part = PartRef;
            OfficeXmlNode tree = EnsureCommonSlideData(part).GetOrAddChild(p, "spTree", PCommonSlideDataView.Order);
            OfficeXmlNode connector = tree.CreateElement(p, "cxnSp");
            OfficeXmlNode nonVisual = connector.Append(connector.CreateElement(p, "nvCxnSpPr"));
            OfficeXmlNode drawing = nonVisual.Append(nonVisual.CreateElement(p, "cNvPr"));
            drawing.SetUInt32Attribute("id", id);
            drawing.SetAttribute("name", name);
            nonVisual.Append(nonVisual.CreateElement(p, "cNvCxnSpPr"));
            nonVisual.Append(nonVisual.CreateElement(p, "nvPr"));

            var properties = new AShapePropertiesView(connector.Append(connector.CreateElement(p, "spPr")));
            ATransform2DView transform = properties.GetOrAddXfrm();
            APoint2DView offset = transform.GetOrAddOff();
            offset.Element.SetInt64Attribute("x", left);
            offset.Element.SetInt64Attribute("y", top);
            APositiveSize2DView extent = transform.GetOrAddExt();
            extent.Cx = width;
            extent.Cy = height;

            APresetGeometry2DView geometry = properties.GetOrAddPrstGeom();
            geometry.Prst = preset;
            geometry.GetOrAddAvLst();
            if (standardOutline) {
                OfficeXmlNode outline = properties.Element.Append(properties.Element.CreateElement(a, "ln"));
                outline.SetInt32Attribute("w", 12700);
                OfficeXmlNode fill = outline.Append(outline.CreateElement(a, "solidFill"));
                fill.Append(fill.CreateElement(a, "schemeClr")).SetAttribute("val", "tx1");
            }

            connector = tree.InsertBefore(connector, FindTerminalPresentationExtension(tree));
            part.MarkEdited();
            return TrackShape((PowerPointConnectionShape)WrapSlideShape(connector, part)!);
        }
    }
}
