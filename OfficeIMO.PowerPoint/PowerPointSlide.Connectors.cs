using DocumentFormat.OpenXml.Presentation;
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
            var drawing = (NonVisualConnectorShapeDrawingProperties)connector.ConnectorDrawingProperties;
            drawing.StartConnection = new A.StartConnection { Id = from.Id!.Value, Index = (uint)fromSite };
            drawing.EndConnection = new A.EndConnection { Id = to.Id!.Value, Index = (uint)toSite };
            _hasGluedConnectors = true;
            connector.RouteGlued(kind);
            return connector;
        }

        /// <summary>Adds a free connector between two points.</summary>
        public PowerPointConnectionShape AddConnector(OfficeLength startX, OfficeLength startY, OfficeLength endX,
            OfficeLength endY, PowerPointConnectorKind kind) {
            PowerPointConnectionShape connector = AddConnectorShape();
            PowerPointConnectorGeometry.Route((ConnectionShape)connector.Element, kind, startX.ToEmus(), startY.ToEmus(),
                endX.ToEmus(), endY.ToEmus(), null, null);
            return connector;
        }

        internal PowerPointShape? FindTopLevelShape(uint id) =>
            ShapeList.FirstOrDefault(shape => shape.Id == id && shape.Element.Parent is ShapeTree);

        internal void RerouteConnectorsAttachedTo(PowerPointShape shape) {
            if (shape is PowerPointConnectionShape || shape.Element.Parent is not ShapeTree) return;
            _hasGluedConnectors ??= SlideRoot.CommonSlideData?.ShapeTree?.Descendants<NonVisualConnectorShapeDrawingProperties>()
                .Any(p => p.StartConnection != null || p.EndConnection != null) == true;
            if (_hasGluedConnectors != true || shape.Id is not uint id) return;
            foreach (PowerPointConnectionShape connector in ShapeList.OfType<PowerPointConnectionShape>()) {
                if (connector.StartShapeId == id || connector.EndShapeId == id) connector.RerouteAfterShapeChange();
            }
        }

        private void ValidateGlueTarget(PowerPointShape shape, string parameterName) {
            ArgumentNullException.ThrowIfNull(shape, parameterName);
            if (!ReferenceEquals(shape.OwnerSlide, this) || shape.Element.Parent is not ShapeTree || shape.Id == null) {
                throw new ArgumentException("Connectors glue to top-level shapes on the same slide.", parameterName);
            }
        }

        private PowerPointConnectionShape AddConnectorShape() {
            var connection = new ConnectionShape(
                new NonVisualConnectionShapeProperties(
                    new NonVisualDrawingProperties { Id = AllocateShapeId(), Name = GenerateUniqueName("Connector") },
                    new NonVisualConnectorShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()),
                new ShapeProperties(
                    new A.Transform2D(new A.Offset { X = 0L, Y = 0L }, new A.Extents { Cx = 0L, Cy = 0L }),
                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.StraightConnector1 },
                    new A.Outline(new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.Text1 })) { Width = 12700 }));
            CommonSlideData data = SlideRoot.CommonSlideData ??= new CommonSlideData(new ShapeTree());
            ShapeTree tree = data.ShapeTree ??= new ShapeTree();
            tree.Append(connection);
            return TrackShape(new PowerPointConnectionShape(connection, _slidePart));
        }
    }
}
