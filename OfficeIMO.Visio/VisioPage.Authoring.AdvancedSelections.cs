namespace OfficeIMO.Visio {
    public partial class VisioPage {
        /// <summary>
        /// Selects shapes with a matching shape name and configures them.
        /// </summary>
        public VisioPage ShapesByName(string name, Action<VisioShapeSelection> configure, StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
            ConfigureShapes(this.SelectByName(name, comparison), configure);
            return this;
        }

        /// <summary>
        /// Selects shapes with a matching universal shape name and configures them.
        /// </summary>
        public VisioPage ShapesByNameU(string nameU, Action<VisioShapeSelection> configure, StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
            ConfigureShapes(this.SelectByNameU(nameU, comparison), configure);
            return this;
        }

        /// <summary>
        /// Selects shapes created from a matching master universal name and configures them.
        /// </summary>
        public VisioPage ShapesByMaster(string masterNameU, Action<VisioShapeSelection> configure, StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
            ConfigureShapes(this.SelectByMaster(masterNameU, comparison), configure);
            return this;
        }

        /// <summary>
        /// Selects Visio-native container shapes and configures them.
        /// </summary>
        public VisioPage ConfigureContainers(Action<VisioShapeSelection> configure) {
            ConfigureShapes(this.SelectContainers(), configure);
            return this;
        }

        /// <summary>
        /// Selects OfficeIMO callout or annotation shapes and configures them.
        /// </summary>
        public VisioPage Callouts(Action<VisioShapeSelection> configure) {
            ConfigureShapes(this.SelectCallouts(), configure);
            return this;
        }

        /// <summary>
        /// Selects shapes that contain a Visio User cell and configures them.
        /// </summary>
        public VisioPage ShapesWithUserCell(string name, Action<VisioShapeSelection> configure) {
            ConfigureShapes(this.SelectWithUserCell(name), configure);
            return this;
        }

        /// <summary>
        /// Selects shapes with a matching Visio User cell value and configures them.
        /// </summary>
        public VisioPage ShapesWithUserCell(string name, string value, Action<VisioShapeSelection> configure, StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
            ConfigureShapes(this.SelectWithUserCell(name, value, comparison), configure);
            return this;
        }

        /// <summary>
        /// Selects shapes with at least one hyperlink and configures them.
        /// </summary>
        public VisioPage ShapesWithHyperlinks(Action<VisioShapeSelection> configure) {
            ConfigureShapes(this.SelectWithHyperlinks(), configure);
            return this;
        }

        /// <summary>
        /// Selects shapes with a matching hyperlink address and configures them.
        /// </summary>
        public VisioPage ShapesWithHyperlink(string address, Action<VisioShapeSelection> configure, StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
            ConfigureShapes(this.SelectWithHyperlink(address, comparison), configure);
            return this;
        }

        /// <summary>
        /// Selects shapes with any explicit protection cell and configures them.
        /// </summary>
        public VisioPage ShapesWithProtection(Action<VisioShapeSelection> configure) {
            ConfigureShapes(this.SelectWithProtection(), configure);
            return this;
        }

        /// <summary>
        /// Selects shapes whose protection state matches a predicate and configures them.
        /// </summary>
        public VisioPage ShapesWithProtection(Func<VisioShapeProtection, bool> predicate, Action<VisioShapeSelection> configure) {
            ConfigureShapes(this.SelectWithProtection(predicate), configure);
            return this;
        }

        /// <summary>
        /// Selects shapes whose bounds intersect the provided page-coordinate bounds and configures them.
        /// </summary>
        public VisioPage ShapesIntersecting(VisioShapeBounds bounds, Action<VisioShapeSelection> configure) {
            ConfigureShapes(this.SelectIntersecting(bounds), configure);
            return this;
        }

        /// <summary>
        /// Selects shapes whose bounds intersect a reference shape and configures them.
        /// </summary>
        public VisioPage ShapesIntersecting(string shapeId, Action<VisioShapeSelection> configure, bool includeSelf = false) {
            ConfigureShapes(this.SelectIntersecting(GetShape(shapeId), includeSelf), configure);
            return this;
        }

        /// <summary>
        /// Selects shapes fully contained by the provided page-coordinate bounds and configures them.
        /// </summary>
        public VisioPage ShapesContainedIn(VisioShapeBounds bounds, Action<VisioShapeSelection> configure) {
            ConfigureShapes(this.SelectContainedIn(bounds), configure);
            return this;
        }

        /// <summary>
        /// Selects shapes fully contained by a reference shape and configures them.
        /// </summary>
        public VisioPage ShapesContainedIn(string containerId, Action<VisioShapeSelection> configure, bool includeContainer = false) {
            ConfigureShapes(this.SelectContainedIn(GetShape(containerId), includeContainer), configure);
            return this;
        }

        /// <summary>
        /// Selects every shape reachable from the provided shape through connectors and configures them.
        /// </summary>
        public VisioPage ConnectedComponent(string shapeId, Action<VisioShapeSelection> configure, bool includeStart = true) {
            ConfigureShapes(this.SelectConnectedComponent(GetShape(shapeId), includeStart), configure);
            return this;
        }

        /// <summary>
        /// Selects the shortest shape path between two connected shapes and configures it.
        /// </summary>
        public VisioPage PathBetween(string fromId, string toId, Action<VisioShapeSelection> configure, bool includeEndpoints = true) {
            ConfigureShapes(this.SelectPathBetween(GetShape(fromId), GetShape(toId), includeEndpoints), configure);
            return this;
        }

        /// <summary>
        /// Selects connectors that start at the provided shape and configures them.
        /// </summary>
        public VisioPage OutgoingConnectors(string shapeId, Action<VisioConnectorSelection> configure) {
            ConfigureConnectors(this.SelectOutgoingConnectors(GetShape(shapeId)), configure);
            return this;
        }

        /// <summary>
        /// Selects connectors that end at the provided shape and configures them.
        /// </summary>
        public VisioPage IncomingConnectors(string shapeId, Action<VisioConnectorSelection> configure) {
            ConfigureConnectors(this.SelectIncomingConnectors(GetShape(shapeId)), configure);
            return this;
        }

        /// <summary>
        /// Selects connectors attached to the provided shape and configures them.
        /// </summary>
        public VisioPage ConnectedConnectors(string shapeId, Action<VisioConnectorSelection> configure) {
            ConfigureConnectors(this.SelectConnectedConnectors(GetShape(shapeId)), configure);
            return this;
        }

        /// <summary>
        /// Selects connectors assigned to a page layer and configures them.
        /// </summary>
        public VisioPage ConnectorsInLayer(string layerName, Action<VisioConnectorSelection> configure, StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
            ConfigureConnectors(this.SelectConnectorsInLayer(layerName, comparison), configure);
            return this;
        }

        /// <summary>
        /// Selects connectors with at least one hyperlink and configures them.
        /// </summary>
        public VisioPage ConnectorsWithHyperlinks(Action<VisioConnectorSelection> configure) {
            ConfigureConnectors(this.SelectConnectorsWithHyperlinks(), configure);
            return this;
        }

        /// <summary>
        /// Selects connectors with a matching hyperlink address and configures them.
        /// </summary>
        public VisioPage ConnectorsWithHyperlink(string address, Action<VisioConnectorSelection> configure, StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
            ConfigureConnectors(this.SelectConnectorsWithHyperlink(address, comparison), configure);
            return this;
        }

        /// <summary>
        /// Selects connectors with any explicit protection cell and configures them.
        /// </summary>
        public VisioPage ConnectorsWithProtection(Action<VisioConnectorSelection> configure) {
            ConfigureConnectors(this.SelectConnectorsWithProtection(), configure);
            return this;
        }

        /// <summary>
        /// Selects connectors whose protection state matches a predicate and configures them.
        /// </summary>
        public VisioPage ConnectorsWithProtection(Func<VisioProtection, bool> predicate, Action<VisioConnectorSelection> configure) {
            ConfigureConnectors(this.SelectConnectorsWithProtection(predicate), configure);
            return this;
        }

    }
}
