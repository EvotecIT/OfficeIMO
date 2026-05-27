using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Query and selection helpers for editing Visio pages by semantics instead of by list indexes.
    /// </summary>
    public static class VisioPageQueryExtensions {
        /// <summary>
        /// Returns all shapes on the page, including shapes nested inside groups.
        /// </summary>
        /// <param name="page">Page to query.</param>
        public static IReadOnlyList<VisioShape> AllShapes(this VisioPage page) {
            if (page == null) {
                throw new ArgumentNullException(nameof(page));
            }

            List<VisioShape> shapes = new();
            foreach (VisioShape shape in page.Shapes) {
                AddShapeAndChildren(shape, shapes);
            }

            return shapes;
        }

        /// <summary>
        /// Finds a shape by identifier, including shapes nested inside groups.
        /// </summary>
        /// <param name="page">Page to query.</param>
        /// <param name="id">Shape identifier.</param>
        public static VisioShape? FindShapeById(this VisioPage page, string id) {
            if (page == null) {
                throw new ArgumentNullException(nameof(page));
            }

            if (string.IsNullOrWhiteSpace(id)) {
                throw new ArgumentException("Shape id cannot be empty.", nameof(id));
            }

            foreach (VisioShape shape in page.Shapes) {
                VisioShape? result = shape.FindDescendantById(id);
                if (result != null) {
                    return result;
                }
            }

            return null;
        }

        /// <summary>
        /// Returns shapes with a matching shape name.
        /// </summary>
        /// <param name="page">Page to query.</param>
        /// <param name="name">Shape name.</param>
        /// <param name="comparison">String comparison used for matching.</param>
        public static IReadOnlyList<VisioShape> ShapesByName(this VisioPage page, string name, StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
            return FilterShapes(page, shape => string.Equals(shape.Name, name, comparison));
        }

        /// <summary>
        /// Returns shapes with a matching universal shape name.
        /// </summary>
        /// <param name="page">Page to query.</param>
        /// <param name="nameU">Universal shape name.</param>
        /// <param name="comparison">String comparison used for matching.</param>
        public static IReadOnlyList<VisioShape> ShapesByNameU(this VisioPage page, string nameU, StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
            return FilterShapes(page, shape => string.Equals(shape.NameU, nameU, comparison));
        }

        /// <summary>
        /// Returns shapes created from a matching master universal name.
        /// </summary>
        /// <param name="page">Page to query.</param>
        /// <param name="masterNameU">Master universal name.</param>
        /// <param name="comparison">String comparison used for matching.</param>
        public static IReadOnlyList<VisioShape> ShapesByMaster(this VisioPage page, string masterNameU, StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
            return FilterShapes(page, shape => string.Equals(shape.MasterNameU, masterNameU, comparison));
        }

        /// <summary>
        /// Returns shapes whose text contains the provided value.
        /// </summary>
        /// <param name="page">Page to query.</param>
        /// <param name="text">Text fragment to find.</param>
        /// <param name="comparison">String comparison used for matching.</param>
        public static IReadOnlyList<VisioShape> ShapesContainingText(this VisioPage page, string text, StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
            if (text == null) {
                throw new ArgumentNullException(nameof(text));
            }

            return FilterShapes(page, shape => shape.Text != null && shape.Text.IndexOf(text, comparison) >= 0);
        }

        /// <summary>
        /// Returns shapes that contain the provided data key.
        /// </summary>
        /// <param name="page">Page to query.</param>
        /// <param name="key">Data key.</param>
        public static IReadOnlyList<VisioShape> ShapesWithData(this VisioPage page, string key) {
            if (string.IsNullOrWhiteSpace(key)) {
                throw new ArgumentException("Data key cannot be empty.", nameof(key));
            }

            return FilterShapes(page, shape => shape.FindShapeData(key) != null || shape.Data.ContainsKey(key));
        }

        /// <summary>
        /// Returns shapes that contain the provided data key and value.
        /// </summary>
        /// <param name="page">Page to query.</param>
        /// <param name="key">Data key.</param>
        /// <param name="value">Data value.</param>
        /// <param name="comparison">String comparison used for matching.</param>
        public static IReadOnlyList<VisioShape> ShapesWithData(this VisioPage page, string key, string value, StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
            if (string.IsNullOrWhiteSpace(key)) {
                throw new ArgumentException("Data key cannot be empty.", nameof(key));
            }

            if (value == null) {
                throw new ArgumentNullException(nameof(value));
            }

            return FilterShapes(page, shape => string.Equals(shape.GetShapeDataValue(key), value, comparison));
        }

        /// <summary>
        /// Returns shapes that contain a Visio Shape Data row.
        /// </summary>
        /// <param name="page">Page to query.</param>
        /// <param name="name">Shape Data row name.</param>
        public static IReadOnlyList<VisioShape> ShapesWithShapeData(this VisioPage page, string name) {
            return page.ShapesWithData(name);
        }

        /// <summary>
        /// Returns shapes with a matching Visio Shape Data value.
        /// </summary>
        /// <param name="page">Page to query.</param>
        /// <param name="name">Shape Data row name.</param>
        /// <param name="value">Shape Data value.</param>
        /// <param name="comparison">String comparison used for matching.</param>
        public static IReadOnlyList<VisioShape> ShapesWithShapeData(this VisioPage page, string name, string value, StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
            return page.ShapesWithData(name, value, comparison);
        }

        /// <summary>
        /// Returns shapes assigned to a page layer.
        /// </summary>
        /// <param name="page">Page to query.</param>
        /// <param name="layerName">Layer name or universal name.</param>
        /// <param name="comparison">String comparison used for matching.</param>
        public static IReadOnlyList<VisioShape> ShapesInLayer(this VisioPage page, string layerName, StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
            if (string.IsNullOrWhiteSpace(layerName)) {
                throw new ArgumentException("Layer name cannot be empty.", nameof(layerName));
            }

            return FilterShapes(page, shape => shape.LayerNames.Any(current => string.Equals(current, layerName, comparison)));
        }

        /// <summary>
        /// Returns shapes marked as Visio-native containers.
        /// </summary>
        /// <param name="page">Page to query.</param>
        public static IReadOnlyList<VisioShape> Containers(this VisioPage page) {
            return FilterShapes(page, shape => shape.IsContainer);
        }

        /// <summary>
        /// Returns shapes marked as OfficeIMO callouts or annotations.
        /// </summary>
        /// <param name="page">Page to query.</param>
        public static IReadOnlyList<VisioShape> Callouts(this VisioPage page) {
            return FilterShapes(page, shape => shape.IsCallout);
        }

        /// <summary>
        /// Returns shapes that contain a Visio User cell.
        /// </summary>
        /// <param name="page">Page to query.</param>
        /// <param name="name">User cell row name.</param>
        public static IReadOnlyList<VisioShape> ShapesWithUserCell(this VisioPage page, string name) {
            if (string.IsNullOrWhiteSpace(name)) {
                throw new ArgumentException("User cell name cannot be empty.", nameof(name));
            }

            return FilterShapes(page, shape => shape.FindUserCell(name) != null);
        }

        /// <summary>
        /// Returns shapes with a matching Visio User cell value.
        /// </summary>
        /// <param name="page">Page to query.</param>
        /// <param name="name">User cell row name.</param>
        /// <param name="value">User cell value.</param>
        /// <param name="comparison">String comparison used for matching.</param>
        public static IReadOnlyList<VisioShape> ShapesWithUserCell(this VisioPage page, string name, string value, StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
            if (string.IsNullOrWhiteSpace(name)) {
                throw new ArgumentException("User cell name cannot be empty.", nameof(name));
            }

            if (value == null) {
                throw new ArgumentNullException(nameof(value));
            }

            return FilterShapes(page, shape => string.Equals(shape.GetUserCellValue(name), value, comparison));
        }

        /// <summary>
        /// Returns shapes that contain at least one hyperlink.
        /// </summary>
        /// <param name="page">Page to query.</param>
        public static IReadOnlyList<VisioShape> ShapesWithHyperlinks(this VisioPage page) {
            return FilterShapes(page, shape => shape.Hyperlinks.Count > 0);
        }

        /// <summary>
        /// Returns shapes that contain a hyperlink with the provided address.
        /// </summary>
        /// <param name="page">Page to query.</param>
        /// <param name="address">Hyperlink address.</param>
        /// <param name="comparison">String comparison used for matching.</param>
        public static IReadOnlyList<VisioShape> ShapesWithHyperlink(this VisioPage page, string address, StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
            if (string.IsNullOrWhiteSpace(address)) {
                throw new ArgumentException("Hyperlink address cannot be empty.", nameof(address));
            }

            return FilterShapes(page, shape => shape.Hyperlinks.Any(hyperlink => string.Equals(hyperlink.Address, address, comparison)));
        }

        /// <summary>
        /// Returns shapes that have at least one explicit protection cell.
        /// </summary>
        /// <param name="page">Page to query.</param>
        public static IReadOnlyList<VisioShape> ShapesWithProtection(this VisioPage page) {
            return FilterShapes(page, shape => shape.Protection.HasAnyLocks);
        }

        /// <summary>
        /// Returns shapes whose protection state matches the predicate.
        /// </summary>
        /// <param name="page">Page to query.</param>
        /// <param name="predicate">Protection predicate.</param>
        public static IReadOnlyList<VisioShape> ShapesWithProtection(this VisioPage page, Func<VisioShapeProtection, bool> predicate) {
            if (predicate == null) {
                throw new ArgumentNullException(nameof(predicate));
            }

            return FilterShapes(page, shape => predicate(shape.Protection));
        }

        /// <summary>
        /// Selects shapes matching a predicate for bulk editing.
        /// </summary>
        /// <param name="page">Page to query.</param>
        /// <param name="predicate">Predicate used to include shapes.</param>
        public static VisioShapeSelection SelectShapes(this VisioPage page, Func<VisioShape, bool> predicate) {
            if (predicate == null) {
                throw new ArgumentNullException(nameof(predicate));
            }

            return new VisioShapeSelection(FilterShapes(page, predicate), page);
        }

        /// <summary>
        /// Selects shapes with a matching shape name.
        /// </summary>
        /// <param name="page">Page to query.</param>
        /// <param name="name">Shape name.</param>
        /// <param name="comparison">String comparison used for matching.</param>
        public static VisioShapeSelection SelectByName(this VisioPage page, string name, StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
            return new VisioShapeSelection(page.ShapesByName(name, comparison), page);
        }

        /// <summary>
        /// Selects shapes with a matching universal shape name.
        /// </summary>
        /// <param name="page">Page to query.</param>
        /// <param name="nameU">Universal shape name.</param>
        /// <param name="comparison">String comparison used for matching.</param>
        public static VisioShapeSelection SelectByNameU(this VisioPage page, string nameU, StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
            return new VisioShapeSelection(page.ShapesByNameU(nameU, comparison), page);
        }

        /// <summary>
        /// Selects shapes created from a matching master universal name.
        /// </summary>
        /// <param name="page">Page to query.</param>
        /// <param name="masterNameU">Master universal name.</param>
        /// <param name="comparison">String comparison used for matching.</param>
        public static VisioShapeSelection SelectByMaster(this VisioPage page, string masterNameU, StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
            return new VisioShapeSelection(page.ShapesByMaster(masterNameU, comparison), page);
        }

        /// <summary>
        /// Selects shapes whose text contains the provided value.
        /// </summary>
        /// <param name="page">Page to query.</param>
        /// <param name="text">Text fragment to find.</param>
        /// <param name="comparison">String comparison used for matching.</param>
        public static VisioShapeSelection SelectContainingText(this VisioPage page, string text, StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
            return new VisioShapeSelection(page.ShapesContainingText(text, comparison), page);
        }

        /// <summary>
        /// Selects shapes that contain the provided data key.
        /// </summary>
        /// <param name="page">Page to query.</param>
        /// <param name="key">Data key.</param>
        public static VisioShapeSelection SelectWithData(this VisioPage page, string key) {
            return new VisioShapeSelection(page.ShapesWithData(key), page);
        }

        /// <summary>
        /// Selects shapes that contain the provided data key and value.
        /// </summary>
        /// <param name="page">Page to query.</param>
        /// <param name="key">Data key.</param>
        /// <param name="value">Data value.</param>
        /// <param name="comparison">String comparison used for matching.</param>
        public static VisioShapeSelection SelectWithData(this VisioPage page, string key, string value, StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
            return new VisioShapeSelection(page.ShapesWithData(key, value, comparison), page);
        }

        /// <summary>
        /// Selects shapes that contain a Visio Shape Data row.
        /// </summary>
        /// <param name="page">Page to query.</param>
        /// <param name="name">Shape Data row name.</param>
        public static VisioShapeSelection SelectWithShapeData(this VisioPage page, string name) {
            return new VisioShapeSelection(page.ShapesWithShapeData(name), page);
        }

        /// <summary>
        /// Selects shapes with a matching Visio Shape Data value.
        /// </summary>
        /// <param name="page">Page to query.</param>
        /// <param name="name">Shape Data row name.</param>
        /// <param name="value">Shape Data value.</param>
        /// <param name="comparison">String comparison used for matching.</param>
        public static VisioShapeSelection SelectWithShapeData(this VisioPage page, string name, string value, StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
            return new VisioShapeSelection(page.ShapesWithShapeData(name, value, comparison), page);
        }

        /// <summary>
        /// Selects shapes assigned to a page layer.
        /// </summary>
        /// <param name="page">Page to query.</param>
        /// <param name="layerName">Layer name or universal name.</param>
        /// <param name="comparison">String comparison used for matching.</param>
        public static VisioShapeSelection SelectLayer(this VisioPage page, string layerName, StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
            return new VisioShapeSelection(page.ShapesInLayer(layerName, comparison), page);
        }

        /// <summary>
        /// Selects shapes marked as Visio-native containers.
        /// </summary>
        /// <param name="page">Page to query.</param>
        public static VisioShapeSelection SelectContainers(this VisioPage page) {
            return new VisioShapeSelection(page.Containers(), page);
        }

        /// <summary>
        /// Selects shapes marked as OfficeIMO callouts or annotations.
        /// </summary>
        /// <param name="page">Page to query.</param>
        public static VisioShapeSelection SelectCallouts(this VisioPage page) {
            return new VisioShapeSelection(page.Callouts(), page);
        }

        /// <summary>
        /// Selects shapes that contain a Visio User cell.
        /// </summary>
        /// <param name="page">Page to query.</param>
        /// <param name="name">User cell row name.</param>
        public static VisioShapeSelection SelectWithUserCell(this VisioPage page, string name) {
            return new VisioShapeSelection(page.ShapesWithUserCell(name), page);
        }

        /// <summary>
        /// Selects shapes with a matching Visio User cell value.
        /// </summary>
        /// <param name="page">Page to query.</param>
        /// <param name="name">User cell row name.</param>
        /// <param name="value">User cell value.</param>
        /// <param name="comparison">String comparison used for matching.</param>
        public static VisioShapeSelection SelectWithUserCell(this VisioPage page, string name, string value, StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
            return new VisioShapeSelection(page.ShapesWithUserCell(name, value, comparison), page);
        }

        /// <summary>
        /// Selects shapes that contain at least one hyperlink.
        /// </summary>
        /// <param name="page">Page to query.</param>
        public static VisioShapeSelection SelectWithHyperlinks(this VisioPage page) {
            return new VisioShapeSelection(page.ShapesWithHyperlinks(), page);
        }

        /// <summary>
        /// Selects shapes that contain a hyperlink with the provided address.
        /// </summary>
        /// <param name="page">Page to query.</param>
        /// <param name="address">Hyperlink address.</param>
        /// <param name="comparison">String comparison used for matching.</param>
        public static VisioShapeSelection SelectWithHyperlink(this VisioPage page, string address, StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
            return new VisioShapeSelection(page.ShapesWithHyperlink(address, comparison), page);
        }

        /// <summary>
        /// Selects shapes that have at least one explicit protection cell.
        /// </summary>
        /// <param name="page">Page to query.</param>
        public static VisioShapeSelection SelectWithProtection(this VisioPage page) {
            return new VisioShapeSelection(page.ShapesWithProtection(), page);
        }

        /// <summary>
        /// Selects shapes whose protection state matches the predicate.
        /// </summary>
        /// <param name="page">Page to query.</param>
        /// <param name="predicate">Protection predicate.</param>
        public static VisioShapeSelection SelectWithProtection(this VisioPage page, Func<VisioShapeProtection, bool> predicate) {
            return new VisioShapeSelection(page.ShapesWithProtection(predicate), page);
        }

        /// <summary>
        /// Returns connectors that start at the provided shape.
        /// </summary>
        /// <param name="page">Page to query.</param>
        /// <param name="shape">Shape connected from.</param>
        public static IReadOnlyList<VisioConnector> OutgoingConnectors(this VisioPage page, VisioShape shape) {
            EnsureShapeBelongsToPage(page, shape);
            return FilterConnectors(page, connector => MatchesShape(connector.From, shape));
        }

        /// <summary>
        /// Returns connectors that end at the provided shape.
        /// </summary>
        /// <param name="page">Page to query.</param>
        /// <param name="shape">Shape connected to.</param>
        public static IReadOnlyList<VisioConnector> IncomingConnectors(this VisioPage page, VisioShape shape) {
            EnsureShapeBelongsToPage(page, shape);
            return FilterConnectors(page, connector => MatchesShape(connector.To, shape));
        }

        /// <summary>
        /// Returns connectors that either start or end at the provided shape.
        /// </summary>
        /// <param name="page">Page to query.</param>
        /// <param name="shape">Shape connected to or from.</param>
        public static IReadOnlyList<VisioConnector> ConnectedConnectors(this VisioPage page, VisioShape shape) {
            EnsureShapeBelongsToPage(page, shape);
            return FilterConnectors(page, connector => MatchesShape(connector.From, shape) || MatchesShape(connector.To, shape));
        }

        /// <summary>
        /// Returns connectors assigned to a page layer.
        /// </summary>
        /// <param name="page">Page to query.</param>
        /// <param name="layerName">Layer name or universal name.</param>
        /// <param name="comparison">String comparison used for matching.</param>
        public static IReadOnlyList<VisioConnector> ConnectorsInLayer(this VisioPage page, string layerName, StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
            if (string.IsNullOrWhiteSpace(layerName)) {
                throw new ArgumentException("Layer name cannot be empty.", nameof(layerName));
            }

            return FilterConnectors(page, connector => connector.LayerNames.Any(current => string.Equals(current, layerName, comparison)));
        }

        /// <summary>
        /// Returns connectors that contain at least one hyperlink.
        /// </summary>
        /// <param name="page">Page to query.</param>
        public static IReadOnlyList<VisioConnector> ConnectorsWithHyperlinks(this VisioPage page) {
            return FilterConnectors(page, connector => connector.Hyperlinks.Count > 0);
        }

        /// <summary>
        /// Returns connectors that contain a hyperlink with the provided address.
        /// </summary>
        /// <param name="page">Page to query.</param>
        /// <param name="address">Hyperlink address.</param>
        /// <param name="comparison">String comparison used for matching.</param>
        public static IReadOnlyList<VisioConnector> ConnectorsWithHyperlink(this VisioPage page, string address, StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
            if (string.IsNullOrWhiteSpace(address)) {
                throw new ArgumentException("Hyperlink address cannot be empty.", nameof(address));
            }

            return FilterConnectors(page, connector => connector.Hyperlinks.Any(hyperlink => string.Equals(hyperlink.Address, address, comparison)));
        }

        /// <summary>
        /// Returns connectors that have at least one explicit protection cell.
        /// </summary>
        /// <param name="page">Page to query.</param>
        public static IReadOnlyList<VisioConnector> ConnectorsWithProtection(this VisioPage page) {
            return FilterConnectors(page, connector => connector.Protection.HasAnyLocks);
        }

        /// <summary>
        /// Returns connectors whose protection state matches the predicate.
        /// </summary>
        /// <param name="page">Page to query.</param>
        /// <param name="predicate">Protection predicate.</param>
        public static IReadOnlyList<VisioConnector> ConnectorsWithProtection(this VisioPage page, Func<VisioProtection, bool> predicate) {
            if (predicate == null) {
                throw new ArgumentNullException(nameof(predicate));
            }

            return FilterConnectors(page, connector => predicate(connector.Protection));
        }

        /// <summary>
        /// Returns shapes connected to the provided shape.
        /// </summary>
        /// <param name="page">Page to query.</param>
        /// <param name="shape">Shape whose neighbors should be returned.</param>
        public static IReadOnlyList<VisioShape> ConnectedShapes(this VisioPage page, VisioShape shape) {
            if (shape == null) {
                throw new ArgumentNullException(nameof(shape));
            }

            List<VisioShape> connectedShapes = new();
            foreach (VisioConnector connector in page.ConnectedConnectors(shape)) {
                VisioShape candidate = MatchesShape(connector.From, shape) ? connector.To : connector.From;
                if (!connectedShapes.Contains(candidate)) {
                    connectedShapes.Add(candidate);
                }
            }

            return connectedShapes;
        }

        /// <summary>
        /// Selects connectors matching a predicate for bulk editing.
        /// </summary>
        /// <param name="page">Page to query.</param>
        /// <param name="predicate">Predicate used to include connectors.</param>
        public static VisioConnectorSelection SelectConnectors(this VisioPage page, Func<VisioConnector, bool> predicate) {
            if (predicate == null) {
                throw new ArgumentNullException(nameof(predicate));
            }

            return new VisioConnectorSelection(FilterConnectors(page, predicate));
        }

        /// <summary>
        /// Selects connectors that start at the provided shape.
        /// </summary>
        /// <param name="page">Page to query.</param>
        /// <param name="shape">Shape connected from.</param>
        public static VisioConnectorSelection SelectOutgoingConnectors(this VisioPage page, VisioShape shape) {
            return new VisioConnectorSelection(page.OutgoingConnectors(shape));
        }

        /// <summary>
        /// Selects connectors that end at the provided shape.
        /// </summary>
        /// <param name="page">Page to query.</param>
        /// <param name="shape">Shape connected to.</param>
        public static VisioConnectorSelection SelectIncomingConnectors(this VisioPage page, VisioShape shape) {
            return new VisioConnectorSelection(page.IncomingConnectors(shape));
        }

        /// <summary>
        /// Selects connectors that either start or end at the provided shape.
        /// </summary>
        /// <param name="page">Page to query.</param>
        /// <param name="shape">Shape connected to or from.</param>
        public static VisioConnectorSelection SelectConnectedConnectors(this VisioPage page, VisioShape shape) {
            return new VisioConnectorSelection(page.ConnectedConnectors(shape));
        }

        /// <summary>
        /// Selects connectors assigned to a page layer.
        /// </summary>
        /// <param name="page">Page to query.</param>
        /// <param name="layerName">Layer name or universal name.</param>
        /// <param name="comparison">String comparison used for matching.</param>
        public static VisioConnectorSelection SelectConnectorsInLayer(this VisioPage page, string layerName, StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
            return new VisioConnectorSelection(page.ConnectorsInLayer(layerName, comparison));
        }

        /// <summary>
        /// Selects connectors that contain at least one hyperlink.
        /// </summary>
        /// <param name="page">Page to query.</param>
        public static VisioConnectorSelection SelectConnectorsWithHyperlinks(this VisioPage page) {
            return new VisioConnectorSelection(page.ConnectorsWithHyperlinks());
        }

        /// <summary>
        /// Selects connectors that contain a hyperlink with the provided address.
        /// </summary>
        /// <param name="page">Page to query.</param>
        /// <param name="address">Hyperlink address.</param>
        /// <param name="comparison">String comparison used for matching.</param>
        public static VisioConnectorSelection SelectConnectorsWithHyperlink(this VisioPage page, string address, StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
            return new VisioConnectorSelection(page.ConnectorsWithHyperlink(address, comparison));
        }

        /// <summary>
        /// Selects connectors that have at least one explicit protection cell.
        /// </summary>
        /// <param name="page">Page to query.</param>
        public static VisioConnectorSelection SelectConnectorsWithProtection(this VisioPage page) {
            return new VisioConnectorSelection(page.ConnectorsWithProtection());
        }

        /// <summary>
        /// Selects connectors whose protection state matches the predicate.
        /// </summary>
        /// <param name="page">Page to query.</param>
        /// <param name="predicate">Protection predicate.</param>
        public static VisioConnectorSelection SelectConnectorsWithProtection(this VisioPage page, Func<VisioProtection, bool> predicate) {
            return new VisioConnectorSelection(page.ConnectorsWithProtection(predicate));
        }

        private static IReadOnlyList<VisioShape> FilterShapes(VisioPage page, Func<VisioShape, bool> predicate) {
            if (page == null) {
                throw new ArgumentNullException(nameof(page));
            }

            return page.AllShapes().Where(predicate).ToList();
        }

        private static IReadOnlyList<VisioConnector> FilterConnectors(VisioPage page, Func<VisioConnector, bool> predicate) {
            if (page == null) {
                throw new ArgumentNullException(nameof(page));
            }

            return page.Connectors.Where(predicate).ToList();
        }

        private static void AddShapeAndChildren(VisioShape shape, List<VisioShape> shapes) {
            shapes.Add(shape);
            foreach (VisioShape child in shape.Children) {
                AddShapeAndChildren(child, shapes);
            }
        }

        private static bool MatchesShape(VisioShape candidate, VisioShape shape) {
            if (shape == null) {
                throw new ArgumentNullException(nameof(shape));
            }

            return ReferenceEquals(candidate, shape);
        }

        private static void EnsureShapeBelongsToPage(VisioPage page, VisioShape shape) {
            if (shape == null) {
                throw new ArgumentNullException(nameof(shape));
            }

            if (page == null) {
                throw new ArgumentNullException(nameof(page));
            }

            if (!page.AllShapes().Contains(shape)) {
                throw new InvalidOperationException("The shape must belong to the page.");
            }
        }
    }
}
