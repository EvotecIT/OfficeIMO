using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Visio {
    public static partial class VisioPageQueryExtensions {
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
    }
}
