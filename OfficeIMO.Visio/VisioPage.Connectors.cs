using System;
using System.Collections.Generic;
using System.Xml.Linq;

namespace OfficeIMO.Visio {
    public partial class VisioPage {

        /// <summary>
        /// Moves a shape from its current location in the page hierarchy into the provided group shape.
        /// </summary>
        /// <param name="shape">The shape to move.</param>
        /// <param name="newParent">The group that should own the shape after the move.</param>
        /// <param name="childIndex">
        /// Optional insertion index within the target group's children.
        /// Use <c>-1</c> to append.
        /// </param>
        public void ReparentShape(VisioShape shape, VisioShape newParent, int childIndex = -1) {
            if (shape == null) {
                throw new ArgumentNullException(nameof(shape));
            }

            if (newParent == null) {
                throw new ArgumentNullException(nameof(newParent));
            }

            if (childIndex < -1) {
                throw new ArgumentOutOfRangeException(nameof(childIndex), "Child index must be -1 or greater.");
            }

            if (ReferenceEquals(shape, newParent)) {
                throw new InvalidOperationException("A shape cannot be reparented into itself.");
            }

            if (!TryFindShapeCollection(shape, out IList<VisioShape>? currentCollection, out int currentIndex)) {
                throw new InvalidOperationException("The shape is not part of this page.");
            }

            if (!TryFindShapeCollection(newParent, out _, out _)) {
                throw new InvalidOperationException("The target parent shape is not part of this page.");
            }

            if (ReferenceEquals(currentCollection, newParent.Children)) {
                if (childIndex < 0) {
                    childIndex = currentCollection.Count;
                }

                if (childIndex == currentIndex) {
                    return;
                }

                if (childIndex > currentCollection.Count) {
                    throw new ArgumentOutOfRangeException(nameof(childIndex), "Child index cannot exceed the number of children in the target group.");
                }

                currentCollection.RemoveAt(currentIndex);
                if (childIndex > currentCollection.Count) {
                    childIndex = currentCollection.Count;
                }

                currentCollection.Insert(childIndex, shape);
                return;
            }

            if (childIndex > newParent.Children.Count) {
                throw new ArgumentOutOfRangeException(nameof(childIndex), "Child index cannot exceed the number of children in the target group.");
            }

            IList<VisioShape> currentOwnerCollection = currentCollection!;
            currentOwnerCollection.RemoveAt(currentIndex);
            try {
                if (childIndex < 0) {
                    newParent.Children.Add(shape);
                } else {
                    newParent.Children.Insert(childIndex, shape);
                }
            } catch {
                currentOwnerCollection.Insert(currentIndex, shape);
                throw;
            }
        }

        /// <summary>
        /// Removes a group shape and promotes its children into the group's former position.
        /// </summary>
        /// <param name="group">The group to ungroup.</param>
        /// <returns>The children that were promoted.</returns>
        public IReadOnlyList<VisioShape> UngroupShape(VisioShape group) {
            if (group == null) {
                throw new ArgumentNullException(nameof(group));
            }

            if (!TryFindShapeCollection(group, out IList<VisioShape>? ownerCollection, out int index)) {
                throw new InvalidOperationException("The group shape is not part of this page.");
            }

            IList<VisioShape> resolvedOwnerCollection = ownerCollection!;

            if (group.Children.Count == 0) {
                resolvedOwnerCollection.RemoveAt(index);
                return Array.Empty<VisioShape>();
            }

            List<VisioShape> promotedChildren = new(group.Children);
            resolvedOwnerCollection.RemoveAt(index);
            group.Children.Clear();

            for (int i = 0; i < promotedChildren.Count; i++) {
                resolvedOwnerCollection.Insert(index + i, promotedChildren[i]);
            }

            return promotedChildren;
        }

        /// <summary>
        /// Reconnects the start of an existing connector to a different shape.
        /// </summary>
        public void ReconnectConnectorStart(VisioConnector connector, VisioShape newFrom, VisioSide side = VisioSide.Auto) {
            if (connector == null) {
                throw new ArgumentNullException(nameof(connector));
            }

            if (newFrom == null) {
                throw new ArgumentNullException(nameof(newFrom));
            }

            EnsureConnectorBelongsToPage(connector);
            EnsureShapeBelongsToPage(newFrom, "The source shape is not part of this page.");

            connector.From = newFrom;
            connector.FromConnectionPoint = ResolveConnectionPoint(newFrom, side);
            connector.PreservedFromConnectionCell = null;
            connector.PreservedBeginConnectAttributes.Clear();
            connector.PreservedBeginConnectAttributeOrder.Clear();
        }

        /// <summary>
        /// Reconnects the end of an existing connector to a different shape.
        /// </summary>
        public void ReconnectConnectorEnd(VisioConnector connector, VisioShape newTo, VisioSide side = VisioSide.Auto) {
            if (connector == null) {
                throw new ArgumentNullException(nameof(connector));
            }

            if (newTo == null) {
                throw new ArgumentNullException(nameof(newTo));
            }

            EnsureConnectorBelongsToPage(connector);
            EnsureShapeBelongsToPage(newTo, "The target shape is not part of this page.");

            connector.To = newTo;
            connector.ToConnectionPoint = ResolveConnectionPoint(newTo, side);
            connector.PreservedToConnectionCell = null;
            connector.PreservedEndConnectAttributes.Clear();
            connector.PreservedEndConnectAttributeOrder.Clear();
        }

        /// <summary>
        /// Reconnects both ends of an existing connector.
        /// </summary>
        public void ReconnectConnector(VisioConnector connector, VisioShape newFrom, VisioShape newTo, VisioSide fromSide = VisioSide.Auto, VisioSide toSide = VisioSide.Auto) {
            if (connector == null) {
                throw new ArgumentNullException(nameof(connector));
            }

            if (newFrom == null) {
                throw new ArgumentNullException(nameof(newFrom));
            }

            if (newTo == null) {
                throw new ArgumentNullException(nameof(newTo));
            }

            EnsureConnectorBelongsToPage(connector);
            EnsureShapeBelongsToPage(newFrom, "The source shape is not part of this page.");
            EnsureShapeBelongsToPage(newTo, "The target shape is not part of this page.");

            connector.From = newFrom;
            connector.To = newTo;
            connector.FromConnectionPoint = ResolveConnectionPoint(newFrom, fromSide);
            connector.ToConnectionPoint = ResolveConnectionPoint(newTo, toSide);
            connector.PreservedFromConnectionCell = null;
            connector.PreservedToConnectionCell = null;
            connector.PreservedBeginConnectAttributes.Clear();
            connector.PreservedEndConnectAttributes.Clear();
            connector.PreservedBeginConnectAttributeOrder.Clear();
            connector.PreservedEndConnectAttributeOrder.Clear();
        }

        /// <summary>
        /// Retargets all connector endpoints on this page that currently reference one shape to another.
        /// </summary>
        /// <param name="oldShape">The existing shape referenced by matching connectors.</param>
        /// <param name="newShape">The replacement shape that matching connectors should reference.</param>
        /// <param name="endpointScope">Controls whether start points, end points, or both are updated.</param>
        /// <param name="fromSide">The side to glue to when a start point is updated.</param>
        /// <param name="toSide">The side to glue to when an end point is updated.</param>
        /// <returns>The connectors that were updated.</returns>
        public IReadOnlyList<VisioConnector> RetargetConnectors(VisioShape oldShape, VisioShape newShape, VisioConnectorEndpointScope endpointScope = VisioConnectorEndpointScope.Both, VisioSide fromSide = VisioSide.Auto, VisioSide toSide = VisioSide.Auto) {
            if (oldShape == null) {
                throw new ArgumentNullException(nameof(oldShape));
            }

            if (newShape == null) {
                throw new ArgumentNullException(nameof(newShape));
            }

            EnsureShapeBelongsToPage(newShape, "The replacement shape is not part of this page.");

            if (ReferenceEquals(oldShape, newShape)) {
                return Array.Empty<VisioConnector>();
            }

            List<VisioConnector> updatedConnectors = new();
            for (int i = 0; i < _connectors.Count; i++) {
                VisioConnector connector = _connectors[i];
                bool updateStart = endpointScope != VisioConnectorEndpointScope.End && ReferenceEquals(connector.From, oldShape);
                bool updateEnd = endpointScope != VisioConnectorEndpointScope.Start && ReferenceEquals(connector.To, oldShape);

                if (!updateStart && !updateEnd) {
                    continue;
                }

                if (updateStart && updateEnd) {
                    ReconnectConnector(connector, newShape, newShape, fromSide, toSide);
                } else if (updateStart) {
                    ReconnectConnectorStart(connector, newShape, fromSide);
                } else {
                    ReconnectConnectorEnd(connector, newShape, toSide);
                }

                updatedConnectors.Add(connector);
            }

            if (updatedConnectors.Count == 0 && !TryFindShapeCollection(oldShape, out _, out _)) {
                throw new InvalidOperationException("The original shape is not part of this page or referenced by any connector on this page.");
            }

            return updatedConnectors;
        }

        private void EnsureConnectorBelongsToPage(VisioConnector connector) {
            if (!_connectors.Contains(connector)) {
                throw new InvalidOperationException("The connector is not part of this page.");
            }
        }

        private void EnsureShapeBelongsToPage(VisioShape shape, string message) {
            if (!TryFindShapeCollection(shape, out _, out _)) {
                throw new InvalidOperationException(message);
            }
        }

        private static VisioConnectionPoint? ResolveConnectionPoint(VisioShape shape, VisioSide side) {
            return side == VisioSide.Auto ? null : shape.EnsureSideConnectionPoint(side);
        }

        private bool TryFindShapeCollection(VisioShape target, out IList<VisioShape>? ownerCollection, out int index) {
            return TryFindShapeCollection(_shapeCollection, target, out ownerCollection, out index);
        }

        private static bool TryFindShapeCollection(IList<VisioShape> collection, VisioShape target, out IList<VisioShape>? ownerCollection, out int index) {
            for (int i = 0; i < collection.Count; i++) {
                VisioShape shape = collection[i];
                if (ReferenceEquals(shape, target)) {
                    ownerCollection = collection;
                    index = i;
                    return true;
                }

                if (TryFindShapeCollection(shape.Children, target, out ownerCollection, out index)) {
                    return true;
                }
            }

            ownerCollection = null;
            index = -1;
            return false;
        }

        /// <summary>
        /// Adds a connector between two shapes.
        /// </summary>
        /// <param name="id">Identifier of the connector.</param>
        /// <param name="from">Shape from which the connector starts.</param>
        /// <param name="to">Shape to which the connector ends.</param>
        /// <param name="kind">Type of connector.</param>
        /// <param name="fromSide">Preferred side on the source shape.</param>
        /// <param name="toSide">Preferred side on the target shape.</param>
        /// <returns>The created connector.</returns>
        public VisioConnector AddConnector(string id, VisioShape from, VisioShape to, ConnectorKind kind, VisioSide fromSide = VisioSide.Auto, VisioSide toSide = VisioSide.Auto) {
            return AddConnectorCore(id, from, to, kind, fromSide, toSide);
        }
    }
}
