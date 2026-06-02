using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace OfficeIMO.Visio {
/// <summary>
    /// Creates deterministic inspection snapshots for generated or loaded Visio documents.
    /// </summary>
    public static class VisioInspectionExtensions {
        /// <summary>
        /// Creates a stable, data-oriented snapshot of the document structure.
        /// </summary>
        public static VisioInspectionSnapshot CreateInspectionSnapshot(this VisioDocument document) {
            if (document == null) {
                throw new ArgumentNullException(nameof(document));
            }

            IReadOnlyList<VisioInspectionMasterSnapshot> masters = document.Masters
                .OrderBy(master => master.NameU, StringComparer.OrdinalIgnoreCase)
                .ThenBy(master => master.Id, StringComparer.OrdinalIgnoreCase)
                .Select(master => new VisioInspectionMasterSnapshot(
                    master.Id,
                    master.NameU,
                    master.Shape.NameU,
                    master.Shape.Text,
                    master.Shape.Width,
                    master.Shape.Height,
                    master.IsPackageBacked,
                    master.StencilId,
                    master.StencilName,
                    master.StencilCategory,
                    master.StencilCatalogName,
                    master.StencilSourcePackagePath,
                    master.StencilKeywords,
                    master.StencilAliases,
                    master.StencilTags,
                    master.StencilIconNameU,
                    master.StencilDefaultWidth,
                    master.StencilDefaultHeight,
                    master.StencilDefaultUnit?.ToString(),
                    master.StencilPreviewImageRelationshipId,
                    master.StencilPreviewImageTarget,
                    master.StencilPreviewImageContentType,
                    master.StencilPreviewImageExtension,
                    master.StencilPreviewImageByteLength))
                .ToList()
                .AsReadOnly();

            IReadOnlyList<VisioInspectionPageSnapshot> pages = document.Pages
                .OrderBy(page => page.Id)
                .ThenBy(page => page.Name, StringComparer.OrdinalIgnoreCase)
                .Select(CreatePageSnapshot)
                .ToList()
                .AsReadOnly();

            return new VisioInspectionSnapshot(
                document.Title,
                document.Author,
                document.Theme != null
                    ? string.IsNullOrWhiteSpace(document.Theme.Name) ? document.Theme.GetType().Name : document.Theme.Name
                    : null,
                document.UseMastersByDefault,
                document.WriteMasterDeltasOnly,
                masters,
                pages);
        }

        private static VisioInspectionPageSnapshot CreatePageSnapshot(VisioPage page) {
            IReadOnlyList<VisioInspectionShapeSnapshot> shapes = page.AllShapes()
                .OrderBy(shape => shape.Id, StringComparer.OrdinalIgnoreCase)
                .Select(CreateShapeSnapshot)
                .ToList()
                .AsReadOnly();

            IReadOnlyList<VisioInspectionConnectorSnapshot> connectors = page.Connectors
                .OrderBy(connector => connector.Id, StringComparer.OrdinalIgnoreCase)
                .Select(CreateConnectorSnapshot)
                .ToList()
                .AsReadOnly();

            IReadOnlyList<string> layers = page.Layers
                .Select(layer => layer.Name)
                .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
                .ToList()
                .AsReadOnly();

            return new VisioInspectionPageSnapshot(
                page.Id,
                page.Name,
                page.NameU,
                page.Width,
                page.Height,
                layers,
                shapes,
                connectors);
        }

        private static VisioInspectionShapeSnapshot CreateShapeSnapshot(VisioShape shape) {
            return new VisioInspectionShapeSnapshot(
                shape.Id,
                shape.Name,
                shape.NameU,
                shape.Type,
                shape.Master?.Id,
                shape.MasterNameU,
                shape.MasterShapeId,
                shape.Parent?.Id,
                shape.Text,
                shape.PinX,
                shape.PinY,
                shape.Width,
                shape.Height,
                shape.Angle,
                shape.LineColor.ToString(),
                shape.FillColor.ToString(),
                shape.LinePattern,
                shape.FillPattern,
                shape.LineWeight,
                shape.IsContainer,
                shape.IsCallout,
                shape.IsBackgroundSurface,
                shape.IsDiagramAdornment,
                shape.CalloutTargetId,
                SortStrings(shape.LayerNames),
                CreateShapeDataSnapshot(shape.ShapeData),
                CreateUserCellSnapshot(shape.UserCells),
                CreateDataSnapshot(shape.Data),
                CreateConnectionPointSnapshot(shape.ConnectionPoints),
                shape.Children.Select(child => child.Id).OrderBy(id => id, StringComparer.OrdinalIgnoreCase).ToList().AsReadOnly());
        }

        private static VisioInspectionConnectorSnapshot CreateConnectorSnapshot(VisioConnector connector) {
            VisioConnectorLabelPlacement? placement = connector.LabelPlacement;
            ResolveConnectorLabelPin(connector, placement, out double? labelResolvedPinX, out double? labelResolvedPinY);

            return new VisioInspectionConnectorSnapshot(
                connector.Id,
                connector.From.Id,
                connector.To.Id,
                connector.Kind.ToString(),
                connector.Label,
                placement != null,
                placement?.Position,
                placement?.OffsetX,
                placement?.OffsetY,
                placement?.PinX,
                placement?.PinY,
                labelResolvedPinX,
                labelResolvedPinY,
                placement?.GetLocPinX(),
                placement?.GetLocPinY(),
                placement?.Width,
                placement?.Height,
                connector.Waypoints
                    .Select(waypoint => new VisioInspectionWaypointSnapshot(waypoint.X, waypoint.Y))
                    .ToList()
                    .AsReadOnly(),
                connector.LineColor.ToString(),
                connector.LinePattern,
                connector.LineWeight,
                connector.BeginArrow?.ToString(),
                connector.EndArrow?.ToString(),
                SortStrings(connector.LayerNames),
                CreateShapeDataSnapshot(connector.ShapeData),
                CreateDataSnapshot(connector.Data));
        }

        private static void ResolveConnectorLabelPin(VisioConnector connector, VisioConnectorLabelPlacement? placement, out double? pinX, out double? pinY) {
            pinX = null;
            pinY = null;
            if (placement == null) {
                return;
            }

            if (placement.PinX.HasValue && placement.PinY.HasValue) {
                pinX = placement.PinX.Value;
                pinY = placement.PinY.Value;
                return;
            }

            List<ConnectorPathPoint> path = BuildConnectorPath(connector);
            if (path.Count == 0) {
                return;
            }

            ConnectorPathPoint point = ResolvePathPoint(path, placement.Position);
            pinX = point.X + placement.OffsetX;
            pinY = point.Y + placement.OffsetY;
        }

        private static List<ConnectorPathPoint> BuildConnectorPath(VisioConnector connector) {
            ResolveEndpoint(connector.From, connector.To, connector.FromConnectionPoint, out double startX, out double startY);
            ResolveEndpoint(connector.To, connector.From, connector.ToConnectionPoint, out double endX, out double endY);

            List<ConnectorPathPoint> points = new() {
                new ConnectorPathPoint(startX, startY)
            };

            if (connector.Waypoints.Count > 0) {
                foreach (VisioConnectorWaypoint waypoint in connector.Waypoints) {
                    points.Add(new ConnectorPathPoint(waypoint.X, waypoint.Y));
                }
            } else if (connector.Kind == ConnectorKind.RightAngle) {
                points.Add(new ConnectorPathPoint(startX, endY));
            }

            points.Add(new ConnectorPathPoint(endX, endY));
            return points;
        }

        private static void ResolveEndpoint(VisioShape shape, VisioShape other, VisioConnectionPoint? connectionPoint, out double x, out double y) {
            if (connectionPoint != null) {
                (x, y) = GetPagePoint(shape, connectionPoint.X, connectionPoint.Y);
                return;
            }

            (double left, double bottom, double right, double top) = GetPageBounds(shape);
            (double otherLeft, double otherBottom, double otherRight, double otherTop) = GetPageBounds(other);
            double centerX = (left + right) / 2D;
            double centerY = (bottom + top) / 2D;
            double otherCenterX = (otherLeft + otherRight) / 2D;
            double otherCenterY = (otherBottom + otherTop) / 2D;
            double dx = otherCenterX - centerX;
            double dy = otherCenterY - centerY;

            if (Math.Abs(dx) >= Math.Abs(dy)) {
                x = dx >= 0 ? right : left;
                y = centerY;
            } else {
                x = centerX;
                y = dy >= 0 ? top : bottom;
            }
        }

        private static (double Left, double Bottom, double Right, double Top) GetPageBounds(VisioShape shape) {
            (double x1, double y1) = GetPagePoint(shape, 0, 0);
            (double x2, double y2) = GetPagePoint(shape, shape.Width, 0);
            (double x3, double y3) = GetPagePoint(shape, 0, shape.Height);
            (double x4, double y4) = GetPagePoint(shape, shape.Width, shape.Height);
            double left = Math.Min(Math.Min(x1, x2), Math.Min(x3, x4));
            double right = Math.Max(Math.Max(x1, x2), Math.Max(x3, x4));
            double bottom = Math.Min(Math.Min(y1, y2), Math.Min(y3, y4));
            double top = Math.Max(Math.Max(y1, y2), Math.Max(y3, y4));
            return (left, bottom, right, top);
        }

        private static (double X, double Y) GetPagePoint(VisioShape shape, double x, double y) {
            (double absX, double absY) = shape.GetAbsolutePoint(x, y);
            return shape.Parent != null
                ? GetPagePoint(shape.Parent, absX, absY)
                : (absX, absY);
        }

        private static ConnectorPathPoint ResolvePathPoint(IReadOnlyList<ConnectorPathPoint> points, double position) {
            double clampedPosition = VisioConnectorLabelPlacement.ClampPosition(position);
            double totalLength = 0D;
            for (int i = 1; i < points.Count; i++) {
                totalLength += Distance(points[i - 1], points[i]);
            }

            if (totalLength <= 0D) {
                return points[0];
            }

            double targetLength = totalLength * clampedPosition;
            double traversed = 0D;
            for (int i = 1; i < points.Count; i++) {
                ConnectorPathPoint from = points[i - 1];
                ConnectorPathPoint to = points[i];
                double segmentLength = Distance(from, to);
                if (segmentLength <= 0D) {
                    continue;
                }

                if (traversed + segmentLength >= targetLength) {
                    double segmentPosition = (targetLength - traversed) / segmentLength;
                    return new ConnectorPathPoint(
                        from.X + ((to.X - from.X) * segmentPosition),
                        from.Y + ((to.Y - from.Y) * segmentPosition));
                }

                traversed += segmentLength;
            }

            return points[points.Count - 1];
        }

        private static double Distance(ConnectorPathPoint from, ConnectorPathPoint to) {
            double dx = to.X - from.X;
            double dy = to.Y - from.Y;
            return Math.Sqrt((dx * dx) + (dy * dy));
        }

        private readonly struct ConnectorPathPoint {
            public ConnectorPathPoint(double x, double y) {
                X = x;
                Y = y;
            }

            public double X { get; }

            public double Y { get; }
        }

        private static IReadOnlyList<VisioInspectionShapeDataSnapshot> CreateShapeDataSnapshot(IEnumerable<VisioShapeDataRow> rows) {
            return rows
                .OrderBy(row => row.Name, StringComparer.OrdinalIgnoreCase)
                .Select(row => new VisioInspectionShapeDataSnapshot(row.Name, row.Label, row.Value, row.Type?.ToString(), row.Format, row.Prompt))
                .ToList()
                .AsReadOnly();
        }

        private static IReadOnlyList<VisioInspectionUserCellSnapshot> CreateUserCellSnapshot(IEnumerable<VisioUserCell> rows) {
            return rows
                .OrderBy(row => row.Name, StringComparer.OrdinalIgnoreCase)
                .Select(row => new VisioInspectionUserCellSnapshot(row.Name, row.Value, row.Formula, row.Prompt))
                .ToList()
                .AsReadOnly();
        }

        private static IReadOnlyList<VisioInspectionConnectionPointSnapshot> CreateConnectionPointSnapshot(IEnumerable<VisioConnectionPoint> points) {
            return points
                .Select((point, index) => new VisioInspectionConnectionPointSnapshot(index, point.SectionIndex, point.X, point.Y, point.DirX, point.DirY))
                .ToList()
                .AsReadOnly();
        }

        private static IReadOnlyDictionary<string, string> CreateDataSnapshot(IDictionary<string, string> data) {
            return new ReadOnlyDictionary<string, string>(
                data.OrderBy(pair => pair.Key, StringComparer.OrdinalIgnoreCase)
                    .ToDictionary(pair => pair.Key, pair => pair.Value, StringComparer.Ordinal));
        }

        private static IReadOnlyList<string> SortStrings(IEnumerable<string> values) {
            return values
                .OrderBy(value => value, StringComparer.OrdinalIgnoreCase)
                .ToList()
                .AsReadOnly();
        }
    }
}
