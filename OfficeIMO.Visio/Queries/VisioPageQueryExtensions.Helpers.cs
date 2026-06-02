using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Visio {
    public static partial class VisioPageQueryExtensions {
        private static string? GetDataValue(VisioShape shape, string key) {
            string? shapeDataValue = shape.GetShapeDataValue(key);
            if (shapeDataValue != null) {
                return shapeDataValue;
            }

            return shape.Data.TryGetValue(key, out string? dataValue) ? dataValue : null;
        }

        private static VisioShapeBounds GetPageShapeBounds(VisioShape shape) {
            (double x1, double y1) = GetPagePoint(shape, 0, 0);
            (double x2, double y2) = GetPagePoint(shape, shape.Width, 0);
            (double x3, double y3) = GetPagePoint(shape, 0, shape.Height);
            (double x4, double y4) = GetPagePoint(shape, shape.Width, shape.Height);
            double left = Math.Min(Math.Min(x1, x2), Math.Min(x3, x4));
            double right = Math.Max(Math.Max(x1, x2), Math.Max(x3, x4));
            double bottom = Math.Min(Math.Min(y1, y2), Math.Min(y3, y4));
            double top = Math.Max(Math.Max(y1, y2), Math.Max(y3, y4));
            return new VisioShapeBounds(left, bottom, right, top);
        }

        private static (double X, double Y) GetPagePoint(VisioShape shape, double x, double y) {
            (double absX, double absY) = shape.GetAbsolutePoint(x, y);
            return shape.Parent != null
                ? GetPagePoint(shape.Parent, absX, absY)
                : (absX, absY);
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

        private static IReadOnlyList<VisioShape> BuildPath(IReadOnlyDictionary<VisioShape, VisioShape?> previous, VisioShape from, VisioShape to, bool includeEndpoints) {
            List<VisioShape> path = new();
            VisioShape? current = to;
            while (current != null) {
                path.Add(current);
                current = previous[current];
            }

            path.Reverse();
            if (!includeEndpoints && path.Count > 0) {
                path.RemoveAt(path.Count - 1);
                if (path.Count > 0) {
                    path.RemoveAt(0);
                }
            }

            return path;
        }

        private static bool Intersects(VisioShapeBounds first, VisioShapeBounds second) {
            if (first.IsEmpty || second.IsEmpty) {
                return false;
            }

            double width = Math.Min(first.Right, second.Right) - Math.Max(first.Left, second.Left);
            double height = Math.Min(first.Top, second.Top) - Math.Max(first.Bottom, second.Bottom);
            return width > 0D && height > 0D;
        }

        private static bool Contains(VisioShapeBounds outer, VisioShapeBounds inner) {
            if (outer.IsEmpty || inner.IsEmpty) {
                return false;
            }

            return inner.Left >= outer.Left &&
                   inner.Right <= outer.Right &&
                   inner.Bottom >= outer.Bottom &&
                   inner.Top <= outer.Top;
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
