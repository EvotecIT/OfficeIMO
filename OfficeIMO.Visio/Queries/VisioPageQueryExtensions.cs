using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Query and selection helpers for editing Visio pages by semantics instead of by list indexes.
    /// </summary>
    public static partial class VisioPageQueryExtensions {
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

            return FilterShapes(page, shape => string.Equals(GetDataValue(shape, key), value, comparison));
        }

        /// <summary>
        /// Returns shapes whose Shape Data value for the provided key matches a predicate.
        /// </summary>
        /// <param name="page">Page to query.</param>
        /// <param name="key">Data key.</param>
        /// <param name="predicate">Predicate that receives the current Shape Data value.</param>
        public static IReadOnlyList<VisioShape> ShapesWithData(this VisioPage page, string key, Func<string?, bool> predicate) {
            if (string.IsNullOrWhiteSpace(key)) {
                throw new ArgumentException("Data key cannot be empty.", nameof(key));
            }

            if (predicate == null) {
                throw new ArgumentNullException(nameof(predicate));
            }

            return FilterShapes(page, shape => predicate(GetDataValue(shape, key)));
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
        /// Returns shapes whose Visio Shape Data value for the provided row name matches a predicate.
        /// </summary>
        /// <param name="page">Page to query.</param>
        /// <param name="name">Shape Data row name.</param>
        /// <param name="predicate">Predicate that receives the current Shape Data value.</param>
        public static IReadOnlyList<VisioShape> ShapesWithShapeData(this VisioPage page, string name, Func<string?, bool> predicate) {
            return page.ShapesWithData(name, predicate);
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
        /// Returns shapes whose bounds intersect the provided bounds.
        /// </summary>
        /// <param name="page">Page to query.</param>
        /// <param name="bounds">Bounds to test against.</param>
        public static IReadOnlyList<VisioShape> ShapesIntersecting(this VisioPage page, VisioShapeBounds bounds) {
            if (bounds.IsEmpty) {
                return Array.Empty<VisioShape>();
            }

            return FilterShapes(page, shape => Intersects(GetPageShapeBounds(shape), bounds));
        }

        /// <summary>
        /// Returns shapes whose bounds intersect the provided shape.
        /// </summary>
        /// <param name="page">Page to query.</param>
        /// <param name="shape">Shape whose bounds are used for the test.</param>
        /// <param name="includeSelf">Whether the reference shape itself should be included.</param>
        public static IReadOnlyList<VisioShape> ShapesIntersecting(this VisioPage page, VisioShape shape, bool includeSelf = false) {
            EnsureShapeBelongsToPage(page, shape);
            VisioShapeBounds bounds = GetPageShapeBounds(shape);
            return FilterShapes(page, candidate => (includeSelf || !ReferenceEquals(candidate, shape)) && Intersects(GetPageShapeBounds(candidate), bounds));
        }

        /// <summary>
        /// Returns shapes whose bounds are fully contained by the provided bounds.
        /// </summary>
        /// <param name="page">Page to query.</param>
        /// <param name="bounds">Containing bounds.</param>
        public static IReadOnlyList<VisioShape> ShapesContainedIn(this VisioPage page, VisioShapeBounds bounds) {
            if (bounds.IsEmpty) {
                return Array.Empty<VisioShape>();
            }

            return FilterShapes(page, shape => Contains(bounds, GetPageShapeBounds(shape)));
        }

        /// <summary>
        /// Returns shapes whose bounds are fully contained by the provided shape.
        /// </summary>
        /// <param name="page">Page to query.</param>
        /// <param name="container">Shape whose bounds are used as the containing area.</param>
        /// <param name="includeContainer">Whether the containing shape itself should be included.</param>
        public static IReadOnlyList<VisioShape> ShapesContainedIn(this VisioPage page, VisioShape container, bool includeContainer = false) {
            EnsureShapeBelongsToPage(page, container);
            VisioShapeBounds bounds = GetPageShapeBounds(container);
            return FilterShapes(page, shape => (includeContainer || !ReferenceEquals(shape, container)) && Contains(bounds, GetPageShapeBounds(shape)));
        }

        /// <summary>
        /// Returns every shape reachable from the provided shape through connectors.
        /// </summary>
        /// <param name="page">Page to query.</param>
        /// <param name="shape">Starting shape.</param>
        /// <param name="includeStart">Whether the starting shape should be included in the returned component.</param>
        public static IReadOnlyList<VisioShape> ConnectedComponent(this VisioPage page, VisioShape shape, bool includeStart = true) {
            EnsureShapeBelongsToPage(page, shape);
            List<VisioShape> component = new();
            Queue<VisioShape> queue = new();
            HashSet<VisioShape> seen = new();
            queue.Enqueue(shape);
            seen.Add(shape);

            while (queue.Count > 0) {
                VisioShape current = queue.Dequeue();
                if (includeStart || !ReferenceEquals(current, shape)) {
                    component.Add(current);
                }

                foreach (VisioShape connected in page.ConnectedShapes(current)) {
                    if (seen.Add(connected)) {
                        queue.Enqueue(connected);
                    }
                }
            }

            return component;
        }

        /// <summary>
        /// Returns the shortest shape path between two connected shapes, or an empty list when no path exists.
        /// </summary>
        /// <param name="page">Page to query.</param>
        /// <param name="from">Starting shape.</param>
        /// <param name="to">Target shape.</param>
        /// <param name="includeEndpoints">Whether the starting and target shapes should be included.</param>
        public static IReadOnlyList<VisioShape> PathBetween(this VisioPage page, VisioShape from, VisioShape to, bool includeEndpoints = true) {
            EnsureShapeBelongsToPage(page, from);
            EnsureShapeBelongsToPage(page, to);
            if (ReferenceEquals(from, to)) {
                return includeEndpoints ? new[] { from } : Array.Empty<VisioShape>();
            }

            Queue<VisioShape> queue = new();
            Dictionary<VisioShape, VisioShape?> previous = new();
            queue.Enqueue(from);
            previous[from] = null;

            while (queue.Count > 0) {
                VisioShape current = queue.Dequeue();
                foreach (VisioShape connected in page.ConnectedShapes(current)) {
                    if (previous.ContainsKey(connected)) {
                        continue;
                    }

                    previous[connected] = current;
                    if (ReferenceEquals(connected, to)) {
                        return BuildPath(previous, from, to, includeEndpoints);
                    }

                    queue.Enqueue(connected);
                }
            }

            return Array.Empty<VisioShape>();
        }
    }
}
