using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Visio {
    public static partial class VisioPageQueryExtensions {
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
        /// Selects shapes whose Shape Data value for the provided key matches a predicate.
        /// </summary>
        /// <param name="page">Page to query.</param>
        /// <param name="key">Data key.</param>
        /// <param name="predicate">Predicate that receives the current Shape Data value.</param>
        public static VisioShapeSelection SelectWithData(this VisioPage page, string key, Func<string?, bool> predicate) {
            return new VisioShapeSelection(page.ShapesWithData(key, predicate), page);
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
        /// Selects shapes whose Visio Shape Data value for the provided row name matches a predicate.
        /// </summary>
        /// <param name="page">Page to query.</param>
        /// <param name="name">Shape Data row name.</param>
        /// <param name="predicate">Predicate that receives the current Shape Data value.</param>
        public static VisioShapeSelection SelectWithShapeData(this VisioPage page, string name, Func<string?, bool> predicate) {
            return new VisioShapeSelection(page.ShapesWithShapeData(name, predicate), page);
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
        /// Selects shapes whose bounds intersect the provided bounds.
        /// </summary>
        /// <param name="page">Page to query.</param>
        /// <param name="bounds">Bounds to test against.</param>
        public static VisioShapeSelection SelectIntersecting(this VisioPage page, VisioShapeBounds bounds) {
            return new VisioShapeSelection(page.ShapesIntersecting(bounds), page);
        }

        /// <summary>
        /// Selects shapes whose bounds intersect the provided shape.
        /// </summary>
        /// <param name="page">Page to query.</param>
        /// <param name="shape">Shape whose bounds are used for the test.</param>
        /// <param name="includeSelf">Whether the reference shape itself should be included.</param>
        public static VisioShapeSelection SelectIntersecting(this VisioPage page, VisioShape shape, bool includeSelf = false) {
            return new VisioShapeSelection(page.ShapesIntersecting(shape, includeSelf), page);
        }

        /// <summary>
        /// Selects shapes whose bounds are fully contained by the provided bounds.
        /// </summary>
        /// <param name="page">Page to query.</param>
        /// <param name="bounds">Containing bounds.</param>
        public static VisioShapeSelection SelectContainedIn(this VisioPage page, VisioShapeBounds bounds) {
            return new VisioShapeSelection(page.ShapesContainedIn(bounds), page);
        }

        /// <summary>
        /// Selects shapes whose bounds are fully contained by the provided shape.
        /// </summary>
        /// <param name="page">Page to query.</param>
        /// <param name="container">Shape whose bounds are used as the containing area.</param>
        /// <param name="includeContainer">Whether the containing shape itself should be included.</param>
        public static VisioShapeSelection SelectContainedIn(this VisioPage page, VisioShape container, bool includeContainer = false) {
            return new VisioShapeSelection(page.ShapesContainedIn(container, includeContainer), page);
        }

        /// <summary>
        /// Selects every shape reachable from the provided shape through connectors.
        /// </summary>
        /// <param name="page">Page to query.</param>
        /// <param name="shape">Starting shape.</param>
        /// <param name="includeStart">Whether the starting shape should be included in the returned component.</param>
        public static VisioShapeSelection SelectConnectedComponent(this VisioPage page, VisioShape shape, bool includeStart = true) {
            return new VisioShapeSelection(page.ConnectedComponent(shape, includeStart), page);
        }

        /// <summary>
        /// Selects the shortest shape path between two connected shapes.
        /// </summary>
        /// <param name="page">Page to query.</param>
        /// <param name="from">Starting shape.</param>
        /// <param name="to">Target shape.</param>
        /// <param name="includeEndpoints">Whether the starting and target shapes should be included.</param>
        public static VisioShapeSelection SelectPathBetween(this VisioPage page, VisioShape from, VisioShape to, bool includeEndpoints = true) {
            return new VisioShapeSelection(page.PathBetween(from, to, includeEndpoints), page);
        }
    }
}
