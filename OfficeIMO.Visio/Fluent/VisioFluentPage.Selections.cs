using System;
using System.Linq;

namespace OfficeIMO.Visio.Fluent {
    public partial class VisioFluentPage {
        /// <summary>
        /// Configures all shapes on the page, including grouped children.
        /// </summary>
        /// <param name="configure">Configuration to apply to the selected shapes.</param>
        public VisioFluentPage Shapes(Action<VisioShapeSelection> configure) {
            return Shapes(_ => true, configure);
        }

        /// <summary>
        /// Selects shapes with a strongly typed predicate and configures the stable selection.
        /// </summary>
        /// <param name="predicate">Predicate used to select shapes.</param>
        /// <param name="configure">Configuration to apply to the selected shapes.</param>
        public VisioFluentPage Shapes(Func<VisioShape, bool> predicate, Action<VisioShapeSelection> configure) {
            if (predicate == null) {
                throw new ArgumentNullException(nameof(predicate));
            }

            ConfigureShapes(Page.SelectShapes(predicate), configure);
            return this;
        }

        /// <summary>
        /// Selects shapes whose text contains the provided value and configures them.
        /// </summary>
        public VisioFluentPage ShapesContainingText(string text, Action<VisioShapeSelection> configure, StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
            ConfigureShapes(Page.SelectContainingText(text, comparison), configure);
            return this;
        }

        /// <summary>
        /// Selects shapes with a matching Shape Data value and configures them.
        /// </summary>
        public VisioFluentPage ShapesWithData(string key, string value, Action<VisioShapeSelection> configure, StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
            ConfigureShapes(Page.SelectWithData(key, value, comparison), configure);
            return this;
        }

        /// <summary>
        /// Selects shapes whose Shape Data value matches a predicate and configures them.
        /// </summary>
        public VisioFluentPage ShapesWithData(string key, Func<string?, bool> predicate, Action<VisioShapeSelection> configure) {
            ConfigureShapes(Page.SelectWithData(key, predicate), configure);
            return this;
        }

        /// <summary>
        /// Selects shapes with a matching typed Shape Data value and configures them.
        /// </summary>
        public VisioFluentPage ShapesWithShapeData(string name, string value, Action<VisioShapeSelection> configure, StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
            ConfigureShapes(Page.SelectWithShapeData(name, value, comparison), configure);
            return this;
        }

        /// <summary>
        /// Selects shapes in a page layer and configures them.
        /// </summary>
        public VisioFluentPage ShapesInLayer(string layerName, Action<VisioShapeSelection> configure, StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
            ConfigureShapes(Page.SelectLayer(layerName, comparison), configure);
            return this;
        }

        /// <summary>
        /// Configures all connectors on the page.
        /// </summary>
        /// <param name="configure">Configuration to apply to the selected connectors.</param>
        public VisioFluentPage Connectors(Action<VisioConnectorSelection> configure) {
            return Connectors(_ => true, configure);
        }

        /// <summary>
        /// Selects connectors with a strongly typed predicate and configures the stable selection.
        /// </summary>
        /// <param name="predicate">Predicate used to select connectors.</param>
        /// <param name="configure">Configuration to apply to the selected connectors.</param>
        public VisioFluentPage Connectors(Func<VisioConnector, bool> predicate, Action<VisioConnectorSelection> configure) {
            if (predicate == null) {
                throw new ArgumentNullException(nameof(predicate));
            }

            ConfigureConnectors(new VisioConnectorSelection(Page.Connectors.Where(predicate)), configure);
            return this;
        }

        private void ConfigureShapes(VisioShapeSelection selection, Action<VisioShapeSelection> configure) {
            if (configure == null) {
                throw new ArgumentNullException(nameof(configure));
            }

            configure(selection);
            RebuildShapeIndex();
        }

        private static void ConfigureConnectors(VisioConnectorSelection selection, Action<VisioConnectorSelection> configure) {
            if (configure == null) {
                throw new ArgumentNullException(nameof(configure));
            }

            configure(selection);
        }
    }
}
