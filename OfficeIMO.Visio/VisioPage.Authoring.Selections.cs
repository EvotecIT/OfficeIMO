using System;
using System.Linq;

namespace OfficeIMO.Visio {
    public partial class VisioPage {
        /// <summary>
        /// Configures all shapes on the page, including grouped children.
        /// </summary>
        /// <param name="configure">Configuration to apply to the selected shapes.</param>
        public VisioPage ConfigureShapes(Action<VisioShapeSelection> configure) {
            return ConfigureShapes(_ => true, configure);
        }

        /// <summary>
        /// Selects shapes with a strongly typed predicate and configures the stable selection.
        /// </summary>
        /// <param name="predicate">Predicate used to select shapes.</param>
        /// <param name="configure">Configuration to apply to the selected shapes.</param>
        public VisioPage ConfigureShapes(Func<VisioShape, bool> predicate, Action<VisioShapeSelection> configure) {
            ArgumentNullException.ThrowIfNull(predicate);

            ConfigureShapes(this.SelectShapes(predicate), configure);
            return this;
        }

        /// <summary>
        /// Selects shapes whose text contains the provided value and configures them.
        /// </summary>
        public VisioPage ShapesContainingText(string text, Action<VisioShapeSelection> configure, StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
            ConfigureShapes(this.SelectContainingText(text, comparison), configure);
            return this;
        }

        /// <summary>
        /// Selects shapes with a matching Shape Data value and configures them.
        /// </summary>
        public VisioPage ShapesWithData(string key, string value, Action<VisioShapeSelection> configure, StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
            ConfigureShapes(this.SelectWithData(key, value, comparison), configure);
            return this;
        }

        /// <summary>
        /// Selects shapes whose Shape Data value matches a predicate and configures them.
        /// </summary>
        public VisioPage ShapesWithData(string key, Func<string?, bool> predicate, Action<VisioShapeSelection> configure) {
            ConfigureShapes(this.SelectWithData(key, predicate), configure);
            return this;
        }

        /// <summary>
        /// Selects shapes with a matching typed Shape Data value and configures them.
        /// </summary>
        public VisioPage ShapesWithShapeData(string name, string value, Action<VisioShapeSelection> configure, StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
            ConfigureShapes(this.SelectWithShapeData(name, value, comparison), configure);
            return this;
        }

        /// <summary>
        /// Selects shapes in a page layer and configures them.
        /// </summary>
        public VisioPage ShapesInLayer(string layerName, Action<VisioShapeSelection> configure, StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
            ConfigureShapes(this.SelectLayer(layerName, comparison), configure);
            return this;
        }

        /// <summary>
        /// Configures all connectors on the page.
        /// </summary>
        /// <param name="configure">Configuration to apply to the selected connectors.</param>
        public VisioPage ConfigureConnectors(Action<VisioConnectorSelection> configure) {
            return ConfigureConnectors(_ => true, configure);
        }

        /// <summary>
        /// Selects connectors with a strongly typed predicate and configures the stable selection.
        /// </summary>
        /// <param name="predicate">Predicate used to select connectors.</param>
        /// <param name="configure">Configuration to apply to the selected connectors.</param>
        public VisioPage ConfigureConnectors(Func<VisioConnector, bool> predicate, Action<VisioConnectorSelection> configure) {
            ArgumentNullException.ThrowIfNull(predicate);

            ConfigureConnectors(new VisioConnectorSelection(this.Connectors.Where(predicate)), configure);
            return this;
        }

        private void ConfigureShapes(VisioShapeSelection selection, Action<VisioShapeSelection> configure) {
            ArgumentNullException.ThrowIfNull(configure);

            configure(selection);
            ValidateShapeIdentifiers();
        }

        private static void ConfigureConnectors(VisioConnectorSelection selection, Action<VisioConnectorSelection> configure) {
            ArgumentNullException.ThrowIfNull(configure);

            configure(selection);
        }
    }
}
