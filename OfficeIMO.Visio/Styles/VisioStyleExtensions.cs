using System;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Convenience helpers for applying reusable Visio styles.
    /// </summary>
    public static class VisioStyleExtensions {
        /// <summary>Applies a reusable style to a shape.</summary>
        public static VisioShape ApplyStyle(this VisioShape shape, VisioShapeStyle style) {
            if (style == null) {
                throw new ArgumentNullException(nameof(style));
            }

            style.ApplyTo(shape);
            return shape;
        }

        /// <summary>Applies a reusable style to a shape selection.</summary>
        public static VisioShapeSelection ApplyStyle(this VisioShapeSelection selection, VisioShapeStyle style) {
            if (selection == null) {
                throw new ArgumentNullException(nameof(selection));
            }

            if (style == null) {
                throw new ArgumentNullException(nameof(style));
            }

            foreach (VisioShape shape in selection) {
                style.ApplyTo(shape);
            }

            return selection;
        }

        /// <summary>Applies a reusable style to a connector.</summary>
        public static VisioConnector ApplyStyle(this VisioConnector connector, VisioConnectorStyle style) {
            if (style == null) {
                throw new ArgumentNullException(nameof(style));
            }

            style.ApplyTo(connector);
            return connector;
        }

        /// <summary>Applies a reusable style to a connector selection.</summary>
        public static VisioConnectorSelection ApplyStyle(this VisioConnectorSelection selection, VisioConnectorStyle style) {
            if (selection == null) {
                throw new ArgumentNullException(nameof(selection));
            }

            if (style == null) {
                throw new ArgumentNullException(nameof(style));
            }

            foreach (VisioConnector connector in selection) {
                style.ApplyTo(connector);
            }

            return selection;
        }

        /// <summary>Applies a reusable text style to a shape.</summary>
        public static VisioShape ApplyTextStyle(this VisioShape shape, VisioTextStyle style) {
            if (shape == null) {
                throw new ArgumentNullException(nameof(shape));
            }

            if (style == null) {
                throw new ArgumentNullException(nameof(style));
            }

            shape.TextStyle = style.Clone();
            return shape;
        }

        /// <summary>Applies a reusable text style to a shape selection.</summary>
        public static VisioShapeSelection ApplyTextStyle(this VisioShapeSelection selection, VisioTextStyle style) {
            if (selection == null) {
                throw new ArgumentNullException(nameof(selection));
            }

            if (style == null) {
                throw new ArgumentNullException(nameof(style));
            }

            foreach (VisioShape shape in selection) {
                shape.TextStyle = style.Clone();
            }

            return selection;
        }

        /// <summary>Applies a reusable text style to a connector label.</summary>
        public static VisioConnector ApplyTextStyle(this VisioConnector connector, VisioTextStyle style) {
            if (connector == null) {
                throw new ArgumentNullException(nameof(connector));
            }

            if (style == null) {
                throw new ArgumentNullException(nameof(style));
            }

            connector.TextStyle = style.Clone();
            return connector;
        }

        /// <summary>Applies a reusable text style to connector labels in a selection.</summary>
        public static VisioConnectorSelection ApplyTextStyle(this VisioConnectorSelection selection, VisioTextStyle style) {
            if (selection == null) {
                throw new ArgumentNullException(nameof(selection));
            }

            if (style == null) {
                throw new ArgumentNullException(nameof(style));
            }

            foreach (VisioConnector connector in selection) {
                connector.TextStyle = style.Clone();
            }

            return selection;
        }
    }
}
