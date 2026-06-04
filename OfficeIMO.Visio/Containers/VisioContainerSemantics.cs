using System;
using System.Globalization;

namespace OfficeIMO.Visio {
    internal static class VisioContainerSemantics {
        internal const string ShapeCategories = "msvShapeCategories";
        internal const string StructureType = "msvStructureType";
        internal const string Margin = "msvSDContainerMargin";
        internal const string Resize = "msvSDContainerResize";
        internal const string Locked = "msvSDContainerLocked";
        internal const string NoHighlight = "msvSDContainerNoHighlight";
        internal const string NoRibbon = "msvSDContainerNoRibbon";
        internal const string ContainerStyle = "msvSDContainerStyle";
        internal const string HeadingStyle = "msvSDHeadingStyle";

        internal static void Apply(VisioShape container, VisioContainerOptions options, VisioMeasurementUnit unit) {
            if (container == null) {
                throw new ArgumentNullException(nameof(container));
            }

            Validate(options);

            if (options.ShapeStyle != null) {
                container.ApplyStyle(options.ShapeStyle);
            } else {
                container.FillColor = options.FillColor;
                container.LineColor = options.LineColor;
                container.LineWeight = options.LineWeight;
                container.FillPattern = 1;
                container.LinePattern = 1;
            }

            if (options.TextStyle != null) {
                container.TextStyle = options.TextStyle.Clone();
            }

            container.SetUserCell(ShapeCategories, "ContainerStyleDefaults", "STR", prompt: string.Empty);
            container.SetUserCell(StructureType, "Container", "STR", prompt: string.Empty);
            container.SetUserCell(Margin, options.Margin.ToInches(unit).ToString(CultureInfo.InvariantCulture), "IN", prompt: string.Empty);
            container.SetUserCell(Resize, options.AutoResize ? "1" : "0", prompt: string.Empty);
            container.SetUserCell(Locked, options.Locked ? "1" : "0", "BOOL", prompt: string.Empty);
            container.SetUserCell(NoHighlight, options.NoHighlight ? "1" : "0", "BOOL", prompt: string.Empty);
            container.SetUserCell(NoRibbon, options.NoRibbon ? "1" : "0", "BOOL", prompt: string.Empty);
            container.SetUserCell(ContainerStyle, options.ContainerStyle.ToString(CultureInfo.InvariantCulture), prompt: string.Empty);
            container.SetUserCell(HeadingStyle, options.HeadingStyle.ToString(CultureInfo.InvariantCulture), prompt: string.Empty);
            container.SetUserCell(VisioSemanticUserCells.ContainerHeadingHeight, options.HeadingHeight.ToInches(unit).ToString(CultureInfo.InvariantCulture), "IN", prompt: "OfficeIMO container heading height");
        }

        internal static VisioContainerOptions CreateOptionsFrom(VisioShape container, VisioMeasurementUnit unit) {
            if (container == null) {
                throw new ArgumentNullException(nameof(container));
            }

            VisioContainerOptions options = new() {
                FillColor = container.FillColor,
                LineColor = container.LineColor,
                LineWeight = container.LineWeight,
                TextStyle = container.TextStyle?.Clone()
            };

            if (TryGetStoredInches(container, Margin, out double margin)) {
                options.Margin = margin.FromInches(unit);
            }

            if (TryGetStoredInches(container, VisioSemanticUserCells.ContainerHeadingHeight, out double headingHeight)) {
                options.HeadingHeight = headingHeight.FromInches(unit);
            }

            options.AutoResize = !IsFalse(container.GetUserCellValue(Resize));
            options.Locked = IsTrue(container.GetUserCellValue(Locked));
            options.NoHighlight = IsTrue(container.GetUserCellValue(NoHighlight));
            options.NoRibbon = IsTrue(container.GetUserCellValue(NoRibbon));
            options.ContainerStyle = ParsePositiveInt(container.GetUserCellValue(ContainerStyle), options.ContainerStyle);
            options.HeadingStyle = ParsePositiveInt(container.GetUserCellValue(HeadingStyle), options.HeadingStyle);
            return options;
        }

        internal static void Validate(VisioContainerOptions options) {
            if (options == null) {
                throw new ArgumentNullException(nameof(options));
            }

            if (!IsFiniteNonNegative(options.Margin)) {
                throw new ArgumentOutOfRangeException(nameof(options.Margin), "Container margin must be a finite non-negative number.");
            }

            if (!IsFiniteNonNegative(options.HeadingHeight)) {
                throw new ArgumentOutOfRangeException(nameof(options.HeadingHeight), "Container heading height must be a finite non-negative number.");
            }

            if (!IsFiniteNonNegative(options.LineWeight)) {
                throw new ArgumentOutOfRangeException(nameof(options.LineWeight), "Container line weight must be a finite non-negative number.");
            }

            if (options.ShapeStyle != null && !IsFiniteNonNegative(options.ShapeStyle.LineWeight)) {
                throw new ArgumentOutOfRangeException(nameof(options.ShapeStyle), "Container shape style line weight must be a finite non-negative number.");
            }

            if (options.ContainerStyle < 0) {
                throw new ArgumentOutOfRangeException(nameof(options.ContainerStyle), "Container style cannot be negative.");
            }

            if (options.HeadingStyle < 0) {
                throw new ArgumentOutOfRangeException(nameof(options.HeadingStyle), "Container heading style cannot be negative.");
            }
        }

        private static bool TryGetStoredInches(VisioShape container, string cellName, out double value) {
            string? raw = container.GetUserCellValue(cellName);
            return double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out value) && IsFiniteNonNegative(value);
        }

        private static int ParsePositiveInt(string? value, int fallback) {
            return int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed) && parsed >= 0
                ? parsed
                : fallback;
        }

        private static bool IsTrue(string? value) {
            return string.Equals(value, "1", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(value, "true", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsFalse(string? value) {
            return string.Equals(value, "0", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(value, "false", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsFiniteNonNegative(double value) {
            return !double.IsNaN(value) && !double.IsInfinity(value) && value >= 0D;
        }
    }
}
