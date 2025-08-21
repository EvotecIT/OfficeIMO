using System;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Supported measurement units for Visio pages.
    /// </summary>
    public enum VisioMeasurementUnit {
        /// <summary>
        /// Measurements expressed in inches.
        /// </summary>
        Inches,

        /// <summary>
        /// Measurements expressed in centimeters.
        /// </summary>
        Centimeters
    }

    internal static class VisioMeasurementUnitExtensions {
        public static double ToInches(this double value, VisioMeasurementUnit unit) {
            return unit switch {
                VisioMeasurementUnit.Centimeters => value / 2.54,
                _ => value
            };
        }

        public static double FromInches(this double value, VisioMeasurementUnit unit) {
            return unit switch {
                VisioMeasurementUnit.Centimeters => value * 2.54,
                _ => value
            };
        }
    }
}
