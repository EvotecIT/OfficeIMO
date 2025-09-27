using System;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Extension helpers for converting to and from inches.
    /// </summary>
    internal static class VisioMeasurementUnitExtensions {
        /// <summary>Converts a value in the specified unit to inches.</summary>
        public static double ToInches(this double value, VisioMeasurementUnit unit) {
            return unit switch {
                VisioMeasurementUnit.Centimeters => value / 2.54,
                VisioMeasurementUnit.Millimeters => value / 25.4,
                _ => value
            };
        }

        /// <summary>Converts a value in inches to the specified unit.</summary>
        public static double FromInches(this double value, VisioMeasurementUnit unit) {
            return unit switch {
                VisioMeasurementUnit.Centimeters => value * 2.54,
                VisioMeasurementUnit.Millimeters => value * 25.4,
                _ => value
            };
        }

        /// <summary>Returns the Visio unit code for the given measurement unit.</summary>
        internal static string ToVisioUnitCode(this VisioMeasurementUnit unit) {
            return unit switch {
                VisioMeasurementUnit.Centimeters => "CM",
                VisioMeasurementUnit.Millimeters => "MM",
                _ => "IN"
            };
        }

        /// <summary>Translates a Visio unit code (e.g. "MM") into a <see cref="VisioMeasurementUnit"/>.</summary>
        internal static VisioMeasurementUnit FromVisioUnitCode(string? unitCode, VisioMeasurementUnit fallback) {
            if (string.IsNullOrEmpty(unitCode)) {
                return fallback;
            }

            return unitCode.ToUpperInvariant() switch {
                "CM" => VisioMeasurementUnit.Centimeters,
                "MM" => VisioMeasurementUnit.Millimeters,
                "IN" => VisioMeasurementUnit.Inches,
                _ => fallback
            };
        }
    }
}

