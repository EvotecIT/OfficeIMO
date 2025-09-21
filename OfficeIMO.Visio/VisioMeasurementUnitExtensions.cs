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
    }
}

