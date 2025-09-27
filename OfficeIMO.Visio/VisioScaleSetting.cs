using System;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Represents a scale value stored on a Visio page (PageScale/DrawingScale).
    /// </summary>
    public struct VisioScaleSetting {
        /// <summary>
        /// Initializes a new <see cref="VisioScaleSetting"/> with the supplied value and unit.
        /// </summary>
        /// <param name="value">Scale magnitude expressed in <paramref name="unit"/>.</param>
        /// <param name="unit">Measurement unit for the scale.</param>
        public VisioScaleSetting(double value, VisioMeasurementUnit unit) {
            Value = value;
            Unit = unit;
        }

        /// <summary>
        /// Gets or sets the scale magnitude expressed in <see cref="Unit"/>.
        /// </summary>
        public double Value { get; set; }

        /// <summary>
        /// Gets or sets the measurement unit associated with <see cref="Value"/>.
        /// </summary>
        public VisioMeasurementUnit Unit { get; set; }

        /// <summary>
        /// Returns a sanitized copy of the scale where invalid magnitudes are replaced with 1.
        /// </summary>
        internal VisioScaleSetting Normalized() {
            double sanitized = Value;
            if (double.IsNaN(sanitized) || double.IsInfinity(sanitized) || sanitized <= 0) {
                sanitized = 1;
            }
            return new VisioScaleSetting(sanitized, Unit);
        }

        /// <summary>
        /// Converts the scale value to inches for serialization.
        /// </summary>
        internal double ToInches() => Value.ToInches(Unit);

        /// <summary>
        /// Returns an equivalent scale expressed in a different measurement unit.
        /// </summary>
        /// <param name="targetUnit">Target measurement unit.</param>
        internal VisioScaleSetting ConvertTo(VisioMeasurementUnit targetUnit) {
            if (Unit == targetUnit) {
                return this;
            }

            double inches = Value.ToInches(Unit);
            double converted = inches.FromInches(targetUnit);
            return new VisioScaleSetting(converted, targetUnit).Normalized();
        }

        /// <summary>
        /// Creates a scale value from an inch-based magnitude captured in the XML.
        /// </summary>
        internal static VisioScaleSetting FromInches(double valueInInches, string? unitCode, VisioMeasurementUnit fallbackUnit) {
            VisioMeasurementUnit unit = VisioMeasurementUnitExtensions.FromVisioUnitCode(unitCode, fallbackUnit);
            double measurementValue = valueInInches.FromInches(unit);
            return new VisioScaleSetting(measurementValue, unit).Normalized();
        }

        /// <summary>
        /// Gets a value indicating whether the scale represents a 1:1 ratio in <see cref="Unit"/>.
        /// </summary>
        internal bool IsDefault => Math.Abs(Value - 1) < 1e-9;

        /// <summary>
        /// Creates a default scale (1:1) for the specified measurement unit.
        /// </summary>
        public static VisioScaleSetting FromUnit(VisioMeasurementUnit unit) => new(1, unit);
    }
}
