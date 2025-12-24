using System;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Describes a column binding for table creation.
    /// </summary>
    public sealed class PowerPointTableColumn<T> {
        /// <summary>
        ///     Initializes a new column binding.
        /// </summary>
        public PowerPointTableColumn(string header, Func<T, object?> valueSelector) {
            Header = header ?? throw new ArgumentNullException(nameof(header));
            ValueSelector = valueSelector ?? throw new ArgumentNullException(nameof(valueSelector));
        }

        /// <summary>
        ///     Column header text.
        /// </summary>
        public string Header { get; }

        /// <summary>
        ///     Selector for cell values.
        /// </summary>
        public Func<T, object?> ValueSelector { get; }

        /// <summary>
        ///     Explicit column width in EMUs.
        /// </summary>
        public long? WidthEmus { get; private set; }

        /// <summary>
        ///     Creates a column binding.
        /// </summary>
        public static PowerPointTableColumn<T> Create(string header, Func<T, object?> valueSelector) {
            return new PowerPointTableColumn<T>(header, valueSelector);
        }

        /// <summary>
        ///     Sets the column width in EMUs.
        /// </summary>
        public PowerPointTableColumn<T> WithWidth(long widthEmus) {
            WidthEmus = widthEmus;
            return this;
        }

        /// <summary>
        ///     Sets the column width in centimeters.
        /// </summary>
        public PowerPointTableColumn<T> WithWidthCm(double widthCm) {
            return WithWidth(PowerPointUnits.FromCentimeters(widthCm));
        }

        /// <summary>
        ///     Sets the column width in inches.
        /// </summary>
        public PowerPointTableColumn<T> WithWidthInches(double widthInches) {
            return WithWidth(PowerPointUnits.FromInches(widthInches));
        }

        /// <summary>
        ///     Sets the column width in points.
        /// </summary>
        public PowerPointTableColumn<T> WithWidthPoints(double widthPoints) {
            return WithWidth(PowerPointUnits.FromPoints(widthPoints));
        }
    }
}
