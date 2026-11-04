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
        ///     Explicit column width as a physical length.
        /// </summary>
        public OfficeLength? Width { get; private set; }

        /// <summary>
        ///     Creates a column binding.
        /// </summary>
        public static PowerPointTableColumn<T> Create(string header, Func<T, object?> valueSelector) {
            return new PowerPointTableColumn<T>(header, valueSelector);
        }
        /// <summary>Sets the physical column width.</summary>
        public PowerPointTableColumn<T> WithWidth(OfficeLength width) {
            Width = width;
            return this;
        }
    }
}