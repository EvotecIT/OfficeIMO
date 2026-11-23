namespace OfficeIMO.Excel {
    /// <summary>
    /// Typed columns for <see cref="ExcelSheet.WriteData{T}(IEnumerable{T}, Action{ExcelDataColumns{T}})"/>.
    /// Text, <c>int</c>, <c>long</c>, <c>double</c>, date and Boolean values are kept as their own types until save, without a boxed value per cell.
    /// </summary>
    /// <typeparam name="T">Type of the rows being written.</typeparam>
    /// <remarks>
    /// Columns write the same cells as the equivalent <c>(header, row =&gt; value)</c> selectors: null stays blank,
    /// an empty string stays text. Text, <c>int</c>, <c>long</c>, <c>double</c>, <see cref="DateTime"/> and <c>bool</c> columns
    /// (and their nullable forms) are stored unboxed; <c>decimal</c>, <c>float</c>, <c>ulong</c>, <c>char</c> and other values go
    /// through <see cref="Add(string, Func{T, object})"/>.
    /// </remarks>
    public sealed class ExcelDataColumns<T> {
        private readonly List<Column> _columns = new();

        internal ExcelDataColumns() {
        }

        internal int Count => _columns.Count;

        internal Column this[int index] => _columns[index];

        /// <summary>Adds a text column.</summary>
        public ExcelDataColumns<T> Add(string header, Func<T, string?> selector) => Add(new TextColumn(header, selector ?? throw new ArgumentNullException(nameof(selector))));

        /// <summary>Adds a 32-bit integer column.</summary>
        public ExcelDataColumns<T> Add(string header, Func<T, int> selector) => Add(new Int32Column(header, selector ?? throw new ArgumentNullException(nameof(selector))));

        /// <summary>Adds a 32-bit integer column whose null values stay blank.</summary>
        public ExcelDataColumns<T> Add(string header, Func<T, int?> selector) => Add(new NullableInt32Column(header, selector ?? throw new ArgumentNullException(nameof(selector))));

        /// <summary>Adds a 64-bit integer column; values beyond Excel's exact range are written as text.</summary>
        public ExcelDataColumns<T> Add(string header, Func<T, long> selector) => Add(new Int64Column(header, selector ?? throw new ArgumentNullException(nameof(selector))));

        /// <summary>Adds a 64-bit integer column whose null values stay blank.</summary>
        public ExcelDataColumns<T> Add(string header, Func<T, long?> selector) => Add(new NullableInt64Column(header, selector ?? throw new ArgumentNullException(nameof(selector))));

        /// <summary>Adds a double-precision number column.</summary>
        public ExcelDataColumns<T> Add(string header, Func<T, double> selector) => Add(new NumberColumn(header, selector ?? throw new ArgumentNullException(nameof(selector))));

        /// <summary>Adds a double-precision number column whose null values stay blank.</summary>
        public ExcelDataColumns<T> Add(string header, Func<T, double?> selector) => Add(new NullableNumberColumn(header, selector ?? throw new ArgumentNullException(nameof(selector))));

        /// <summary>Adds a date and time column, written with the tabular date format.</summary>
        public ExcelDataColumns<T> Add(string header, Func<T, DateTime> selector) => Add(new DateTimeColumn(header, selector ?? throw new ArgumentNullException(nameof(selector))));

        /// <summary>Adds a date and time column whose null values stay blank.</summary>
        public ExcelDataColumns<T> Add(string header, Func<T, DateTime?> selector) => Add(new NullableDateTimeColumn(header, selector ?? throw new ArgumentNullException(nameof(selector))));

        /// <summary>Adds a Boolean column.</summary>
        public ExcelDataColumns<T> Add(string header, Func<T, bool> selector) => Add(new BooleanColumn(header, selector ?? throw new ArgumentNullException(nameof(selector))));

        /// <summary>Adds a Boolean column whose null values stay blank.</summary>
        public ExcelDataColumns<T> Add(string header, Func<T, bool?> selector) => Add(new NullableBooleanColumn(header, selector ?? throw new ArgumentNullException(nameof(selector))));

        /// <summary>Adds a decimal column, written as the equivalent object selector writes it.</summary>
        public ExcelDataColumns<T> Add(string header, Func<T, decimal> selector) {
            ArgumentNullException.ThrowIfNull(selector);
            return Add(new ObjectColumn(header, row => selector(row)));
        }

        /// <summary>Adds a single-precision number column, written as the equivalent object selector writes it.</summary>
        public ExcelDataColumns<T> Add(string header, Func<T, float> selector) {
            ArgumentNullException.ThrowIfNull(selector);
            return Add(new ObjectColumn(header, row => selector(row)));
        }

        /// <summary>Adds an unsigned 64-bit integer column, written as the equivalent object selector writes it.</summary>
        public ExcelDataColumns<T> Add(string header, Func<T, ulong> selector) {
            ArgumentNullException.ThrowIfNull(selector);
            return Add(new ObjectColumn(header, row => selector(row)));
        }

        /// <summary>Adds a character column, written as the equivalent object selector writes it.</summary>
        public ExcelDataColumns<T> Add(string header, Func<T, char> selector) {
            ArgumentNullException.ThrowIfNull(selector);
            return Add(new ObjectColumn(header, row => selector(row)));
        }

        /// <summary>Adds a nullable decimal column, written as the equivalent object selector writes it.</summary>
        public ExcelDataColumns<T> Add(string header, Func<T, decimal?> selector) {
            ArgumentNullException.ThrowIfNull(selector);
            return Add(new ObjectColumn(header, row => selector(row)));
        }

        /// <summary>Adds a nullable single-precision number column, written as the equivalent object selector writes it.</summary>
        public ExcelDataColumns<T> Add(string header, Func<T, float?> selector) {
            ArgumentNullException.ThrowIfNull(selector);
            return Add(new ObjectColumn(header, row => selector(row)));
        }

        /// <summary>Adds a nullable unsigned 64-bit integer column, written as the equivalent object selector writes it.</summary>
        public ExcelDataColumns<T> Add(string header, Func<T, ulong?> selector) {
            ArgumentNullException.ThrowIfNull(selector);
            return Add(new ObjectColumn(header, row => selector(row)));
        }

        /// <summary>Adds a nullable character column, written as the equivalent object selector writes it.</summary>
        public ExcelDataColumns<T> Add(string header, Func<T, char?> selector) {
            ArgumentNullException.ThrowIfNull(selector);
            return Add(new ObjectColumn(header, row => selector(row)));
        }

        /// <summary>Adds a column of any value, written as a <c>(header, selector)</c> column of <c>WriteData</c> writes it.</summary>
        public ExcelDataColumns<T> Add(string header, Func<T, object?> selector) => Add(new ObjectColumn(header, selector ?? throw new ArgumentNullException(nameof(selector))));

        private ExcelDataColumns<T> Add(Column column) {
            _columns.Add(column);
            return this;
        }

        /// <summary>One column: its header and how a row's value is stored.</summary>
        internal abstract class Column {
            protected Column(string header) {
                Header = header;
            }

            internal string Header { get; }

            /// <summary>Stores the row's value at <paramref name="index"/>.</summary>
            /// <returns>The runtime type of the stored value, or null when the cell stays blank.</returns>
            internal abstract Type? Store(T row, ExcelTabularCellValues values, int index);
        }

        private sealed class TextColumn(string header, Func<T, string?> selector) : Column(header) {
            internal override Type? Store(T row, ExcelTabularCellValues values, int index) {
                string? value = selector(row);
                values.SetText(index, value);
                return value == null ? null : typeof(string);
            }
        }

        private sealed class Int32Column(string header, Func<T, int> selector) : Column(header) {
            internal override Type? Store(T row, ExcelTabularCellValues values, int index) {
                values.SetInt32(index, selector(row));
                return typeof(int);
            }
        }

        private sealed class NullableInt32Column(string header, Func<T, int?> selector) : Column(header) {
            internal override Type? Store(T row, ExcelTabularCellValues values, int index) {
                int? value = selector(row);
                if (!value.HasValue) {
                    values.SetBlank(index);
                    return null;
                }
                values.SetInt32(index, value.Value);
                return typeof(int);
            }
        }

        private sealed class Int64Column(string header, Func<T, long> selector) : Column(header) {
            internal override Type? Store(T row, ExcelTabularCellValues values, int index) {
                values.SetInt64(index, selector(row));
                return typeof(long);
            }
        }

        private sealed class NullableInt64Column(string header, Func<T, long?> selector) : Column(header) {
            internal override Type? Store(T row, ExcelTabularCellValues values, int index) {
                long? value = selector(row);
                if (!value.HasValue) {
                    values.SetBlank(index);
                    return null;
                }
                values.SetInt64(index, value.Value);
                return typeof(long);
            }
        }

        private sealed class NumberColumn(string header, Func<T, double> selector) : Column(header) {
            internal override Type? Store(T row, ExcelTabularCellValues values, int index) {
                values.SetNumber(index, selector(row));
                return typeof(double);
            }
        }

        private sealed class NullableNumberColumn(string header, Func<T, double?> selector) : Column(header) {
            internal override Type? Store(T row, ExcelTabularCellValues values, int index) {
                double? value = selector(row);
                if (!value.HasValue) {
                    values.SetBlank(index);
                    return null;
                }
                values.SetNumber(index, value.Value);
                return typeof(double);
            }
        }

        private sealed class DateTimeColumn(string header, Func<T, DateTime> selector) : Column(header) {
            internal override Type? Store(T row, ExcelTabularCellValues values, int index) {
                values.SetDateTime(index, selector(row));
                return typeof(DateTime);
            }
        }

        private sealed class NullableDateTimeColumn(string header, Func<T, DateTime?> selector) : Column(header) {
            internal override Type? Store(T row, ExcelTabularCellValues values, int index) {
                DateTime? value = selector(row);
                if (!value.HasValue) {
                    values.SetBlank(index);
                    return null;
                }
                values.SetDateTime(index, value.Value);
                return typeof(DateTime);
            }
        }

        private sealed class BooleanColumn(string header, Func<T, bool> selector) : Column(header) {
            internal override Type? Store(T row, ExcelTabularCellValues values, int index) {
                values.SetBoolean(index, selector(row));
                return typeof(bool);
            }
        }

        private sealed class NullableBooleanColumn(string header, Func<T, bool?> selector) : Column(header) {
            internal override Type? Store(T row, ExcelTabularCellValues values, int index) {
                bool? value = selector(row);
                if (!value.HasValue) {
                    values.SetBlank(index);
                    return null;
                }
                values.SetBoolean(index, value.Value);
                return typeof(bool);
            }
        }

        private sealed class ObjectColumn(string header, Func<T, object?> selector) : Column(header) {
            internal override Type? Store(T row, ExcelTabularCellValues values, int index) {
                object? value = selector(row);
                values.Set(index, value);
                return value == null || value == DBNull.Value ? null : value.GetType();
            }
        }
    }
}
