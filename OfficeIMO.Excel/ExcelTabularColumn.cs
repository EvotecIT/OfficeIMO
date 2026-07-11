namespace OfficeIMO.Excel {
    /// <summary>
    /// Describes a strongly typed column for package-native object exports.
    /// </summary>
    public abstract class ExcelTabularColumn<T> {
        private protected ExcelTabularColumn(string header) {
            if (string.IsNullOrWhiteSpace(header)) throw new ArgumentException("Column header must not be empty.", nameof(header));
            Header = header;
        }

        /// <summary>Column header written to the worksheet.</summary>
        public string Header { get; }

        internal abstract Type DataType { get; }

        internal abstract object? GetValue(T row);

        internal abstract void WriteValue(ExcelDocument.ExcelTabularRowWriter writer, T row);

        /// <summary>Creates a text column.</summary>
        public static ExcelTabularColumn<T> Create(string header, Func<T, string?> selector) => new StringColumn(header, selector);

        /// <summary>Creates a Boolean column.</summary>
        public static ExcelTabularColumn<T> Create(string header, Func<T, bool> selector) => new BooleanColumn(header, selector);

        /// <summary>Creates a date and time column.</summary>
        public static ExcelTabularColumn<T> Create(string header, Func<T, DateTime> selector) => new DateTimeColumn(header, selector);

        /// <summary>Creates a date, time, and offset column.</summary>
        public static ExcelTabularColumn<T> Create(string header, Func<T, DateTimeOffset> selector) => new DateTimeOffsetColumn(header, selector);

        /// <summary>Creates a duration column.</summary>
        public static ExcelTabularColumn<T> Create(string header, Func<T, TimeSpan> selector) => new TimeSpanColumn(header, selector);

        /// <summary>Creates a double-precision numeric column.</summary>
        public static ExcelTabularColumn<T> Create(string header, Func<T, double> selector) => new DoubleColumn(header, selector);

        /// <summary>Creates a single-precision numeric column.</summary>
        public static ExcelTabularColumn<T> Create(string header, Func<T, float> selector) => new SingleColumn(header, selector);

        /// <summary>Creates a decimal column.</summary>
        public static ExcelTabularColumn<T> Create(string header, Func<T, decimal> selector) => new DecimalColumn(header, selector);

        /// <summary>Creates a signed 32-bit integer column.</summary>
        public static ExcelTabularColumn<T> Create(string header, Func<T, int> selector) => new Int32Column(header, selector);

        /// <summary>Creates a signed 64-bit integer column.</summary>
        public static ExcelTabularColumn<T> Create(string header, Func<T, long> selector) => new Int64Column(header, selector);

        /// <summary>Creates an unsigned 64-bit integer column.</summary>
        public static ExcelTabularColumn<T> Create(string header, Func<T, ulong> selector) => new UInt64Column(header, selector);

        /// <summary>Creates a column for another value type.</summary>
        public static ExcelTabularColumn<T> Create<TValue>(string header, Func<T, TValue> selector) => new ObjectColumn<TValue>(header, selector);

        private sealed class StringColumn : ExcelTabularColumn<T> {
            private readonly Func<T, string?> _selector;
            internal StringColumn(string header, Func<T, string?> selector) : base(header) => _selector = selector ?? throw new ArgumentNullException(nameof(selector));
            internal override Type DataType => typeof(string);
            internal override object? GetValue(T row) => _selector(row);
            internal override void WriteValue(ExcelDocument.ExcelTabularRowWriter writer, T row) => writer.Write(_selector(row));
        }

        private sealed class BooleanColumn : ExcelTabularColumn<T> {
            private readonly Func<T, bool> _selector;
            internal BooleanColumn(string header, Func<T, bool> selector) : base(header) => _selector = selector ?? throw new ArgumentNullException(nameof(selector));
            internal override Type DataType => typeof(bool);
            internal override object GetValue(T row) => _selector(row);
            internal override void WriteValue(ExcelDocument.ExcelTabularRowWriter writer, T row) => writer.Write(_selector(row));
        }

        private sealed class DateTimeColumn : ExcelTabularColumn<T> {
            private readonly Func<T, DateTime> _selector;
            internal DateTimeColumn(string header, Func<T, DateTime> selector) : base(header) => _selector = selector ?? throw new ArgumentNullException(nameof(selector));
            internal override Type DataType => typeof(DateTime);
            internal override object GetValue(T row) => _selector(row);
            internal override void WriteValue(ExcelDocument.ExcelTabularRowWriter writer, T row) => writer.Write(_selector(row));
        }

        private sealed class DateTimeOffsetColumn : ExcelTabularColumn<T> {
            private readonly Func<T, DateTimeOffset> _selector;
            internal DateTimeOffsetColumn(string header, Func<T, DateTimeOffset> selector) : base(header) => _selector = selector ?? throw new ArgumentNullException(nameof(selector));
            internal override Type DataType => typeof(DateTimeOffset);
            internal override object GetValue(T row) => _selector(row);
            internal override void WriteValue(ExcelDocument.ExcelTabularRowWriter writer, T row) => writer.Write(_selector(row));
        }

        private sealed class TimeSpanColumn : ExcelTabularColumn<T> {
            private readonly Func<T, TimeSpan> _selector;
            internal TimeSpanColumn(string header, Func<T, TimeSpan> selector) : base(header) => _selector = selector ?? throw new ArgumentNullException(nameof(selector));
            internal override Type DataType => typeof(TimeSpan);
            internal override object GetValue(T row) => _selector(row);
            internal override void WriteValue(ExcelDocument.ExcelTabularRowWriter writer, T row) => writer.Write(_selector(row));
        }

        private sealed class DoubleColumn : ExcelTabularColumn<T> {
            private readonly Func<T, double> _selector;
            internal DoubleColumn(string header, Func<T, double> selector) : base(header) => _selector = selector ?? throw new ArgumentNullException(nameof(selector));
            internal override Type DataType => typeof(double);
            internal override object GetValue(T row) => _selector(row);
            internal override void WriteValue(ExcelDocument.ExcelTabularRowWriter writer, T row) => writer.Write(_selector(row));
        }

        private sealed class SingleColumn : ExcelTabularColumn<T> {
            private readonly Func<T, float> _selector;
            internal SingleColumn(string header, Func<T, float> selector) : base(header) => _selector = selector ?? throw new ArgumentNullException(nameof(selector));
            internal override Type DataType => typeof(float);
            internal override object GetValue(T row) => _selector(row);
            internal override void WriteValue(ExcelDocument.ExcelTabularRowWriter writer, T row) => writer.Write(_selector(row));
        }

        private sealed class DecimalColumn : ExcelTabularColumn<T> {
            private readonly Func<T, decimal> _selector;
            internal DecimalColumn(string header, Func<T, decimal> selector) : base(header) => _selector = selector ?? throw new ArgumentNullException(nameof(selector));
            internal override Type DataType => typeof(decimal);
            internal override object GetValue(T row) => _selector(row);
            internal override void WriteValue(ExcelDocument.ExcelTabularRowWriter writer, T row) => writer.Write(_selector(row));
        }

        private sealed class Int32Column : ExcelTabularColumn<T> {
            private readonly Func<T, int> _selector;
            internal Int32Column(string header, Func<T, int> selector) : base(header) => _selector = selector ?? throw new ArgumentNullException(nameof(selector));
            internal override Type DataType => typeof(int);
            internal override object GetValue(T row) => _selector(row);
            internal override void WriteValue(ExcelDocument.ExcelTabularRowWriter writer, T row) => writer.Write(_selector(row));
        }

        private sealed class Int64Column : ExcelTabularColumn<T> {
            private readonly Func<T, long> _selector;
            internal Int64Column(string header, Func<T, long> selector) : base(header) => _selector = selector ?? throw new ArgumentNullException(nameof(selector));
            internal override Type DataType => typeof(long);
            internal override object GetValue(T row) => _selector(row);
            internal override void WriteValue(ExcelDocument.ExcelTabularRowWriter writer, T row) => writer.Write(_selector(row));
        }

        private sealed class UInt64Column : ExcelTabularColumn<T> {
            private readonly Func<T, ulong> _selector;
            internal UInt64Column(string header, Func<T, ulong> selector) : base(header) => _selector = selector ?? throw new ArgumentNullException(nameof(selector));
            internal override Type DataType => typeof(ulong);
            internal override object GetValue(T row) => _selector(row);
            internal override void WriteValue(ExcelDocument.ExcelTabularRowWriter writer, T row) => writer.Write(_selector(row));
        }

        private sealed class ObjectColumn<TValue> : ExcelTabularColumn<T> {
            private readonly Func<T, TValue> _selector;
            internal ObjectColumn(string header, Func<T, TValue> selector) : base(header) => _selector = selector ?? throw new ArgumentNullException(nameof(selector));
            internal override Type DataType => typeof(TValue);
            internal override object? GetValue(T row) => _selector(row);
            internal override void WriteValue(ExcelDocument.ExcelTabularRowWriter writer, T row) => writer.Write(_selector(row));
        }
    }
}
