namespace OfficeIMO.Access {
    /// <summary>Field types reached by the in-memory foundation examples. Native field decoding is separately qualified.</summary>
    public enum AccessDataType {
        /// <summary>Engine-generated integer; omission does not allocate a value in the model.</summary>
        AutoNumber,
        /// <summary>Signed 32-bit integer.</summary>
        Int32,
        /// <summary>Unicode text with a declared maximum length.</summary>
        ShortText,
        /// <summary>Long Unicode text.</summary>
        LongText,
        /// <summary>Currency represented by Decimal with four fractional digits.</summary>
        Currency,
        /// <summary>64-bit floating-point value.</summary>
        Double,
        /// <summary>Date/time value without implicit timezone conversion.</summary>
        DateTime,
        /// <summary>Boolean value.</summary>
        Boolean,
        /// <summary>128-bit identifier.</summary>
        Guid,
        /// <summary>Binary data. Values are copied at the model boundary.</summary>
        Binary,
        /// <summary>Unsigned 8-bit integer.</summary>
        Byte,
        /// <summary>Signed 16-bit integer.</summary>
        Int16,
        /// <summary>32-bit floating-point value.</summary>
        Single,
        /// <summary>Fixed-precision decimal value. Values beyond CLR Decimal retain an exact numeric representation.</summary>
        Decimal,
        /// <summary>Signed 64-bit Large Number.</summary>
        Int64,
        /// <summary>Structured native complex value, such as an attachment or multivalued field.</summary>
        Complex,
        /// <summary>Extended native date/time whose precision is retained explicitly.</summary>
        ExtendedDateTime,
        /// <summary>Unqualified native field; exact field bytes are retained as an opaque value.</summary>
        Unknown
    }

    /// <summary>Input field values. A missing key differs from an explicit null.</summary>
    public sealed class AccessRowValues : Dictionary<string, object?> {
        /// <summary>Creates a case-insensitive Access field-value map.</summary>
        public AccessRowValues() : base(StringComparer.OrdinalIgnoreCase) { }
    }

    /// <summary>A typed table in a document model.</summary>
    public sealed class AccessTable : AccessNamedObject {
        internal readonly List<Dictionary<string, object?>> Rows = new List<Dictionary<string, object?>>();
        internal AccessNativeTable? NativeTable;
        internal AccessTable(AccessDocument document, string name) : base(document, name) {
            Columns = new AccessColumnCollection(this); Indexes = new AccessIndexCollection(this);
        }
        /// <summary>Field definitions in ordinal order.</summary>
        public AccessColumnCollection Columns { get; }
        /// <summary>Typed index definitions.</summary>
        public AccessIndexCollection Indexes { get; }
        /// <summary>Number of modeled rows, without reading any native table.</summary>
        public long RowCount { get { EnsureAttached(); if (IsLinked) throw new NotSupportedException("Linked-table row counts require a separately authorized provider; targets are never resolved by the document codec."); return Document.ResolveNativeReadTable(NativeTable)?.RowCount ?? Rows.Count; } }
        /// <summary>Whether this is a system or hidden table kept outside the user-table collection.</summary>
        public bool IsSystem { get; internal set; }
        /// <summary>Whether this definition refers to an external table whose target is never opened.</summary>
        public bool IsLinked => LinkedTable != null;
        /// <summary>Inert, credential-redacted linked-table metadata, or null for a local table.</summary>
        public AccessLinkedTableInfo? LinkedTable { get; internal set; }
        /// <summary>Persisted table properties; unknown property values retain their raw representation.</summary>
        public IReadOnlyDictionary<string, object?> Properties { get; internal set; } = new System.Collections.ObjectModel.ReadOnlyDictionary<string, object?>(new Dictionary<string, object?>());
        /// <summary>Exact persisted property-map bytes, including uninterpreted chunks. Null for a new model or an absent map.</summary>
        public AccessOpaqueValue? NativeProperties { get; internal set; }
        /// <summary>Appends validated values to the model, retaining omitted fields separately from explicit nulls.</summary>
        public void AppendRow(AccessRowValues values) {
            EnsureAttached(); Document.EnsureMutable();
            if (values == null) throw new ArgumentNullException(nameof(values));
            Dictionary<string, object?> row = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
            foreach (KeyValuePair<string, object?> pair in values) {
                AccessColumn column = Columns[pair.Key];
                column.ValidateValue(pair.Value);
                row.Add(column.Name, CopyValue(pair.Value));
            }
            Rows.Add(row); Document.Changed(() => Rows.RemoveAt(Rows.Count - 1), Id, "row.append");
        }
        internal static object? CopyValue(object? value) => value is byte[] bytes ? (byte[])bytes.Clone() : value;
        /// <summary>Opens a forward-only reader over modeled rows. Its lease blocks edits until disposal.</summary>
        public AccessDataReader OpenDataReader(CancellationToken cancellationToken = default) {
            EnsureAttached(); cancellationToken.ThrowIfCancellationRequested();
            if (IsLinked) throw new NotSupportedException("Linked tables are inert metadata; no target connection is opened.");
            return new AccessDataReader(this, cancellationToken);
        }
    }

    /// <summary>Field definitions for one table.</summary>
    public sealed class AccessColumnCollection : AccessObjectCollection<AccessColumn> {
        private readonly AccessTable _table;
        internal AccessColumnCollection(AccessTable table) : base(table.Document) { _table = table; }
        /// <summary>Adds a sequential AutoNumber field whose first omitted value is seed. The model retains omission; native creation allocates values.</summary>
        public AccessColumn AddAutoNumber(string name, int seed = 1) {
            if (seed < 1) throw new ArgumentOutOfRangeException(nameof(seed));
            AccessColumn column = Add(name, AccessDataType.AutoNumber); column.AutoNumberSeed = seed; return column;
        }
        /// <summary>Adds a Decimal field with explicit precision and scale. Native saving rejects values requiring rounding.</summary>
        public AccessColumn AddDecimal(string name, int precision = 28, int scale = 0) {
            if (precision < 1 || precision > 28) throw new ArgumentOutOfRangeException(nameof(precision));
            if (scale < 0 || scale > precision) throw new ArgumentOutOfRangeException(nameof(scale));
            AccessColumn column = Add(name, AccessDataType.Decimal); column.Precision = precision; column.Scale = scale; return column;
        }
        /// <summary>Adds a typed column before rows have been appended. ShortText defaults to 255 characters.</summary>
        public AccessColumn Add(string name, AccessDataType type, int? maxLength = null) {
            _table.EnsureAttached(); Document.EnsureMutable();
            if (_table.Rows.Count != 0) throw new InvalidOperationException("Define columns before appending rows.");
            if (Items.Count == 255) throw new InvalidOperationException("An Access table cannot declare more than 255 columns.");
            AccessColumn column = new AccessColumn(_table, name, type, maxLength); AddItem(column); return column;
        }
    }

    /// <summary>A typed column with immutable definition in the foundation slice.</summary>
    public sealed class AccessColumn : AccessNamedObject {
        internal AccessColumn(AccessTable table, string name, AccessDataType type, int? maxLength) : base(table.Document, name) {
            if (!Enum.IsDefined(typeof(AccessDataType), type)) throw new ArgumentOutOfRangeException(nameof(type));
            if (type == AccessDataType.ShortText) { maxLength ??= 255; if (maxLength < 1 || maxLength > 255) throw new ArgumentOutOfRangeException(nameof(maxLength)); }
            else if (maxLength != null) throw new ArgumentException("maxLength is declared only for ShortText.", nameof(maxLength));
            Table = table; DataType = type; MaxLength = maxLength; IsAutoNumber = type == AccessDataType.AutoNumber; AutoNumberSeed = IsAutoNumber ? 1 : (int?)null;
        }
        /// <summary>Owning table.</summary>
        public AccessTable Table { get; }
        /// <summary>Declared field type.</summary>
        public AccessDataType DataType { get; }
        /// <summary>Maximum text length, or null for other types.</summary>
        public int? MaxLength { get; }
        /// <summary>Whether the native field allocates an engine-generated integer or GUID. Reading does not allocate new values.</summary>
        public bool IsAutoNumber { get; internal set; }
        /// <summary>Initial sequential seed for an authored AutoNumber field. Null for loaded native fields whose initial seed has not been decoded.</summary>
        public int? AutoNumberSeed { get; internal set; }
        /// <summary>Whether this native text field stores hyperlink syntax. Its stored text is retained unchanged.</summary>
        public bool IsHyperlink { get; internal set; }
        /// <summary>Whether this native memo field has rich-text formatting. Reading does not render or strip markup.</summary>
        public bool IsRichText { get; internal set; }
        /// <summary>Whether the native field is calculated. Loading never evaluates its expression.</summary>
        public bool IsCalculated { get; internal set; }
        /// <summary>Native structured-field definition, or null for a scalar/model field.</summary>
        public AccessComplexDefinition? ComplexDefinition { get; internal set; }
        /// <summary>Native numeric precision, or null when no precision is declared.</summary>
        public int? Precision { get; internal set; }
        /// <summary>Native numeric scale, or null when no scale is declared.</summary>
        public int? Scale { get; internal set; }
        /// <summary>Persisted field properties, including required/default/validation and lookup metadata. Expressions remain inert.</summary>
        public IReadOnlyDictionary<string, object?> Properties { get; internal set; } = new System.Collections.ObjectModel.ReadOnlyDictionary<string, object?>(new Dictionary<string, object?>());
        internal void ValidateValue(object? value) {
            if (value == null) return;
            bool valid = DataType switch {
                AccessDataType.AutoNumber or AccessDataType.Int32 => value is int,
                AccessDataType.ShortText => value is string text && text.Length <= MaxLength,
                AccessDataType.LongText => value is string,
                AccessDataType.Currency => value is decimal amount && amount >= -922337203685477.5808m && amount <= 922337203685477.5807m && decimal.Round(amount, 4) == amount,
                AccessDataType.Double => value is double number && !double.IsNaN(number) && !double.IsInfinity(number),
                AccessDataType.DateTime or AccessDataType.ExtendedDateTime => value is DateTime,
                AccessDataType.Boolean => value is bool,
                AccessDataType.Guid => value is Guid,
                AccessDataType.Binary => value is byte[], AccessDataType.Byte => value is byte,
                AccessDataType.Int16 => value is short, AccessDataType.Int64 => value is long,
                AccessDataType.Single => value is float single && !float.IsNaN(single) && !float.IsInfinity(single),
                AccessDataType.Decimal => value is decimal, _ => false
            };
            if (!valid) throw new ArgumentException($"Value for '{Name}' does not satisfy {DataType}.");
        }
    }
}
