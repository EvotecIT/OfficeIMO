namespace OfficeIMO.Access {
    /// <summary>One inert native saved-query record, including uninterpreted attributes and exact row bytes.</summary>
    public sealed class AccessQueryRecord {
        internal AccessQueryRecord(byte attribute, string? name1, string? name2, string? expression, short? flag, int? extra, byte[]? order, byte[] nativeBytes) {
            Attribute = attribute; Name1 = name1; Name2 = name2; Expression = expression; Flag = flag; Extra = extra; _order = order;
            NativeBytes = new AccessOpaqueValue(attribute, nativeBytes, "Exact MSysQueries row; reading never executes SQL.");
        }
        private readonly byte[]? _order;
        /// <summary>Persisted query-record attribute.</summary>
        public byte Attribute { get; }
        /// <summary>First persisted name/argument.</summary>
        public string? Name1 { get; }
        /// <summary>Second persisted name/argument.</summary>
        public string? Name2 { get; }
        /// <summary>Persisted SQL fragment, evaluated by no document-library operation.</summary>
        public string? Expression { get; }
        /// <summary>Optional persisted flag/type.</summary>
        public short? Flag { get; }
        /// <summary>Optional persisted extra argument.</summary>
        public int? Extra { get; }
        /// <summary>Exact record representation.</summary>
        public AccessOpaqueValue NativeBytes { get; }
        /// <summary>Returns a defensive copy of the native ordering key.</summary>
        public byte[]? GetOrderBytes() => _order == null ? null : (byte[])_order.Clone();
    }

    /// <summary>Inert declared query parameter. Unknown type codes remain available without coercion.</summary>
    public sealed class AccessQueryParameter {
        internal AccessQueryParameter(string name, int nativeType, AccessDataType type) { Name = name; NativeType = nativeType; DataType = type; }
        /// <summary>Parameter name.</summary>
        public string Name { get; }
        /// <summary>Native declared type code.</summary>
        public int NativeType { get; }
        /// <summary>Common type, or Unknown for an unqualified declaration.</summary>
        public AccessDataType DataType { get; }
    }
}
