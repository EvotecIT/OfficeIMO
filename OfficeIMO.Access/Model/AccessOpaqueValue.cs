namespace OfficeIMO.Access {
    /// <summary>Unqualified native value retained without coercion. The raw representation is inert and returned as a defensive copy.</summary>
    public sealed class AccessOpaqueValue {
        private readonly byte[] _bytes;
        internal AccessOpaqueValue(uint type, byte[] bytes, string reason) { NativeType = type; _bytes = bytes; Reason = reason; }
        /// <summary>Persisted native type code, including full 32-bit designer property types.</summary>
        public uint NativeType { get; }
        /// <summary>Reason typed interpretation is unavailable. It contains no native content or credentials.</summary>
        public string Reason { get; }
        /// <summary>Number of retained raw bytes.</summary>
        public int Length => _bytes.Length;
        /// <summary>Copies the exact retained field representation. Calling this does not execute or resolve its content.</summary>
        public byte[] GetBytes() => (byte[])_bytes.Clone();
    }
}
