using System;

namespace OfficeIMO.Shared {
    /// <summary>
    /// Represents one value from an OLE property set stream.
    /// </summary>
    internal sealed class OfficeOlePropertyValue {
        internal OfficeOlePropertyValue(ushort type, object? value) {
            Type = type;
            Value = value;
        }

        internal ushort Type { get; }

        internal object? Value { get; }

        internal string? AsString() {
            return Value as string;
        }

        internal DateTime? AsDateTime() {
            return Value as DateTime?;
        }
    }
}
