using System.Collections.Generic;

namespace OfficeIMO.Drawing.Internal {
    /// <summary>
    /// Represents a parsed section inside an OLE property set stream.
    /// </summary>
    internal sealed class OfficeOlePropertySection {
        internal OfficeOlePropertySection(
            IReadOnlyDictionary<uint, OfficeOlePropertyValue> properties,
            IReadOnlyDictionary<uint, string> dictionary) {
            Properties = properties;
            Dictionary = dictionary;
        }

        internal IReadOnlyDictionary<uint, OfficeOlePropertyValue> Properties { get; }

        internal IReadOnlyDictionary<uint, string> Dictionary { get; }
    }
}
