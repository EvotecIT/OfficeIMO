using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing.Internal {
    /// <summary>
    /// Represents a parsed section inside an OLE property set stream.
    /// </summary>
    internal sealed class OfficeOlePropertySection {
        internal OfficeOlePropertySection(Guid formatId,
            IReadOnlyDictionary<uint, OfficeOlePropertyValue> properties,
            IReadOnlyDictionary<uint, string> dictionary) {
            FormatId = formatId;
            Properties = properties;
            Dictionary = dictionary;
        }

        internal Guid FormatId { get; }

        internal IReadOnlyDictionary<uint, OfficeOlePropertyValue> Properties { get; }

        internal IReadOnlyDictionary<uint, string> Dictionary { get; }
    }
}
