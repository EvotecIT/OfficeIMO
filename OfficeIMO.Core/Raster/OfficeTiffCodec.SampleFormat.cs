using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public static partial class OfficeTiffCodec {
    // The managed subset decodes unsigned eight-bit components. Signed, floating
    // point and undefined component encodings must not be interpreted as RGB bytes.
    private static bool HasUnsignedSamples(byte[] bytes, IReadOnlyDictionary<int, TiffEntry> entries,
        bool littleEndian, int samples) =>
        !entries.ContainsKey(339) ||
        TryReadValues(bytes, entries, 339, littleEndian, samples, out int[] formats) &&
        Array.TrueForAll(formats, format => format == 1);
}
