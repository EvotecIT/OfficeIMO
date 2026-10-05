using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public static partial class OfficeTiffCodec {
    // The managed subset decodes unsigned eight- or sixteen-bit components. Signed, floating
    // point and undefined component encodings must not be interpreted as RGB bytes.
    private static bool TryGetSampleByteCount(byte[] bytes, IReadOnlyDictionary<int, TiffEntry> entries,
        bool littleEndian, int samples, int photometric, out int sampleBytes) {
        sampleBytes = 0;
        if (!TryReadScalarOrDefault(bytes, entries, 266, littleEndian, 1, out int fillOrder) ||
            fillOrder != 1 ||
            !TryReadValues(bytes, entries, 258, littleEndian, samples, out int[] bits) ||
            (bits[0] != 8 && bits[0] != 16) ||
            Array.Exists(bits, bit => bit != bits[0]) ||
            (photometric == 3 && bits[0] != 8) ||
            !HasUnsignedSamples(bytes, entries, littleEndian, samples)) return false;
        sampleBytes = bits[0] / 8;
        return true;
    }

    private static bool HasUnsignedSamples(byte[] bytes, IReadOnlyDictionary<int, TiffEntry> entries,
        bool littleEndian, int samples) =>
        !entries.ContainsKey(339) ||
        TryReadValues(bytes, entries, 339, littleEndian, samples, out int[] formats) &&
        Array.TrueForAll(formats, format => format == 1);
}
