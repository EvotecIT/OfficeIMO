using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public static partial class OfficeTiffCodec {
    private static bool TryGetSampleByteCount(byte[] bytes, IReadOnlyDictionary<int, TiffEntry> entries,
        bool littleEndian, int samples, int photometric, out int sampleBytes, out bool floating, out int packedBits) {
        sampleBytes = 0; floating = false; packedBits = 0;
        if (!TryReadScalarOrDefault(bytes, entries, 266, littleEndian, 1, out int fillOrder) ||
            fillOrder != 1) return false;
        int[] bits;
        if (entries.ContainsKey(258)) {
            if (!TryReadValues(bytes, entries, 258, littleEndian, samples, out bits) ||
                Array.Exists(bits, bit => bit != bits[0])) return false;
        } else bits = new[] { 1 }; // Baseline bilevel default.

        int format = 1;
        if (entries.ContainsKey(339)) {
            if (!TryReadValues(bytes, entries, 339, littleEndian, samples, out int[] formats) ||
                Array.Exists(formats, value => value != formats[0])) return false;
            format = formats[0];
        }
        if (format == 1 && (bits[0] == 1 || bits[0] == 4) && samples == 1 &&
            (photometric == 0 || photometric == 1 || photometric == 3)) {
            sampleBytes = 1;
            packedBits = bits[0];
            return true;
        }
        floating = format == 3;
        if (floating ? (bits[0] != 16 && bits[0] != 24 && bits[0] != 32 && bits[0] != 64) || photometric == 3
            : format != 1 || (bits[0] != 8 && bits[0] != 16) || (photometric == 3 && bits[0] != 8)) return false;
        sampleBytes = bits[0] / 8;
        return true;
    }

    private static bool IsSupportedSamplePredictor(int predictor, bool floating, int compression) =>
        predictor == 1 || !floating && predictor == 2 || floating && predictor == 3 &&
        (compression == 5 || compression == 8 || compression == 32946);
}
