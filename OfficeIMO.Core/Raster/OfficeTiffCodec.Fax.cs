using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace OfficeIMO.Drawing;

public static partial class OfficeTiffCodec {
    private static bool IsTiffFaxCompression(int compression) => compression >= 2 && compression <= 4;

    private static bool TryGetTiffFaxSettings(byte[] bytes, IReadOnlyDictionary<int, TiffEntry> entries,
        bool littleEndian, int compression, out int k, out int fillOrder) {
        k = 0; fillOrder = 1;
        if (!IsTiffFaxCompression(compression)) return true;
        if (!TryReadScalarOrDefault(bytes, entries, 266, littleEndian, 1, out fillOrder) ||
            (fillOrder != 1 && fillOrder != 2)) return false;
        if (compression == 3) {
            if (!TryReadScalarOrDefault(bytes, entries, 292, littleEndian, 0, out int options) ||
                (options & ~7) != 0) return false;
            k = (options & 1) != 0 ? 1 : 0;
        } else if (compression == 4) {
            if (!TryReadScalarOrDefault(bytes, entries, 293, littleEndian, 0, out int options) || (options & ~2) != 0) return false;
            k = -1;
        }
        return true;
    }

    private static bool TryDecodeTiffFax(byte[] bytes, int offset, int count, int columns, int rows,
        int compression, int k, int fillOrder, byte[] target, int expected, CancellationToken token) {
        // The packed segment caller accounts for both this copy and the fax output.
        var encoded = new byte[count];
        CopyWithCancellation(bytes, offset, encoded, 0, count, token);
        if (fillOrder == 2) {
            for (int i = 0; i < encoded.Length; i++) {
                if ((i & 4095) == 0) token.ThrowIfCancellationRequested();
                int value = encoded[i];
                value = ((value & 0x55) << 1) | ((value >> 1) & 0x55);
                value = ((value & 0x33) << 2) | ((value >> 2) & 0x33);
                encoded[i] = (byte)((value << 4) | (value >> 4));
            }
        }
        try {
            // TIFF permits stopping after the declared rows without checking EOFB/RTC.
            // Keep raw zero/one samples; the common photometric path applies polarity.
            byte[] decoded = OfficeFaxDecoder.Decode(encoded, columns, rows, k,
                endOfLine: compression == 3, byteAligned: compression == 2,
                blackIsOne: true, endOfBlock: false, maximumBytes: expected, cancellationToken: token);
            if (decoded.Length != expected) return false;
            CopyWithCancellation(decoded, 0, target, 0, expected, token);
            return true;
        } catch (InvalidDataException) {
            return false;
        }
    }
}
