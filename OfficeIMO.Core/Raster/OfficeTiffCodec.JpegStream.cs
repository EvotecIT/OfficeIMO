using System;
using System.Threading;

namespace OfficeIMO.Drawing;

public static partial class OfficeTiffCodec {
    // TIFF owns component order and color interpretation. APP metadata must not
    // override its orientation, CMYK polarity or photometric tags.
    private static bool TryNormalizeTiffJpeg(byte[] data, bool tablesOnly, int width, int height,
        int samples, int horizontal, int vertical, int inheritedTables, CancellationToken token,
        out int definedTables) {
        definedTables = 0;
        if (data.Length < 4 || data[0] != 255 || data[1] != 216) return false;
        int offset = 2;
        int[] ids = new int[samples];
        bool frame = false, scan = false;
        while (offset < data.Length) {
            token.ThrowIfCancellationRequested();
            if (data[offset++] != 255) return false;
            while (offset < data.Length && data[offset] == 255) {
                if ((offset & 4095) == 0) token.ThrowIfCancellationRequested();
                offset++;
            }
            if (offset >= data.Length) return false;
            int markerOffset = offset, marker = data[offset++];
            if (marker == 217) return offset == data.Length && (tablesOnly || frame && scan);
            if (offset + 2 > data.Length) return false;
            int length = (data[offset] << 8) | data[offset + 1];
            if (length < 2 || length > data.Length - offset) return false;
            int start = offset + 2, end = offset + length;
            if (tablesOnly && (marker == 204 || marker == 221)) data[markerOffset] = 254; // SOI resets DAC/DRI before each image.
            else if (marker >= 224 && marker <= 239) data[markerOffset] = 254; // COM is ignored by the JPEG owner.
            else if (marker == 219 || marker == 196) {
                int p = start;
                while (p < end) {
                    int info = data[p++], id = info & 15, kind = info >> 4;
                    if (id > 3 || (marker == 219 ? kind != 0 : kind > 1)) return false;
                    int mask = 1 << (marker == 219 ? id : 4 + kind * 4 + id);
                    if ((inheritedTables & mask) != 0) return false;
                    definedTables |= mask;
                    int count = 64;
                    if (marker == 196) {
                        if (end - p < 16) return false;
                        count = 0;
                        for (int i = 0; i < 16; i++) count += data[p++];
                    }
                    if (count > end - p) return false;
                    p += count;
                }
            } else if (marker == 192 && !tablesOnly) {
                if (frame || length != 8 + samples * 3 || data[start] != 8 ||
                    ((data[start + 1] << 8) | data[start + 2]) != height ||
                    ((data[start + 3] << 8) | data[start + 4]) != width || data[start + 5] != samples) return false;
                for (int i = 0; i < samples; i++) {
                    int p = start + 6 + 3 * i;
                    ids[i] = data[p];
                    for (int j = 0; j < i; j++) if (ids[j] == ids[i]) return false;
                    if (data[p + 1] != (i == 0 ? (horizontal << 4) | vertical : 17)) return false;
                    data[p] = (byte)(i + 1);
                }
                frame = true;
            } else if (marker == 218 && !tablesOnly) {
                if (!frame || length < 6) return false;
                int count = data[start];
                if (count < 1 || count > samples || length != 6 + 2 * count) return false;
                for (int i = 0; i < count; i++) {
                    int p = start + 1 + 2 * i, index = Array.IndexOf(ids, (int)data[p]);
                    if (index < 0) return false;
                    data[p] = (byte)(index + 1);
                }
                scan = true;
                offset = end;
                int scanSteps = 0;
                while (offset < data.Length) {
                    if ((scanSteps++ & 4095) == 0) token.ThrowIfCancellationRequested();
                    if (data[offset] != 255) { offset++; continue; }
                    if (offset + 1 == data.Length) return false;
                    int next = data[offset + 1];
                    if (next == 0 || next >= 208 && next <= 215) { offset += 2; continue; }
                    break;
                }
                continue;
            } else if (marker != 254 && !(marker == 221 && !tablesOnly && length == 4)) return false;
            offset = end;
        }
        return false;
    }
}
