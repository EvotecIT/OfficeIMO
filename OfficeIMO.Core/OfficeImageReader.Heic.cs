using System.Threading;

namespace OfficeIMO.Drawing;

public static partial class OfficeImageReader {
    internal static bool HasHeicSignature(byte[] bytes, CancellationToken cancellationToken) =>
        ProbeHeicSignature(bytes, cancellationToken, rejectUninspected: false);

    // An oversized brand table is indeterminate, not permission to bypass inspection.
    internal static bool MayContainHeic(byte[] bytes, CancellationToken cancellationToken) =>
        ProbeHeicSignature(bytes, cancellationToken, rejectUninspected: true);

    private static bool ProbeHeicSignature(byte[] bytes, CancellationToken cancellationToken, bool rejectUninspected) {
        for (int p = 0, boxes = 0; boxes < 4096 && p <= bytes.Length - 16; boxes++) {
            cancellationToken.ThrowIfCancellationRequested();
            ulong size = ReadUInt32BigEndian(bytes, p);
            int header = size == 1 ? 16 : 8;
            if (size == 1) {
                if (p > bytes.Length - 24) return false;
                size = (ulong)ReadUInt32BigEndian(bytes, p + 8) << 32 | ReadUInt32BigEndian(bytes, p + 12);
            } else if (size == 0) size = (ulong)(bytes.Length - p);
            if (size < (ulong)header || size > (ulong)(bytes.Length - p)) return false;
            if (ReadUInt32BigEndian(bytes, p + 4) == 0x66747970) { // ftyp
                if (size < (ulong)header + 8 || (size - (ulong)header) % 4 != 0) return false;
                // A recognized major brand remains a rejection guard even when the rest
                // of the table is too large for metadata parsing.
                if (IsHeicBrand(ReadUInt32BigEndian(bytes, p + header))) return true;
                if ((size - (ulong)header) / 4 > 4096) return rejectUninspected;
                for (int b = p + header + 8; b <= p + (int)size - 4; b += 4) {
                    cancellationToken.ThrowIfCancellationRequested();
                    uint brand = ReadUInt32BigEndian(bytes, b);
                    if (IsHeicBrand(brand)) return true;
                }
                return false;
            }
            p += (int)size;
        }
        return false;
    }

    private static bool IsHeicBrand(uint brand) => brand is 0x68656963 or 0x68656978 or 0x68657663 or
        0x6865696D or 0x68656973 or 0x6D696631 or 0x6D736631;

    private static bool TryReadHeic(byte[] bytes, CancellationToken cancellationToken, out OfficeImageInfo info) {
        info = new OfficeImageInfo(OfficeImageFormat.Unknown, 0, 0);
        if (!OfficeHeifImageComposer.TryIdentify(bytes, new OfficeRasterDecodeOptions { CancellationToken = cancellationToken }, out int width, out int height)) return false;
        info = new OfficeImageInfo(OfficeImageFormat.Heic, width, height);
        return true;
    }
}
