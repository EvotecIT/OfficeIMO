using System.Threading;

namespace OfficeIMO.Drawing;

public static partial class OfficeImageReader {
    /// <summary>Recognizes an AVIF brand even when later container metadata is malformed.</summary>
    internal static bool HasAvifSignature(byte[] bytes, CancellationToken cancellationToken) {
        int offset = 0;
        for (int boxes = 0; boxes < 4096 && offset <= bytes.Length - 8; boxes++) {
            cancellationToken.ThrowIfCancellationRequested();
            ulong length = ReadUInt32BigEndian(bytes, offset);
            int header = length == 1 ? 16 : 8;
            if (offset > bytes.Length - header) return false;
            if (length == 1)
                length = (ulong)ReadUInt32BigEndian(bytes, offset + 8) << 32 | ReadUInt32BigEndian(bytes, offset + 12);
            else if (length == 0) length = (ulong)(bytes.Length - offset);
            if (GetAscii(bytes, offset + 4, 4) == "ftyp") {
                int start = offset + header;
                if (start <= bytes.Length - 4 && IsAvifBrand(bytes, start)) return true;
                int end = offset + (int)System.Math.Min(length, (ulong)(bytes.Length - offset));
                for (int p = start + 8; p <= end - 4; p += 4) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (IsAvifBrand(bytes, p)) return true;
                }
            }
            if (length < (ulong)header || length > (ulong)(bytes.Length - offset)) return false;
            offset += (int)length;
        }
        return false;
    }

    private static bool IsAvifBrand(byte[] bytes, int offset) =>
        bytes[offset] == 'a' && bytes[offset + 1] == 'v' && bytes[offset + 2] == 'i' &&
        (bytes[offset + 3] == 'f' || bytes[offset + 3] == 's');

    /// <summary>Identifies whole bounded still items without decoding their entropy payload.</summary>
    private static bool TryReadAvif(byte[] bytes, CancellationToken cancellationToken, out OfficeImageInfo info) {
        info = new OfficeImageInfo(OfficeImageFormat.Unknown, 0, 0);
        if (!OfficeAvifContainerReader.TryRead(bytes, new OfficeRasterDecodeOptions { CancellationToken = cancellationToken }, out var container))
            return false;
        info = new OfficeImageInfo(OfficeImageFormat.Avif, container!.Color.Width, container.Color.Height);
        return true;
    }
}
