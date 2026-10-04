using System.Threading;

namespace OfficeIMO.IWork.Internal;

internal static partial class IWorkSnappy {
    /// <summary>Emits literal-only raw Snappy blocks inside Apple's IWA framing, with at most 64 KiB decoded per block.</summary>
    internal static byte[] EncodeIwa(byte[] source, int maximumBytes, CancellationToken cancellationToken) {
        using var output = new MemoryStream();
        for (int offset = 0; offset < source.Length;) {
            cancellationToken.ThrowIfCancellationRequested();
            int length = Math.Min(64 * 1024, source.Length - offset);
            using var raw = new MemoryStream(length + 8);
            IWorkProtoWriter.WriteVarint(raw, (ulong)length);
            int encoded = length - 1;
            if (length <= 60) {
                raw.WriteByte((byte)(encoded << 2));
            } else {
                int width = encoded <= byte.MaxValue ? 1 : 2;
                raw.WriteByte((byte)((59 + width) << 2));
                for (int index = 0; index < width; index++) raw.WriteByte((byte)(encoded >> (8 * index)));
            }
            raw.Write(source, offset, length);
            int rawLength = checked((int)raw.Length);
            if (output.Length > (long)maximumBytes - rawLength - 4) {
                throw new InvalidDataException($"Native Keynote IWA data exceeds the configured {maximumBytes}-byte limit.");
            }
            output.WriteByte(0);
            output.WriteByte((byte)rawLength);
            output.WriteByte((byte)(rawLength >> 8));
            output.WriteByte((byte)(rawLength >> 16));
            raw.WriteTo(output);
            offset += length;
        }
        cancellationToken.ThrowIfCancellationRequested();
        return output.ToArray();
    }
}
