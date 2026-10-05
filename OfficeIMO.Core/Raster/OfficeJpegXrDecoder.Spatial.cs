using System;
using System.Threading;

namespace OfficeIMO.Drawing;

internal static partial class OfficeJpegXrDecoder {
    internal sealed class ReconstructedFrame {
        internal DclpPlane Primary = new();
        internal DclpPlane? Alpha;
        internal BandData? Highpass;
    }

    private sealed class SpatialPlane {
        internal readonly PlaneHeader Header;
        internal readonly int[] Dc, Prediction, LpIndices, Highpass, Patterns;
        internal readonly byte[] ModelBits;
        internal readonly DclpPlane Result;
        internal int[] DcQuant = Array.Empty<int>();
        internal int[][] LpQuant = Array.Empty<int[]>(), HpQuant = Array.Empty<int[]>();
        internal bool ReuseLp;
        internal DcContext DcReader = null!;
        internal LpContext? LpReader;
        internal HpContext? HpReader;

        internal SpatialPlane(PlaneHeader header, int count) {
            Header = header;
            Dc = new int[checked(count * header.Components)];
            Prediction = new int[checked(count * header.Components * 16)];
            LpIndices = header.Bands == 3 ? Array.Empty<int>() : new int[count];
            Highpass = header.Bands < 2 ? new int[checked(count * header.Components * 256)] : Array.Empty<int>();
            Patterns = header.Bands < 2 ? new int[checked(count * header.Components)] : Array.Empty<int>();
            ModelBits = header.Bands < 2 ? new byte[checked(count * 2)] : Array.Empty<byte>();
            Result = new DclpPlane { Coefficients = new int[Prediction.Length], HighpassModes = new byte[count] };
        }

        internal void StartTile(Bits bits) {
            DcQuant = Header.DcQuant ?? ReadQuantization(bits, Header.Components);
            LpQuant = ReadLpQuantization(bits, Header, DcQuant);
            HpQuant = ReadHpQuantization(bits, Header, LpQuant, out ReuseLp);
            DcReader = new DcContext(Header.Components, Header.Color);
            LpReader = Header.Bands < 3 ? new LpContext(Header.Components, Header.Color) : null;
            HpReader = Header.Bands < 2 ? new HpContext(Header.Components, Header.Color) : null;
        }

        internal void Read(Bits bits, int columns, int mb, int x, int y, int tileWidth, int trim) {
            bool adapt = x == tileWidth - 1 || x % 16 == 0;
            int lpIndex = LpReader == null ? 0 : ReadQuantizerIndex(bits, LpQuant.Length);
            int hpIndex = HpReader == null ? 0 : ReuseLp ? lpIndex : ReadQuantizerIndex(bits, HpQuant.Length);
            if (LpReader != null) LpIndices[mb] = lpIndex;
            DcReader.Read(bits, Dc, mb * Header.Components, adapt);
            LpReader?.Read(bits, Prediction, mb * Header.Components * 16, x % 16 == 0, adapt);
            ReconstructMacroblockDclp(Header, columns, mb, x == 0, y == 0, Dc, Prediction,
                LpReader == null ? null : LpIndices, DcQuant, LpReader == null ? null : LpQuant[lpIndex], Result);
            if (HpReader != null) {
                HpReader.Read(bits, Highpass, mb * Header.Components * 256, Patterns, mb * Header.Components,
                    (mb - 1) * Header.Components, (mb - columns) * Header.Components,
                    x == 0, y == 0, x % 16 == 0, adapt, Result.HighpassModes[mb], ModelBits, mb * 2,
                    Header.Bands == 0, trim);
                ReconstructMacroblockHp(Highpass, mb * Header.Components * 256, Header, HpQuant[hpIndex], Result.HighpassModes[mb]);
            }
        }
    }

    // T.832 8.7.1: spatial tile headers contain all bands for the primary plane,
    // then alpha. Each macroblock follows that same complete-plane order.
    internal static ReconstructedFrame ReadSpatial(byte[] bytes, FrameHeader frame, PacketMap packets,
            CancellationToken cancellation) {
        if (frame.Frequency) throw new FormatException("JPEG-XR spatial parsing requires spatial packets.");
        int columns = (frame.Width + frame.Left + frame.Right) / 16;
        int rows = (frame.Height + frame.Top + frame.Bottom) / 16, count = checked(columns * rows);
        long components = frame.Primary.Components + (frame.AlphaPlane?.Components ?? 0);
        if (count * components * (256L + 33) * 4 + bytes.Length > 256 * 1024 * 1024)
            throw new FormatException("JPEG-XR spatial working set exceeds the managed limit.");
        var primary = new SpatialPlane(frame.Primary, count);
        SpatialPlane? alpha = frame.AlphaPlane == null ? null : new SpatialPlane(frame.AlphaPlane, count);
        int tile = 0, top = 0;
        foreach (int height in frame.TileHeights) {
            int left = 0;
            foreach (int width in frame.TileWidths) {
                cancellation.ThrowIfCancellationRequested();
                int start = packets.Offsets[tile];
                var bits = new Bits(bytes, start + 4, packets.Lengths[tile] - 4, cancellation);
                int trim = frame.Trim ? (int)bits.Read(4) : 0;
                primary.StartTile(bits); alpha?.StartTile(bits);
                for (int y = 0; y < height; y++) for (int x = 0; x < width; x++) {
                    cancellation.ThrowIfCancellationRequested();
                    int mb = (top + y) * columns + left + x;
                    primary.Read(bits, columns, mb, x, y, width, trim);
                    alpha?.Read(bits, columns, mb, x, y, width, trim);
                }
                bits.AlignZero(); left += width; tile++;
            }
            top += height;
        }
        return new ReconstructedFrame {
            Primary = primary.Result, Alpha = alpha?.Result,
            Highpass = primary.Highpass.Length == 0 && (alpha == null || alpha.Highpass.Length == 0) ? null : new BandData {
                Primary = primary.Highpass, Alpha = alpha?.Highpass ?? Array.Empty<int>(),
                ModelBits = primary.ModelBits, AlphaModelBits = alpha?.ModelBits ?? Array.Empty<byte>(),
                Patterns = primary.Patterns, AlphaPatterns = alpha?.Patterns ?? Array.Empty<int>()
            }
        };
    }
}
