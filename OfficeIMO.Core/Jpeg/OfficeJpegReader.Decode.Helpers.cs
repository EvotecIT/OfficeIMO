using System;
using System.Threading;

namespace OfficeIMO.Drawing;

internal static partial class OfficeJpegReader {
    private const int HuffmanFastBits = 9;
    private const string JpegDimensionsLimitMessage = "JPEG dimensions exceed limits.";
    private static readonly int[] CrToR = new int[256];
    private static readonly int[] CrToG = new int[256];
    private static readonly int[] CbToG = new int[256];
    private static readonly int[] CbToB = new int[256];

    static OfficeJpegReader() {
        for (var i = 0; i < 256; i++) {
            var d = i - 128;
            CrToR[i] = (91881 * d + 32768) >> 16;
            CrToG[i] = (46802 * d + 32768) >> 16;
            CbToG[i] = (22554 * d + 32768) >> 16;
            CbToB[i] = (116130 * d + 32768) >> 16;
        }
    }

    private static void DecodeBlock(
        ref JpegBitReader reader,
        HuffmanTable dcTable,
        HuffmanTable acTable,
        int[] quant,
        ref int prevDc,
        int[] coeffs,
        byte[] pixels,
        int[] workspace) {
        Array.Clear(coeffs, 0, 64);

        var t = DecodeHuffman(ref reader, dcTable, useFast: true);
        var diff = t == 0 ? 0 : Extend(reader.ReadBits(t), t);
        var dc = prevDc + diff;
        prevDc = dc;
        coeffs[0] = dc * quant[0];

        var k = 1;
        while (k < 64) {
            var rs = DecodeHuffman(ref reader, acTable, useFast: true);
            if (rs == 0) break;
            var r = rs >> 4;
            var s = rs & 0x0F;
            if (s == 0) {
                if (r == 15) {
                    k += 16;
                    continue;
                }
                break;
            }

            k += r;
            if (k >= 64) break;
            var ac = Extend(reader.ReadBits(s), s);
            var zig = ZigZag[k];
            coeffs[zig] = ac * quant[zig];
            k++;
        }

        InverseDct(coeffs, pixels, workspace);
    }

    private static int DecodeHuffman(ref JpegBitReader reader, HuffmanTable table, bool useFast) {
        if (useFast && table.Fast is not null && reader.TryPeekBits(HuffmanFastBits, out int peek)) {
            var entry = table.Fast[peek];
            if (entry >= 0) {
                var size = entry >> 8;
                reader.SkipBits(size);
                return entry & 0xFF;
            }
        }

        var node = 0;
        while (true) {
            var bit = reader.ReadBit();
            node = bit == 0 ? table.Left[node] : table.Right[node];
            if (node < 0) {
                if (reader.AllowTruncated) return 0;
                throw new FormatException("Invalid JPEG Huffman code.");
            }
            var symbol = table.Symbols[node];
            if (symbol >= 0) return symbol;
        }
    }

    private static int Extend(int value, int bits) {
        if (bits == 0) return 0;
        var limit = 1 << (bits - 1);
        if (value < limit) value -= (1 << bits) - 1;
        return value;
    }

    private static void WriteBlock(byte[] buffer, int stride, int blockX, int blockY, byte[] pixels) {
        var baseX = blockX * 8;
        var baseY = blockY * 8;
        for (var y = 0; y < 8; y++) {
            var row = (baseY + y) * stride + baseX;
            var src = y * 8;
            Buffer.BlockCopy(pixels, src, buffer, row, 8);
        }
    }

    private static byte ClampToByte(int value) {
        if (value <= 0) return 0;
        if (value >= 255) return 255;
        return (byte)value;
    }

    private static JpegFrame ParseFrameHeader(OfficeByteView data) {
        var precision = data[0];
        if (precision != 8) throw new FormatException("Unsupported JPEG precision.");
        var height = ReadUInt16BE(data, 1);
        var width = ReadUInt16BE(data, 3);
        var components = data[5];
        if (width == 0 || height == 0) throw new FormatException("Invalid JPEG dimensions.");
        if (!OfficeRasterGuards.TryEnsurePixelCount(width, height, out _)) {
            throw new FormatException(JpegDimensionsLimitMessage);
        }
        if (components < 1 || components > 4) {
            throw new FormatException("Unsupported JPEG component count.");
        }
        if (data.Length < 6 + components * 3) throw new FormatException("Invalid JPEG SOF segment.");

        var frame = new JpegFrame {
            Width = width,
            Height = height,
            ComponentCount = components,
            Components = new Component[components]
        };

        var offset = 6;
        var maxH = 0;
        var maxV = 0;
        var samplingUnits = 0;
        for (var i = 0; i < components; i++) {
            var id = data[offset++];
            var sampling = data[offset++];
            var h = sampling >> 4;
            var v = sampling & 0x0F;
            var qt = data[offset++];
            if (h == 0 || v == 0 || h > 4 || v > 4) throw new FormatException("Invalid JPEG sampling factors.");
            samplingUnits = checked(samplingUnits + h * v);
            if (samplingUnits > 10) throw new FormatException("JPEG sampling factors exceed supported limits.");
            if (qt >= 4) throw new FormatException("Unsupported JPEG quantization table.");
            frame.Components[i] = new Component {
                Id = id,
                H = h,
                V = v,
                QuantId = qt
            };
            if (h > maxH) maxH = h;
            if (v > maxV) maxV = v;
        }
        frame.MaxH = maxH;
        frame.MaxV = maxV;
        return frame;
    }

    internal static bool IsSupportedRgbaFrameHeader(byte[] data, int offset, int length) {
        try {
            JpegFrame frame = ParseFrameHeader(new OfficeByteView(data).Slice(offset, length));
            return frame.ComponentCount is 1 or 3 or 4;
        } catch (Exception ex) when (ex is FormatException || ex is ArgumentException ||
                                     ex is IndexOutOfRangeException || ex is OverflowException) {
            return false;
        }
    }

    private static ScanHeader ParseScanHeader(OfficeByteView data, ref JpegFrame frame) {
        var components = data[0];
        if (components == 0 || components > frame.ComponentCount) throw new FormatException("Invalid JPEG scan component count.");
        if (data.Length < 1 + components * 2 + 3) throw new FormatException("Invalid JPEG scan header.");

        var indices = new int[components];
        var seenComponents = new bool[frame.ComponentCount];
        var offset = 1;
        for (var i = 0; i < components; i++) {
            var id = data[offset++];
            var table = data[offset++];
            var dc = table >> 4;
            var ac = table & 0x0F;
            var index = FindComponentIndex(frame.Components, id);
            if (index < 0) throw new FormatException("Unknown JPEG component in scan.");
            if (seenComponents[index]) throw new FormatException("Duplicate JPEG component in scan.");
            seenComponents[index] = true;
            frame.Components[index].DcTable = (byte)dc;
            frame.Components[index].AcTable = (byte)ac;
            indices[i] = index;
        }

        var ss = data[offset++];
        var se = data[offset++];
        var ahal = data[offset++];

        return new ScanHeader {
            ComponentIndices = indices,
            Ss = ss,
            Se = se,
            Ah = (byte)(ahal >> 4),
            Al = (byte)(ahal & 0x0F)
        };
    }

    private static int FindComponentIndex(Component[] components, int id) {
        for (var i = 0; i < components.Length; i++) {
            if (components[i].Id == id) return i;
        }
        return -1;
    }

    private static int FindScanEnd(OfficeByteView data, int start, CancellationToken cancellationToken) {
        var i = start;
        while (i + 1 < data.Length) {
            if ((i & 0x3FFF) == 0) cancellationToken.ThrowIfCancellationRequested();
            if (data[i] == 0xFF) {
                var j = i + 1;
                SkipFillBytes(data, ref j, cancellationToken);
                if (j >= data.Length) return data.Length;
                var marker = data[j];
                if (marker == 0x00) {
                    i = j + 1;
                    continue;
                }
                if (marker >= 0xD0 && marker <= 0xD7) {
                    i = j + 1;
                    continue;
                }
                return i;
            }
            i++;
        }
        return data.Length;
    }

    private static bool TryReadAdobeTransform(OfficeByteView data, out int transform) {
        transform = 0;
        if (data.Length < 12) return false;
        if (data[0] != (byte)'A' || data[1] != (byte)'d' || data[2] != (byte)'o' || data[3] != (byte)'b' || data[4] != (byte)'e') {
            return false;
        }
        transform = data[11];
        return true;
    }

    private static byte[] ApplyOrientation(byte[] rgba, ref int width, ref int height,
        int orientation, CancellationToken cancellationToken) =>
        OfficeRasterOrientation.Apply(rgba, ref width, ref height, orientation, cancellationToken, JpegDimensionsLimitMessage);

    private static double[,] BuildCosTable() {
        var table = new double[8, 8];
        for (var x = 0; x < 8; x++) {
            for (var u = 0; u < 8; u++) {
                table[x, u] = Math.Cos(((2 * x + 1) * u * Math.PI) / 16.0);
            }
        }
        return table;
    }

    private static ushort ReadUInt16BE(OfficeByteView data, int offset) {
        return (ushort)((data[offset] << 8) | data[offset + 1]);
    }

    private struct Component {
        public byte Id;
        public int H;
        public int V;
        public byte QuantId;
        public byte DcTable;
        public byte AcTable;
    }

    private struct JpegFrame {
        public int Width;
        public int Height;
        public int ComponentCount;
        public Component[] Components;
        public int MaxH;
        public int MaxV;
    }

    private struct ScanHeader {
        public int[] ComponentIndices;
        public byte Ss;
        public byte Se;
        public byte Ah;
        public byte Al;
    }

    internal static bool TryInitializeDecodeWorkingSet(
        long retainedEncodedBytes,
        int width,
        int height,
        int orientation,
        out long reservedBytes) {
        reservedBytes = 0L;
        if (retainedEncodedBytes < 0L || width < 1 || height < 1 || orientation < 1 || orientation > 8) {
            return false;
        }
        try {
            long rgbaBytes = checked((long)width * height * 4L);
            reservedBytes = checked(
                retainedEncodedBytes + rgbaBytes * (orientation > 1 ? 2L : 1L) + 64L * 1024L);
            return reservedBytes <= OfficeRasterGuards.MaximumDecodedBytes;
        } catch (OverflowException) {
            reservedBytes = 0L;
            return false;
        }
    }

    internal static bool TryReserveOrientationCanvas(
        int width,
        int height,
        ref long reservedBytes,
        ref bool orientationCanvasReserved) {
        if (orientationCanvasReserved) return true;
        if (width < 1 || height < 1 || reservedBytes < 0L) return false;
        try {
            long rgbaBytes = checked((long)width * height * 4L);
            long updatedBytes = checked(reservedBytes + rgbaBytes);
            if (updatedBytes > OfficeRasterGuards.MaximumDecodedBytes) return false;
            reservedBytes = updatedBytes;
            orientationCanvasReserved = true;
            return true;
        } catch (OverflowException) {
            return false;
        }
    }

    private sealed class BaselineState {
        public BaselineComponentState[] Components = Array.Empty<BaselineComponentState>();
        public bool[] DecodedComponents = Array.Empty<bool>();
        public int McuCols;
        public int McuRows;
        private long _reservedBytes;
        private bool _orientationCanvasReserved;

        public static BaselineState Create(JpegFrame frame, int orientation, long retainedEncodedBytes) {
            var mcuWidth = frame.MaxH * 8;
            var mcuHeight = frame.MaxV * 8;
            var mcuCols = (frame.Width + mcuWidth - 1) / mcuWidth;
            var mcuRows = (frame.Height + mcuHeight - 1) / mcuHeight;
            var components = new BaselineComponentState[frame.ComponentCount];
            if (!TryInitializeDecodeWorkingSet(
                    retainedEncodedBytes, frame.Width, frame.Height, orientation, out long aggregateBytes)) {
                throw new FormatException(JpegDimensionsLimitMessage);
            }
            for (var i = 0; i < frame.ComponentCount; i++) {
                var component = frame.Components[i];
                var blocksPerRow = OfficeRasterGuards.EnsureByteCount((long)mcuCols * component.H, JpegDimensionsLimitMessage);
                var blocksPerCol = OfficeRasterGuards.EnsureByteCount((long)mcuRows * component.V, JpegDimensionsLimitMessage);
                components[i] = new BaselineComponentState(component, blocksPerRow, blocksPerCol, ref aggregateBytes);
            }

            return new BaselineState {
                Components = components,
                DecodedComponents = new bool[frame.ComponentCount],
                McuCols = mcuCols,
                McuRows = mcuRows,
                _reservedBytes = aggregateBytes,
                _orientationCanvasReserved = orientation > 1
            };
        }

        public void ReserveOrientationCanvas(JpegFrame frame) {
            if (!TryReserveOrientationCanvas(
                    frame.Width, frame.Height, ref _reservedBytes, ref _orientationCanvasReserved)) {
                throw new FormatException(JpegDimensionsLimitMessage);
            }
        }

        public byte[] RenderRgba(
            JpegFrame frame,
            int? adobeTransform,
            bool highQualityChroma,
            CancellationToken cancellationToken) {
            for (var i = 0; i < DecodedComponents.Length; i++) {
                if (!DecodedComponents[i]) throw new FormatException("Missing JPEG component scan.");
            }

            return ComposeRgba(frame, Components, adobeTransform, highQualityChroma, cancellationToken);
        }

        public byte[] RenderColorComponents(
            JpegFrame frame,
            int? adobeTransform,
            int? requestedColorTransform,
            bool usePdfColorTransformDefault,
            bool highQualityChroma,
            out int componentCount) {
            for (var i = 0; i < DecodedComponents.Length; i++) {
                if (!DecodedComponents[i]) throw new FormatException("Missing JPEG component scan.");
            }

            return ComposeColorComponents(
                frame,
                Components,
                adobeTransform,
                requestedColorTransform,
                usePdfColorTransformDefault,
                highQualityChroma,
                outputRgba: false,
                CancellationToken.None,
                out componentCount);
        }
    }

    private sealed class BaselineComponentState {
        public Component Component;
        public byte[] Buffer;
        public int[] BlockCoeffs;
        public byte[] BlockPixels;
        public int[] BlockWorkspace;
        public int Stride;
        public int BlocksPerRow;
        public int BlocksPerCol;
        public int PrevDc;

        public BaselineComponentState(Component component, int blocksPerRow, int blocksPerCol, ref long aggregateBytes) {
            Component = component;
            BlocksPerRow = blocksPerRow;
            BlocksPerCol = blocksPerCol;
            Stride = OfficeRasterGuards.EnsureByteCount((long)blocksPerRow * 8, JpegDimensionsLimitMessage);
            var bufferLength = OfficeRasterGuards.EnsureByteArrayLength((long)Stride * blocksPerCol * 8, ref aggregateBytes, JpegDimensionsLimitMessage);
            Buffer = new byte[bufferLength];
            BlockCoeffs = new int[OfficeRasterGuards.EnsureInt32ArrayLength(64, ref aggregateBytes, JpegDimensionsLimitMessage)];
            BlockPixels = new byte[OfficeRasterGuards.EnsureByteArrayLength(64, ref aggregateBytes, JpegDimensionsLimitMessage)];
            BlockWorkspace = new int[OfficeRasterGuards.EnsureInt32ArrayLength(64, ref aggregateBytes, JpegDimensionsLimitMessage)];
            PrevDc = 0;
        }

        public static BaselineComponentState FromDecodedBuffer(
            Component component,
            int blocksPerRow,
            int blocksPerCol,
            int stride,
            byte[] buffer) {
            return new BaselineComponentState {
                Component = component,
                BlocksPerRow = blocksPerRow,
                BlocksPerCol = blocksPerCol,
                Stride = stride,
                Buffer = buffer,
                BlockCoeffs = Array.Empty<int>(),
                BlockPixels = Array.Empty<byte>(),
                BlockWorkspace = Array.Empty<int>()
            };
        }

        private BaselineComponentState() {
            Buffer = Array.Empty<byte>();
            BlockCoeffs = Array.Empty<int>();
            BlockPixels = Array.Empty<byte>();
            BlockWorkspace = Array.Empty<int>();
        }
    }

    private sealed class ProgressiveState {
        public ProgressiveComponentState[] Components = Array.Empty<ProgressiveComponentState>();
        public int McuCols;
        public int McuRows;
        private long _reservedBytes;
        private bool _orientationCanvasReserved;

        public static ProgressiveState Create(
            JpegFrame frame,
            int[][] quantTables,
            int orientation,
            long retainedEncodedBytes) {
            var maxH = frame.MaxH;
            var maxV = frame.MaxV;
            var mcuWidth = maxH * 8;
            var mcuHeight = maxV * 8;
            var mcuCols = (frame.Width + mcuWidth - 1) / mcuWidth;
            var mcuRows = (frame.Height + mcuHeight - 1) / mcuHeight;

            var components = new ProgressiveComponentState[frame.ComponentCount];
            if (!TryInitializeDecodeWorkingSet(
                    retainedEncodedBytes, frame.Width, frame.Height, orientation, out long aggregateBytes)) {
                throw new FormatException(JpegDimensionsLimitMessage);
            }
            for (var i = 0; i < frame.ComponentCount; i++) {
                var comp = frame.Components[i];
                if (comp.QuantId >= quantTables.Length || quantTables[comp.QuantId] is null) {
                    throw new FormatException("Missing JPEG quantization table.");
                }
                var blocksPerRow = OfficeRasterGuards.EnsureByteCount((long)mcuCols * comp.H, JpegDimensionsLimitMessage);
                var blocksPerCol = OfficeRasterGuards.EnsureByteCount((long)mcuRows * comp.V, JpegDimensionsLimitMessage);
                components[i] = new ProgressiveComponentState(
                    comp,
                    blocksPerRow,
                    blocksPerCol,
                    quantTables[comp.QuantId],
                    ref aggregateBytes);
            }

            return new ProgressiveState {
                Components = components,
                McuCols = mcuCols,
                McuRows = mcuRows,
                _reservedBytes = aggregateBytes,
                _orientationCanvasReserved = orientation > 1
            };
        }

        public void ReserveOrientationCanvas(JpegFrame frame) {
            if (!TryReserveOrientationCanvas(
                    frame.Width, frame.Height, ref _reservedBytes, ref _orientationCanvasReserved)) {
                throw new FormatException(JpegDimensionsLimitMessage);
            }
        }

        public byte[] RenderRgba(
            JpegFrame frame,
            int? adobeTransform,
            bool highQualityChroma,
            CancellationToken cancellationToken) {
            BaselineComponentState[] baselineStates = CreateBaselineStates(cancellationToken);
            return ComposeRgba(frame, baselineStates, adobeTransform, highQualityChroma, cancellationToken);
        }

        public byte[] RenderColorComponents(
            JpegFrame frame,
            int? adobeTransform,
            int? requestedColorTransform,
            bool usePdfColorTransformDefault,
            bool highQualityChroma,
            out int componentCount) {
            BaselineComponentState[] baselineStates = CreateBaselineStates(CancellationToken.None);
            return ComposeColorComponents(
                frame,
                baselineStates,
                adobeTransform,
                requestedColorTransform,
                usePdfColorTransformDefault,
                highQualityChroma,
                outputRgba: false,
                CancellationToken.None,
                out componentCount);
        }

        private BaselineComponentState[] CreateBaselineStates(CancellationToken cancellationToken) {
            for (var i = 0; i < Components.Length; i++) {
                var compState = Components[i];
                for (var by = 0; by < compState.BlocksPerCol; by++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    for (var bx = 0; bx < compState.BlocksPerRow; bx++) {
                        var baseIndex = (by * compState.BlocksPerRow + bx) * 64;
                        for (int coefficient = 0; coefficient < 64; coefficient++) {
                            compState.BlockCoeffs[coefficient] =
                                compState.Coeffs[baseIndex + coefficient] * compState.Quantization[coefficient];
                        }
                        InverseDct(compState.BlockCoeffs, compState.BlockPixels, compState.BlockWorkspace);
                        WriteBlock(compState.Buffer, compState.Stride, bx, by, compState.BlockPixels);
                    }
                }
            }

            var baselineStates = new BaselineComponentState[Components.Length];
            for (var i = 0; i < Components.Length; i++) {
                var compState = Components[i];
                baselineStates[i] = BaselineComponentState.FromDecodedBuffer(
                    compState.Component,
                    compState.BlocksPerRow,
                    compState.BlocksPerCol,
                    compState.Stride,
                    compState.Buffer);
            }

            return baselineStates;
        }
    }

    private static byte ApplyCmyk(int c, int k) {
        var v = c + k;
        if (v > 255) v = 255;
        return (byte)(255 - v);
    }

    private sealed class ProgressiveComponentState {
        public Component Component;
        public int BlocksPerRow;
        public int BlocksPerCol;
        public short[] Coeffs;
        public int[] Quantization;
        public byte[] Buffer;
        public int[] BlockCoeffs;
        public byte[] BlockPixels;
        public int[] BlockWorkspace;
        public int Stride;
        public int PrevDc;

        public ProgressiveComponentState(
            Component component,
            int blocksPerRow,
            int blocksPerCol,
            int[] quantization,
            ref long aggregateBytes) {
            Component = component;
            BlocksPerRow = blocksPerRow;
            BlocksPerCol = blocksPerCol;
            Quantization = quantization;
            Stride = OfficeRasterGuards.EnsureByteCount((long)blocksPerRow * 8, JpegDimensionsLimitMessage);
            var coeffLength = OfficeRasterGuards.EnsureInt16ArrayLength((long)BlocksPerRow * BlocksPerCol * 64, ref aggregateBytes, JpegDimensionsLimitMessage);
            var bufferLength = OfficeRasterGuards.EnsureByteArrayLength((long)Stride * blocksPerCol * 8, ref aggregateBytes, JpegDimensionsLimitMessage);
            Coeffs = new short[coeffLength];
            Buffer = new byte[bufferLength];
            BlockCoeffs = new int[OfficeRasterGuards.EnsureInt32ArrayLength(64, ref aggregateBytes, JpegDimensionsLimitMessage)];
            BlockPixels = new byte[OfficeRasterGuards.EnsureByteArrayLength(64, ref aggregateBytes, JpegDimensionsLimitMessage)];
            BlockWorkspace = new int[OfficeRasterGuards.EnsureInt32ArrayLength(64, ref aggregateBytes, JpegDimensionsLimitMessage)];
            PrevDc = 0;
        }
    }

    private struct HuffmanTable {
        public int[] Left;
        public int[] Right;
        public int[] Symbols;
        public short[]? Fast;
        public bool IsValid;

        public static HuffmanTable Build(OfficeByteView counts, byte[] values) {
            var left = new int[512];
            var right = new int[512];
            var symbols = new int[512];
            var fast = new short[1 << HuffmanFastBits];
            for (var i = 0; i < left.Length; i++) left[i] = -1;
            for (var i = 0; i < right.Length; i++) right[i] = -1;
            for (var i = 0; i < symbols.Length; i++) symbols[i] = -1;
            for (var i = 0; i < fast.Length; i++) fast[i] = -1;

            var next = 1;
            var code = 0;
            var k = 0;
            for (var i = 1; i <= 16; i++) {
                var count = counts[i - 1];
                for (var j = 0; j < count; j++) {
                    var symbol = values[k++];
                    var node = 0;
                    for (var bit = i - 1; bit >= 0; bit--) {
                        var b = (code >> bit) & 1;
                        if (b == 0) {
                            if (left[node] < 0) {
                                if (next >= left.Length) throw new FormatException("Invalid JPEG Huffman tree.");
                                left[node] = next++;
                            }
                            node = left[node];
                        } else {
                            if (right[node] < 0) {
                                if (next >= right.Length) throw new FormatException("Invalid JPEG Huffman tree.");
                                right[node] = next++;
                            }
                            node = right[node];
                        }
                    }
                    symbols[node] = symbol;
                    if (i <= HuffmanFastBits) {
                        var fill = 1 << (HuffmanFastBits - i);
                        var start = code << (HuffmanFastBits - i);
                        var entry = (short)((i << 8) | symbol);
                        for (var f = 0; f < fill; f++) {
                            fast[start + f] = entry;
                        }
                    }
                    code++;
                }
                code <<= 1;
            }

            return new HuffmanTable {
                Left = left,
                Right = right,
                Symbols = symbols,
                Fast = fast,
                IsValid = true
            };
        }
    }

    private ref struct JpegBitReader {
        private readonly OfficeByteView _data;
        private readonly bool _allowTruncated;
        private readonly CancellationToken _cancellationToken;
        private int _pos;
        private int _bitBuffer;
        private int _bitCount;

        public bool RestartMarkerSeen;

        public JpegBitReader(
            OfficeByteView data,
            bool allowTruncated,
            CancellationToken cancellationToken) {
            _data = data;
            _allowTruncated = allowTruncated;
            _cancellationToken = cancellationToken;
            _pos = 0;
            _bitBuffer = 0;
            _bitCount = 0;
            RestartMarkerSeen = false;
        }

        public bool AllowTruncated => _allowTruncated;

        public bool TryPeekBits(int count, out int value) {
            int originalPosition = _pos;
            int originalBitBuffer = _bitBuffer;
            int originalBitCount = _bitCount;
            bool originalRestartMarkerSeen = RestartMarkerSeen;
            while (_bitCount < count) {
                if (!TryReadByte(out int next)) {
                    RestorePeekState(originalPosition, originalBitBuffer, originalBitCount, originalRestartMarkerSeen);
                    value = 0;
                    return false;
                }
                if (RestartMarkerSeen != originalRestartMarkerSeen) {
                    RestorePeekState(originalPosition, originalBitBuffer, originalBitCount, originalRestartMarkerSeen);
                    value = 0;
                    return false;
                }
                _bitBuffer = (_bitBuffer << 8) | next;
                _bitCount += 8;
            }
            value = (_bitBuffer >> (_bitCount - count)) & ((1 << count) - 1);
            return true;
        }

        private void RestorePeekState(int position, int bitBuffer, int bitCount, bool restartMarkerSeen) {
            _pos = position;
            _bitBuffer = bitBuffer;
            _bitCount = bitCount;
            RestartMarkerSeen = restartMarkerSeen;
        }

        public void SkipBits(int count) {
            if (count == 0) return;
            _bitCount -= count;
            if (_bitCount <= 0) {
                _bitCount = 0;
                _bitBuffer = 0;
            } else {
                _bitBuffer &= (1 << _bitCount) - 1;
            }
        }

        public int ReadBit() {
            EnsureBits(1);
            var bit = (_bitBuffer >> (_bitCount - 1)) & 1;
            _bitCount--;
            if (_bitCount == 0) {
                _bitBuffer = 0;
            } else {
                _bitBuffer &= (1 << _bitCount) - 1;
            }
            return bit;
        }

        public int ReadBits(int count) {
            if (count == 0) return 0;
            EnsureBits(count);
            var value = (_bitBuffer >> (_bitCount - count)) & ((1 << count) - 1);
            _bitCount -= count;
            if (_bitCount == 0) {
                _bitBuffer = 0;
            } else {
                _bitBuffer &= (1 << _bitCount) - 1;
            }
            return value;
        }


        public void ExpectRestartMarker() {
            _bitBuffer = 0;
            _bitCount = 0;
            while (_pos < _data.Length) {
                var b = _data[_pos++];
                if (b != 0xFF) continue;
                SkipFillBytes(_data, ref _pos, _cancellationToken);
                if (_pos >= _data.Length) throw new FormatException("Unexpected JPEG end.");
                var marker = _data[_pos++];
                if (marker >= 0xD0 && marker <= 0xD7) {
                    RestartMarkerSeen = false;
                    return;
                }
                if (marker == 0x00) continue;
                throw new FormatException("Unexpected JPEG marker in scan.");
            }
            throw new FormatException("Missing JPEG restart marker.");
        }

        private void EnsureBits(int count) {
            while (_bitCount < count) {
                var b = ReadByte();
                _bitBuffer = (_bitBuffer << 8) | b;
                _bitCount += 8;
            }
        }

        private int ReadByte() {
            if (TryReadByte(out int value)) return value;
            throw new FormatException("Unexpected JPEG end.");
        }

        private bool TryReadByte(out int value) {
            while (_pos < _data.Length) {
                var b = _data[_pos++];
                if (b != 0xFF) {
                    value = b;
                    return true;
                }
                SkipFillBytes(_data, ref _pos, _cancellationToken);
                if (_pos >= _data.Length) {
                    if (_allowTruncated) {
                        value = 0;
                        return true;
                    }
                    value = 0;
                    return false;
                }
                var marker = _data[_pos++];
                if (marker == 0x00) {
                    value = 0xFF;
                    return true;
                }
                if (marker >= 0xD0 && marker <= 0xD7) {
                    RestartMarkerSeen = true;
                    continue;
                }
                throw new FormatException("Unexpected JPEG marker in scan.");
            }
            if (_allowTruncated) {
                value = 0;
                return true;
            }
            value = 0;
            return false;
        }
    }

}
