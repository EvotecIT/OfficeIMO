namespace OfficeIMO.Drawing;

internal static partial class OfficeJpegReader {
    private struct Component {
        public byte Id;
        public int H;
        public int V;
        public byte QuantId;
        public byte DcTable;
        public byte AcTable;
    }

    private struct JpegFrame {
        public int Precision;
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
        out long reservedBytes,
        int outputComponents = 4, int outputSampleBytes = 1) {
        reservedBytes = 0L;
        if (retainedEncodedBytes < 0L || width < 1 || height < 1 || orientation < 1 || orientation > 8 || outputComponents < 1 || outputComponents > 255 || outputSampleBytes < 1 || outputSampleBytes > 2) {
            return false;
        }
        try {
            long rgbaBytes = checked((long)width * height * 4L);
            reservedBytes = checked(
                retainedEncodedBytes + checked((long)width * height * outputComponents * outputSampleBytes) +
                (orientation > 1 ? rgbaBytes : 0L) + 64L * 1024L);
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

        public static BaselineState Create(JpegFrame frame, int orientation, long retainedEncodedBytes, bool preserveRaw16 = false) {
            var mcuWidth = frame.MaxH * 8;
            var mcuHeight = frame.MaxV * 8;
            var mcuCols = (frame.Width + mcuWidth - 1) / mcuWidth;
            var mcuRows = (frame.Height + mcuHeight - 1) / mcuHeight;
            var components = new BaselineComponentState[frame.ComponentCount];
            if (!TryInitializeDecodeWorkingSet(
                    retainedEncodedBytes, frame.Width, frame.Height, orientation, out long aggregateBytes, Math.Max(4, frame.ComponentCount), preserveRaw16 ? 2 : 1)) {
                throw new FormatException(JpegDimensionsLimitMessage);
            }
            for (var i = 0; i < frame.ComponentCount; i++) {
                var component = frame.Components[i];
                var blocksPerRow = OfficeRasterGuards.EnsureByteCount((long)mcuCols * component.H, JpegDimensionsLimitMessage);
                var blocksPerCol = OfficeRasterGuards.EnsureByteCount((long)mcuRows * component.V, JpegDimensionsLimitMessage);
                components[i] = new BaselineComponentState(component, blocksPerRow, blocksPerCol, ref aggregateBytes, frame.Precision);
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

        public void PrepareAcLookup(ref HuffmanTable table) =>
            table.PrepareBaselineAcLookup(ref _reservedBytes);

        // Decode budget not yet reserved, for optional rendering tables.
        public long AvailableBytes => OfficeRasterGuards.MaximumDecodedBytes - _reservedBytes;

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

            return ComposeRgba(frame, Components, adobeTransform, highQualityChroma, AvailableBytes, cancellationToken);
        }

        public byte[] RenderColorComponents(
            JpegFrame frame,
            int? adobeTransform,
            int? requestedColorTransform,
            bool usePdfColorTransformDefault,
            bool highQualityChroma,
            CancellationToken cancellationToken,
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
                AvailableBytes,
                cancellationToken,
                out componentCount);
        }
    }

    private sealed class BaselineComponentState {
        public Component Component;
        public byte[] Buffer;
        public ushort[]? WideBuffer;
        public int SampleCount => WideBuffer?.Length ?? Buffer.Length;
        public int ReadSample(int index) => WideBuffer == null ? Buffer[index] : WideBuffer[index];
        public void WriteSample(int index, int value) {
            if (WideBuffer == null) Buffer[index] = (byte)value; else WideBuffer[index] = (ushort)value;
        }
        public void CopySamples(int from, int to, int count) {
            if (WideBuffer == null) Array.Copy(Buffer, from, Buffer, to, count);
            else Array.Copy(WideBuffer, from, WideBuffer, to, count);
        }
        public int[] BlockCoeffs;
        public byte[] BlockPixels;
        public int[] BlockWorkspace;
        public int Stride;
        public int BlocksPerRow;
        public int BlocksPerCol;
        public int PrevDc;
        public int SampleMaximum = 255;
        public double[] WideWorkspace;

        private readonly int _planeLength;

        // Progressive decoding reserves the plane with the rest of its working set but
        // allocates it only when the image is rendered through component planes.
        public BaselineComponentState(Component component, int blocksPerRow, int blocksPerCol, ref long aggregateBytes, int precision = 8,
            bool allocatePlane = true) {
            Component = component;
            BlocksPerRow = blocksPerRow;
            BlocksPerCol = blocksPerCol;
            Stride = OfficeRasterGuards.EnsureByteCount((long)blocksPerRow * 8, JpegDimensionsLimitMessage);
            SampleMaximum = (1 << precision) - 1;
            int sampleBytes = precision > 8 ? 2 : 1;
            _planeLength = OfficeRasterGuards.EnsureByteArrayLength((long)Stride * blocksPerCol * 8 * sampleBytes, ref aggregateBytes, JpegDimensionsLimitMessage);
            Buffer = Array.Empty<byte>();
            if (allocatePlane) AllocatePlane();
            BlockCoeffs = new int[OfficeRasterGuards.EnsureInt32ArrayLength(64, ref aggregateBytes, JpegDimensionsLimitMessage)];
            BlockPixels = new byte[OfficeRasterGuards.EnsureByteArrayLength(64, ref aggregateBytes, JpegDimensionsLimitMessage)];
            BlockWorkspace = new int[OfficeRasterGuards.EnsureInt32ArrayLength(64, ref aggregateBytes, JpegDimensionsLimitMessage)];
            WideWorkspace = precision == 12
                ? new double[OfficeRasterGuards.EnsureByteArrayLength(64 * sizeof(double), ref aggregateBytes, JpegDimensionsLimitMessage) / sizeof(double)]
                : Array.Empty<double>();
            PrevDc = 0;
        }

        public void AllocatePlane() {
            if (Buffer.Length != 0 || WideBuffer != null) return;
            if (SampleMaximum > 255) WideBuffer = new ushort[_planeLength / 2];
            else Buffer = new byte[_planeLength];
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
            long retainedEncodedBytes, bool arithmeticCoding = false) {
            // Arithmetic scans may define quantization tables late and never use the
            // Huffman refinement masks.
            var maxH = frame.MaxH;
            var maxV = frame.MaxV;
            var mcuWidth = maxH * 8;
            var mcuHeight = maxV * 8;
            var mcuCols = (frame.Width + mcuWidth - 1) / mcuWidth;
            var mcuRows = (frame.Height + mcuHeight - 1) / mcuHeight;

            var components = new ProgressiveComponentState[frame.ComponentCount];
            if (!TryInitializeDecodeWorkingSet(
                    retainedEncodedBytes, frame.Width, frame.Height, orientation, out long aggregateBytes, Math.Max(4, frame.ComponentCount))) {
                throw new FormatException(JpegDimensionsLimitMessage);
            }
            try {
                for (var i = 0; i < frame.ComponentCount; i++) {
                    var comp = frame.Components[i];
                    if (comp.QuantId >= quantTables.Length || !arithmeticCoding && quantTables[comp.QuantId] is null) {
                        throw new FormatException("Missing JPEG quantization table.");
                    }
                    var blocksPerRow = OfficeRasterGuards.EnsureByteCount((long)mcuCols * comp.H, JpegDimensionsLimitMessage);
                    var blocksPerCol = OfficeRasterGuards.EnsureByteCount((long)mcuRows * comp.V, JpegDimensionsLimitMessage);
                    components[i] = new ProgressiveComponentState(
                        comp,
                        blocksPerRow,
                        blocksPerCol,
                        quantTables[comp.QuantId] ?? Array.Empty<int>(),
                        ref aggregateBytes, frame.Precision, trackNonzeroPositions: !arithmeticCoding);
                }
                // Exact storage for every component is reserved above; pool rounding may use
                // only what is left after room for a later orientation canvas and lookups.
                long headroom = (orientation > 1 ? 0L : (long)frame.Width * frame.Height * 4L) + 64L * 1024L;
                foreach (ProgressiveComponentState component in components) component.AllocateBuffers(ref aggregateBytes, headroom);
            } catch {
                foreach (ProgressiveComponentState? created in components) created?.ReturnBuffers();
                throw;
            }

            return new ProgressiveState {
                Components = components,
                McuCols = mcuCols,
                McuRows = mcuRows,
                _reservedBytes = aggregateBytes,
                _orientationCanvasReserved = orientation > 1
            };
        }

        public void PrepareAcLookup(ref HuffmanTable table) =>
            table.PrepareBaselineAcLookup(ref _reservedBytes);

        // Decode budget not yet reserved, for optional rendering tables.
        public long AvailableBytes => OfficeRasterGuards.MaximumDecodedBytes - _reservedBytes;

        public void ReserveOrientationCanvas(JpegFrame frame) {
            if (!TryReserveOrientationCanvas(
                    frame.Width, frame.Height, ref _reservedBytes, ref _orientationCanvasReserved)) {
                throw new FormatException(JpegDimensionsLimitMessage);
            }
        }

        // The decoder calls this once the image is rendered or decoding failed.
        public void ReturnBuffers() {
            foreach (ProgressiveComponentState component in Components) component.ReturnBuffers();
        }

        public byte[] RenderRgba(
            JpegFrame frame,
            int? adobeTransform,
            bool highQualityChroma,
            CancellationToken cancellationToken) {
            if (highQualityChroma && TryRenderYccHighQualityBands(frame, this, adobeTransform, cancellationToken, out byte[] banded)) {
                return banded;
            }
            BaselineComponentState[] baselineStates = CreateBaselineStates(cancellationToken);
            return ComposeRgba(frame, baselineStates, adobeTransform, highQualityChroma, AvailableBytes, cancellationToken);
        }

        public byte[] RenderColorComponents(
            JpegFrame frame,
            int? adobeTransform,
            int? requestedColorTransform,
            bool usePdfColorTransformDefault,
            bool highQualityChroma,
            CancellationToken cancellationToken,
            out int componentCount) {
            BaselineComponentState[] baselineStates = CreateBaselineStates(cancellationToken);
            return ComposeColorComponents(
                frame,
                baselineStates,
                adobeTransform,
                requestedColorTransform,
                usePdfColorTransformDefault,
                highQualityChroma,
                outputRgba: false,
                AvailableBytes,
                cancellationToken,
                out componentCount);
        }

        private BaselineComponentState[] CreateBaselineStates(CancellationToken cancellationToken) {
            for (var i = 0; i < Components.Length; i++) {
                var compState = Components[i];
                var samples = compState.Samples;
                samples.AllocatePlane();
                // Sixteen-bit coefficients times quantizers in [0, 65535] cannot overflow Int32.
                bool boundedQuantization = HasBoundedQuantization(compState.Quantization);
                for (var by = 0; by < compState.BlocksPerCol; by++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    for (var bx = 0; bx < compState.BlocksPerRow; bx++) {
                        var baseIndex = (by * compState.BlocksPerRow + bx) * 64;
                        if (boundedQuantization) {
                            Dequantize(compState.Coeffs, baseIndex, compState.Quantization, samples.BlockCoeffs);
                        } else {
                            for (int coefficient = 0; coefficient < 64; coefficient++) {
                                samples.BlockCoeffs[coefficient] =
                                    checked(compState.Coeffs[baseIndex + coefficient] * compState.Quantization[coefficient]);
                            }
                        }
                        WriteDctBlock(samples, bx, by);
                    }
                }
            }

            var baselineStates = new BaselineComponentState[Components.Length];
            for (var i = 0; i < Components.Length; i++) {
                var compState = Components[i];
                baselineStates[i] = compState.Samples;
            }

            return baselineStates;
        }
    }

    private sealed class ProgressiveComponentState {
        public Component Component;
        public int BlocksPerRow;
        public int BlocksPerCol;
        public short[] Coeffs;
        // Per block, bit j set when the coefficient at zigzag position j is nonzero;
        // maintained by the Huffman AC scans for refinement.
        public ulong[] NonzeroPositions;
        public int[] Quantization;
        public BaselineComponentState Samples;
        public int PrevDc;

        public ProgressiveComponentState(
            Component component,
            int blocksPerRow,
            int blocksPerCol,
            int[] quantization,
            ref long aggregateBytes, int precision, bool trackNonzeroPositions) {
            Component = component;
            BlocksPerRow = blocksPerRow;
            BlocksPerCol = blocksPerCol;
            Quantization = quantization;
            var coeffLength = OfficeRasterGuards.EnsureInt16ArrayLength((long)BlocksPerRow * BlocksPerCol * 64, ref aggregateBytes, JpegDimensionsLimitMessage);
            int maskLength = trackNonzeroPositions
                ? OfficeRasterGuards.EnsureByteArrayLength(coeffLength / 64 * 8L, ref aggregateBytes, JpegDimensionsLimitMessage) / 8
                : 0;
            Samples = new BaselineComponentState(component, blocksPerRow, blocksPerCol, ref aggregateBytes, precision, allocatePlane: false);
            _coefficientLength = coeffLength;
            _maskLength = maskLength;
            Coeffs = Array.Empty<short>();
            NonzeroPositions = Array.Empty<ulong>();
            PrevDc = 0;
        }

        private readonly int _coefficientLength;
        private readonly int _maskLength;

        // Called once every component's exact storage is reserved.
        public void AllocateBuffers(ref long aggregateBytes, long headroom) {
            Coeffs = RentCleared<short>(_coefficientLength, sizeof(short), ref aggregateBytes, headroom);
            NonzeroPositions = RentCleared<ulong>(_maskLength, sizeof(ulong), ref aggregateBytes, headroom);
        }

        // Coefficient planes are as large as the image and needed only while one image
        // decodes, so they come from the shared pool. The pool rounds lengths up to a power
        // of two; that extra capacity is charged to the decode budget, and a plane whose
        // rounding does not fit (keeping headroom free) is allocated at its exact length
        // instead. Only the requested length is cleared and used.
        private static T[] RentCleared<T>(int length, int elementSize, ref long aggregateBytes, long headroom) {
            if (length == 0) return Array.Empty<T>();
#if NET8_0_OR_GREATER
            if (length <= 1 << 30) {
                long capacity = Math.Max(16L, (long)System.Numerics.BitOperations.RoundUpToPowerOf2((uint)length));
                long rounded = checked(aggregateBytes + (capacity - length) * elementSize);
                if (rounded + headroom <= OfficeRasterGuards.MaximumDecodedBytes) {
                    T[] array = System.Buffers.ArrayPool<T>.Shared.Rent(length);
                    if (array.Length <= capacity) {
                        aggregateBytes = rounded;
                        Array.Clear(array, 0, length);
                        return array;
                    }
                    System.Buffers.ArrayPool<T>.Shared.Return(array);
                }
            }
#endif
            return new T[length];
        }

        public void ReturnBuffers() {
#if NET8_0_OR_GREATER
            // Exact-length fallbacks are left to the collector.
            if (IsPooledLength(Coeffs.Length)) System.Buffers.ArrayPool<short>.Shared.Return(Coeffs);
            if (IsPooledLength(NonzeroPositions.Length)) System.Buffers.ArrayPool<ulong>.Shared.Return(NonzeroPositions);
#endif
            Coeffs = Array.Empty<short>();
            NonzeroPositions = Array.Empty<ulong>();
        }

        private static bool IsPooledLength(int length) => length >= 16 && (length & (length - 1)) == 0;
    }

}
