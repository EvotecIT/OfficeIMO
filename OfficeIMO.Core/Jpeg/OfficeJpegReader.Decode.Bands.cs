namespace OfficeIMO.Drawing;

internal static partial class OfficeJpegReader {
    // Progressive eight-bit YCbCr with high-quality chroma renders one MCU row (band) at
    // a time: each band's blocks are dequantized and transformed into small rotating
    // buffers and converted straight into the RGBA output, so no full-frame component
    // planes are allocated. Chroma interpolation reads at most one row from each
    // neighboring band, which stays resident in the three-band ring. Pixels are the
    // ones ComposeColorComponents produces from complete planes.
    private static bool TryRenderYccHighQualityBands(
        JpegFrame frame,
        ProgressiveState state,
        int? adobeTransform,
        CancellationToken cancellationToken,
        out byte[] rgba) {
        rgba = Array.Empty<byte>();
        if (frame.Precision != 8 || frame.ComponentCount != 3 || state.Components.Length != 3) return false;
        var components = frame.Components;
        bool hasRgbComponentIds = FindComponentIndex(components, 'R') >= 0 &&
            FindComponentIndex(components, 'G') >= 0 && FindComponentIndex(components, 'B') >= 0;
        bool transformToRgb = adobeTransform.HasValue ? adobeTransform.Value == 1 : !hasRgbComponentIds;
        if (!transformToRgb) return false;
        int lumaIndex = FindComponentIndex(components, 1);
        int cbIndex = FindComponentIndex(components, 2);
        int crIndex = FindComponentIndex(components, 3);
        if (lumaIndex < 0 || cbIndex < 0 || crIndex < 0) {
            lumaIndex = 0;
            cbIndex = 1;
            crIndex = 2;
        }

        ProgressiveComponentState luma = state.Components[lumaIndex];
        ProgressiveComponentState cb = state.Components[cbIndex];
        ProgressiveComponentState cr = state.Components[crIndex];
        int maxH = frame.MaxH;
        int maxV = frame.MaxV;
        int bands = state.McuRows;
        int workers = (long)frame.Width * frame.Height >= 262_144 && frame.Height >= 16 && Environment.ProcessorCount > 1
            ? Math.Min(Math.Min(8, Environment.ProcessorCount), bands)
            : 1;
        // Band rings and axis tables replace the component planes the state reserved but
        // never allocated; together they must fit that reservation plus the unreserved budget.
        long planeBytes = 0L;
        long bandBytes = 0L;
        foreach (ProgressiveComponentState component in state.Components) {
            long componentBandBytes = (long)component.Samples.Stride * component.Component.V * 8L;
            planeBytes += componentBandBytes * component.BlocksPerCol / component.Component.V;
            bandBytes += 3L * (componentBandBytes + 24L + 4L * 64L * 3L);
        }
        if (workers * bandBytes + GetChromaAxisBytes(frame.Width, frame.Height) > planeBytes + state.AvailableBytes) {
            return false;
        }

        // Bounded tables keep dequantization exception-free inside the parallel bands.
        if (!HasBoundedQuantization(luma.Quantization) || !HasBoundedQuantization(cb.Quantization) ||
            !HasBoundedQuantization(cr.Quantization) ||
            luma.Component.H != maxH || luma.Component.V != maxV ||
            !TryCreateChromaAxis(cb.Component.H, maxH, frame.Width, out ChromaAxis cbColumns) ||
            !TryCreateChromaAxis(cr.Component.H, maxH, frame.Width, out ChromaAxis crColumns) ||
            !TryCreateChromaAxis(cb.Component.V, maxV, frame.Height, out ChromaAxis cbRows) ||
            !TryCreateChromaAxis(cr.Component.V, maxV, frame.Height, out ChromaAxis crRows)) {
            return false;
        }

        byte[] output = OfficeRasterGuards.AllocateRgba32(frame.Width, frame.Height, JpegDimensionsLimitMessage);
        var plan = new BandPlan(frame, state, output, luma, cb, cr, cbColumns, crColumns, cbRows, crRows);
        if (workers <= 1) {
            plan.RenderBands(0, bands, cancellationToken);
        } else {
            Parallel.For(0, workers, new ParallelOptions {
                CancellationToken = cancellationToken,
                MaxDegreeOfParallelism = workers
            }, worker => plan.RenderBands(
                (int)((long)bands * worker / workers),
                (int)((long)bands * (worker + 1) / workers),
                cancellationToken));
        }
        rgba = output;
        return true;
    }

    private sealed class BandPlan {
        private readonly JpegFrame _frame;
        private readonly ProgressiveState _state;
        private readonly byte[] _output;
        private readonly ProgressiveComponentState _luma;
        private readonly ProgressiveComponentState _cb;
        private readonly ProgressiveComponentState _cr;
        private readonly ChromaAxis _cbColumns;
        private readonly ChromaAxis _crColumns;
        private readonly ChromaAxis _cbRows;
        private readonly ChromaAxis _crRows;

        public BandPlan(JpegFrame frame, ProgressiveState state, byte[] output,
            ProgressiveComponentState luma, ProgressiveComponentState cb, ProgressiveComponentState cr,
            ChromaAxis cbColumns, ChromaAxis crColumns, ChromaAxis cbRows, ChromaAxis crRows) {
            _frame = frame;
            _state = state;
            _output = output;
            _luma = luma;
            _cb = cb;
            _cr = cr;
            _cbColumns = cbColumns;
            _crColumns = crColumns;
            _cbRows = cbRows;
            _crRows = crRows;
        }

        // Renders bands [firstBand, endBand) with worker-local band buffers. Bands just
        // outside the range are transformed again here as chroma context.
        public void RenderBands(int firstBand, int endBand, CancellationToken cancellationToken) {
            if (firstBand >= endBand) return;
            var lumaRing = new BandRing(_luma);
            var cbRing = new BandRing(_cb);
            var crRing = new BandRing(_cr);
            int bandHeight = _frame.MaxV * 8;
            int width = _frame.Width;
            for (int band = firstBand; band < endBand; band++) {
                cancellationToken.ThrowIfCancellationRequested();
                for (int neighbor = band - 1; neighbor <= band + 1; neighbor++) {
                    if (neighbor < 0 || neighbor >= _state.McuRows) continue;
                    lumaRing.Load(neighbor);
                    cbRing.Load(neighbor);
                    crRing.Load(neighbor);
                }
                int firstRow = band * bandHeight;
                int endRow = Math.Min(_frame.Height, firstRow + bandHeight);
                BaselineComponentState lumaBand = lumaRing.Get(band);
                for (int y = firstRow; y < endRow; y++) {
                    cbRing.Locate(_cbRows.First[y], out byte[] cbTop, out int cbTopOffset);
                    cbRing.Locate(_cbRows.Second[y], out byte[] cbBottom, out int cbBottomOffset);
                    crRing.Locate(_crRows.First[y], out byte[] crTop, out int crTopOffset);
                    crRing.Locate(_crRows.Second[y], out byte[] crBottom, out int crBottomOffset);
                    ComposeYccRow(_output, y * width * 4, width, outputRgba: true,
                        lumaBand.Buffer, (y - firstRow) * lumaBand.Stride,
                        cbTop, cbTopOffset, cbBottom, cbBottomOffset, _cbRows.Weight[y], _cbColumns,
                        crTop, crTopOffset, crBottom, crBottomOffset, _crRows.Weight[y], _crColumns);
                }
            }
        }
    }

    // Three one-band sample buffers for a component, slot = band % 3.
    private sealed class BandRing {
        private readonly ProgressiveComponentState _component;
        private readonly BaselineComponentState[] _slots = new BaselineComponentState[3];
        private readonly int[] _bands = { -1, -1, -1 };
        private readonly int _bandRows;

        public BandRing(ProgressiveComponentState component) {
            _component = component;
            _bandRows = component.Component.V * 8;
            long unusedBudget = 0L;
            for (int slot = 0; slot < _slots.Length; slot++) {
                _slots[slot] = new BaselineComponentState(component.Component, component.BlocksPerRow,
                    component.Component.V, ref unusedBudget);
            }
        }

        public BaselineComponentState Get(int band) => _slots[band % 3];

        public void Load(int band) {
            int slot = band % 3;
            if (_bands[slot] == band) return;
            BaselineComponentState samples = _slots[slot];
            int verticalBlocks = _component.Component.V;
            for (int row = 0; row < verticalBlocks; row++) {
                int blockY = band * verticalBlocks + row;
                for (int blockX = 0; blockX < _component.BlocksPerRow; blockX++) {
                    int baseIndex = (blockY * _component.BlocksPerRow + blockX) * 64;
                    Dequantize(_component.Coeffs, baseIndex, _component.Quantization, samples.BlockCoeffs);
                    WriteDctBlock(samples, blockX, row);
                }
            }
            _bands[slot] = band;
        }

        // Component sample row to its band buffer and row offset.
        public void Locate(int row, out byte[] samples, out int offset) {
            int band = row / _bandRows;
            BaselineComponentState state = _slots[band % 3];
            samples = state.Buffer;
            offset = (row - band * _bandRows) * state.Stride;
        }
    }
}
