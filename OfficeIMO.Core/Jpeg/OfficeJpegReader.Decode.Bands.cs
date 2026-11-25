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
        if (!TryCreateBandPlan(frame, state, adobeTransform, out BandPlan? plan, out int workers) || plan == null) return false;
        int bands = state.McuRows;
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
        rgba = plan.Output;
        return true;
    }

    private static bool TryCreateBandPlan(
        JpegFrame frame,
        ProgressiveState state,
        int? adobeTransform,
        out BandPlan? plan,
        out int workers) {
        plan = null;
        workers = 1;
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
        workers = (long)frame.Width * frame.Height >= 262_144 && frame.Height >= 16 && Environment.ProcessorCount > 1
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

        // Every band writes all of its rows, so the output needs no clearing.
        byte[] output = OfficeRasterGuards.AllocateRgba32Uninitialized(frame.Width, frame.Height, JpegDimensionsLimitMessage);
        plan = new BandPlan(frame, state, output, luma, cb, cr, cbColumns, crColumns, cbRows, crRows);
        return true;
    }

    // Renders the bands while the final progressive scan is still decoding. The scan reports
    // how many of its MCU rows are complete; a band is rendered once every row it reads,
    // including the neighboring bands used as chroma context, has its final coefficients.
    // Workers take chunks of bands in order, and the decoding thread joins in once the scan
    // ends. The coefficient planes are pooled, so the decoder waits for every worker before
    // it returns them, whether the scan finished or failed.
    private sealed class BandOverlap {
        private const int ChunkBands = 4;
        private readonly BandPlan _plan;
        private readonly int _bands;
        private readonly int _chunks;
        private readonly int _scanRows;
        // Scan rows a band's last neighbor needs: bands * this, capped at the scan's rows.
        private readonly int _rowsPerBand;
        private readonly CancellationToken _cancellationToken;
        private readonly object _gate = new object();
        private readonly Task[] _tasks;
        private int _nextChunk = -1;
        private int _completedRows;
        // The fewest completed rows any waiting worker needs; the scan takes the lock only then.
        private int _wakeAt = int.MaxValue;
        private bool _stopped;

        private BandOverlap(BandPlan plan, int bands, int scanRows, int rowsPerBand, int workers,
            CancellationToken cancellationToken) {
            _plan = plan;
            _bands = bands;
            _chunks = (bands + ChunkBands - 1) / ChunkBands;
            _scanRows = scanRows;
            _rowsPerBand = rowsPerBand;
            _cancellationToken = cancellationToken;
            _tasks = new Task[Math.Min(workers - 1, _chunks)];
            for (int i = 0; i < _tasks.Length; i++) {
                _tasks[i] = Task.Factory.StartNew(() => RenderChunks(waitForRows: true), CancellationToken.None,
                    TaskCreationOptions.DenyChildAttach, TaskScheduler.Default);
            }
        }

        public byte[] Output => _plan.Output;

        // Starts the workers when the banded render applies and at least one other worker is available.
        public static BandOverlap? TryStart(JpegFrame frame, ProgressiveState state, int? adobeTransform,
            ScanHeader scan, int scanRows, CancellationToken cancellationToken) {
            if (!TryCreateBandPlan(frame, state, adobeTransform, out BandPlan? plan, out int workers) ||
                plan == null || workers <= 1) return null;
            // An interleaved scan row is one band of every component; a single-component row is one block row.
            int rowsPerBand = scan.ComponentIndices.Length == 1 ? frame.Components[scan.ComponentIndices[0]].V : 1;
            return new BandOverlap(plan, state.McuRows, scanRows, rowsPerBand, workers, cancellationToken);
        }

        // Called by the scan after each of its MCU rows; the coefficients of those rows are final.
        // Both sides publish with a full fence before reading the other's field, so a worker that
        // registers its row count either sees these rows or is woken for them.
        public void Publish(int completedRows) {
            Interlocked.Exchange(ref _completedRows, completedRows);
            if (completedRows < Volatile.Read(ref _wakeAt)) return;
            lock (_gate) {
                _wakeAt = int.MaxValue;
                Monitor.PulseAll(_gate);
            }
        }

        // The scan has ended: render what is left on this thread too, then wait for the workers.
        public void Finish() {
            Publish(int.MaxValue);
            try {
                RenderChunks(waitForRows: false);
            } finally {
                WaitForWorkers(observeFailures: true);
            }
        }

        // The scan failed: stop handing out chunks and wait for the workers before the buffers are returned.
        public void Stop() {
            lock (_gate) {
                _stopped = true;
                Monitor.PulseAll(_gate);
            }
            WaitForWorkers(observeFailures: false);
        }

        private void WaitForWorkers(bool observeFailures) {
            List<Exception>? failures = null;
            foreach (Task task in _tasks) {
                try {
                    task.Wait();
                } catch (AggregateException exception) {
                    if (observeFailures) (failures ??= new List<Exception>()).AddRange(exception.InnerExceptions);
                }
            }
            if (failures == null) return;
            _cancellationToken.ThrowIfCancellationRequested();
            if (failures.Count == 1) System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(failures[0]).Throw();
            throw new AggregateException(failures);
        }

        private void RenderChunks(bool waitForRows) {
            BandRing[]? rings = null;
            while (true) {
                int chunk = Interlocked.Increment(ref _nextChunk);
                if (chunk >= _chunks) return;
                int firstBand = chunk * ChunkBands;
                int endBand = Math.Min(_bands, firstBand + ChunkBands);
                if (waitForRows && !WaitForBand(endBand - 1)) return;
                _plan.RenderBands(firstBand, endBand, rings ??= _plan.CreateRings(), _cancellationToken);
            }
        }

        private bool WaitForBand(int band) {
            int needed = (int)Math.Min((long)(band + 2) * _rowsPerBand, _scanRows);
            if (Volatile.Read(ref _completedRows) >= needed) return true;
            lock (_gate) {
                while (!_stopped) {
                    int wakeAt;
                    do {
                        wakeAt = _wakeAt;
                        if (wakeAt <= needed) break;
                    } while (Interlocked.CompareExchange(ref _wakeAt, needed, wakeAt) != wakeAt);
                    if (Volatile.Read(ref _completedRows) >= needed) return true;
                    _cancellationToken.ThrowIfCancellationRequested();
                    Monitor.Wait(_gate, 50);
                }
                return false;
            }
        }
    }

    private sealed class BandPlan {
        private readonly JpegFrame _frame;
        private readonly ProgressiveState _state;
        private readonly byte[] _output;
        public byte[] Output => _output;
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
        public void RenderBands(int firstBand, int endBand, CancellationToken cancellationToken) =>
            RenderBands(firstBand, endBand, CreateRings(), cancellationToken);

        public BandRing[] CreateRings() => new[] { new BandRing(_luma), new BandRing(_cb), new BandRing(_cr) };

        // A worker can pass the same rings for several ranges: a slot is reused only for the band it
        // already holds, and bands are rendered only once their coefficients are final.
        public void RenderBands(int firstBand, int endBand, BandRing[] rings, CancellationToken cancellationToken) {
            if (firstBand >= endBand) return;
            BandRing lumaRing = rings[0];
            BandRing cbRing = rings[1];
            BandRing crRing = rings[2];
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
                        new ChromaRows(cbTop, cbTopOffset, cbBottom, cbBottomOffset, _cbRows.Weight[y], _cbColumns),
                        new ChromaRows(crTop, crTopOffset, crBottom, crBottomOffset, _crRows.Weight[y], _crColumns));
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
