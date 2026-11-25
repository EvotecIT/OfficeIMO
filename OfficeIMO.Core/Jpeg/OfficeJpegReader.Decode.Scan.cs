namespace OfficeIMO.Drawing;

internal static partial class OfficeJpegReader {
    private static void DecodeBaselineScan(
        OfficeByteView scanData,
        ScanHeader scan,
        JpegFrame frame,
        BaselineState state,
        int[][] quantTables,
        HuffmanTable[] dcTables,
        HuffmanTable[] acTables,
        int restartInterval,
        bool allowTruncated,
        CancellationToken cancellationToken) {
        ValidateBaselineScan(scan, frame, quantTables, dcTables, acTables);

        var states = state.Components;
        for (var i = 0; i < scan.ComponentIndices.Length; i++) {
            var componentIndex = scan.ComponentIndices[i];
            var comp = frame.Components[componentIndex];
            states[componentIndex].Component = comp;
            states[componentIndex].PrevDc = 0;
            state.PrepareAcLookup(ref acTables[comp.AcTable]);
        }

        var reader = new JpegBitReader(scanData, allowTruncated, cancellationToken);
        var mcuIndex = 0;
        var isSingle = scan.ComponentIndices.Length == 1;
        var scanComponent = frame.Components[scan.ComponentIndices[0]];
        var scanMcuCols = isSingle ? GetNonInterleavedBlockCount(frame.Width, scanComponent.H, frame.MaxH) : state.McuCols;
        var scanMcuRows = isSingle ? GetNonInterleavedBlockCount(frame.Height, scanComponent.V, frame.MaxV) : state.McuRows;

        for (var my = 0; my < scanMcuRows; my++) {
            cancellationToken.ThrowIfCancellationRequested();
            for (var mx = 0; mx < scanMcuCols; mx++) {
                if (restartInterval > 0 && mcuIndex > 0 && mcuIndex % restartInterval == 0) {
                    reader.ExpectRestartMarker();
                    for (var i = 0; i < scan.ComponentIndices.Length; i++) {
                        states[scan.ComponentIndices[i]].PrevDc = 0;
                    }
                }

                if (isSingle) {
                    var compIndex = scan.ComponentIndices[0];
                    var componentState = states[compIndex];
                    DecodeBlockCoefficients(
                        ref reader,
                        dcTables[componentState.Component.DcTable],
                        acTables[componentState.Component.AcTable],
                        quantTables[componentState.Component.QuantId],
                        ref componentState.PrevDc,
                        componentState.BlockCoeffs);
                    WriteDctBlock(componentState, mx, my);
                } else {
                    for (var ci = 0; ci < scan.ComponentIndices.Length; ci++) {
                        var compIndex = scan.ComponentIndices[ci];
                        var componentState = states[compIndex];
                        var blocks = componentState.Component.H * componentState.Component.V;
                        for (var b = 0; b < blocks; b++) {
                            DecodeBlockCoefficients(
                                ref reader,
                                dcTables[componentState.Component.DcTable],
                                acTables[componentState.Component.AcTable],
                                quantTables[componentState.Component.QuantId],
                                ref componentState.PrevDc,
                                componentState.BlockCoeffs);

                            var blockX = mx * componentState.Component.H + (b % componentState.Component.H);
                            var blockY = my * componentState.Component.V + (b / componentState.Component.H);
                            WriteDctBlock(componentState, blockX, blockY);
                        }
                    }
                }

                if (reader.RestartMarkerSeen) {
                    for (var i = 0; i < scan.ComponentIndices.Length; i++) {
                        states[scan.ComponentIndices[i]].PrevDc = 0;
                    }
                    reader.RestartMarkerSeen = false;
                }

                mcuIndex++;
            }
        }

        for (var i = 0; i < scan.ComponentIndices.Length; i++) {
            state.DecodedComponents[scan.ComponentIndices[i]] = true;
        }
    }

    private static void ValidateBaselineScan(
        ScanHeader scan,
        JpegFrame frame,
        int[][] quantTables,
        HuffmanTable[] dcTables,
        HuffmanTable[] acTables) {
        if (frame.ComponentCount == 0) throw new FormatException("Invalid JPEG frame.");
        if (frame.ComponentCount < 1 || frame.ComponentCount > 255) {
            throw new FormatException("Unsupported JPEG component count.");
        }
        if (scan.Ss != 0 || scan.Se != 63 || scan.Ah != 0 || scan.Al != 0) {
            throw new FormatException("Invalid baseline JPEG scan parameters.");
        }

        if (frame.Precision == 8) EnsureStandardHuffmanTables(dcTables, acTables);

        for (var i = 0; i < scan.ComponentIndices.Length; i++) {
            var componentIndex = scan.ComponentIndices[i];
            var comp = frame.Components[componentIndex];
            if (comp.QuantId >= quantTables.Length || quantTables[comp.QuantId] is null) {
                throw new FormatException("Missing JPEG quantization table.");
            }
            if (comp.DcTable >= dcTables.Length || !dcTables[comp.DcTable].IsValid) {
                throw new FormatException("Missing JPEG DC Huffman table.");
            }
            if (comp.AcTable >= acTables.Length || !acTables[comp.AcTable].IsValid) {
                throw new FormatException("Missing JPEG AC Huffman table.");
            }
        }
    }

    private static void DecodeProgressiveScan(
        OfficeByteView scanData,
        ScanHeader scan,
        JpegFrame frame,
        ProgressiveState state,
        int[][] quantTables,
        HuffmanTable[] dcTables,
        HuffmanTable[] acTables,
        int restartInterval,
        bool allowTruncated,
        CancellationToken cancellationToken,
        BandOverlap? overlap = null) {
        ValidateProgressiveScan(scan, frame, quantTables, dcTables, acTables);
        // Progressive scans are lenient to match historical behavior.
        var reader = new JpegBitReader(scanData, allowTruncated, cancellationToken);
        var mcuIndex = 0;
        var eobRun = 0;
        var isSingle = scan.ComponentIndices.Length == 1;
        var scanComponent = frame.Components[scan.ComponentIndices[0]];
        var scanMcuCols = isSingle ? GetNonInterleavedBlockCount(frame.Width, scanComponent.H, frame.MaxH) : state.McuCols;
        var scanMcuRows = GetProgressiveScanRows(scan, frame, state);

        for (var i = 0; i < scan.ComponentIndices.Length; i++) {
            int componentIndex = scan.ComponentIndices[i];
            // SOS headers may select different Huffman tables for later progressive scans.
            // Keep the retained coefficient state synchronized with the current scan declaration.
            state.Components[componentIndex].Component = frame.Components[componentIndex];
            state.Components[componentIndex].PrevDc = 0;
            if (scan.Se > 0 && scan.Ah == 0) state.PrepareAcLookup(ref acTables[frame.Components[componentIndex].AcTable]);
        }

        // Validation fixed the tables for this scan, so the block kind and each
        // component's tables are resolved once. Malformed spectral ranges keep the
        // general block decoder, which reports them.
        ProgressiveBlockKind kind = scan.Ss == 0 && scan.Se == 0
            ? (scan.Ah == 0 ? ProgressiveBlockKind.DcFirst : ProgressiveBlockKind.DcRefine)
            : scan.Ss > 0 && scan.Ss <= scan.Se && scan.Se <= 63
                ? (scan.Ah == 0 ? ProgressiveBlockKind.AcFirst : ProgressiveBlockKind.AcRefine)
                : ProgressiveBlockKind.General;
        var singleState = state.Components[scan.ComponentIndices[0]];
        if (kind == ProgressiveBlockKind.DcFirst || kind == ProgressiveBlockKind.DcRefine) {
            DecodeProgressiveDcScan(ref reader, kind, scan, state, dcTables, scanMcuCols, scanMcuRows,
                restartInterval, cancellationToken);
            return;
        }

        for (var my = 0; my < scanMcuRows; my++) {
            cancellationToken.ThrowIfCancellationRequested();
            for (var mx = 0; mx < scanMcuCols; mx++) {
                if (restartInterval > 0 && mcuIndex > 0 && mcuIndex % restartInterval == 0) {
                    reader.ExpectRestartMarker();
                    for (var i = 0; i < scan.ComponentIndices.Length; i++) {
                        state.Components[scan.ComponentIndices[i]].PrevDc = 0;
                    }
                    eobRun = 0;
                }

                if (isSingle) {
                    DecodeProgressiveBlock(ref reader, kind, scan, singleState, dcTables, acTables, quantTables,
                        mx, my, ref eobRun);
                } else {
                    for (var ci = 0; ci < scan.ComponentIndices.Length; ci++) {
                        var compState = state.Components[scan.ComponentIndices[ci]];
                        int horizontal = compState.Component.H;
                        int vertical = compState.Component.V;
                        for (var v = 0; v < vertical; v++) {
                            for (var h = 0; h < horizontal; h++) {
                                DecodeProgressiveBlock(ref reader, kind, scan, compState, dcTables, acTables, quantTables,
                                    mx * horizontal + h, my * vertical + v, ref eobRun);
                            }
                        }
                    }
                }

                if (reader.RestartMarkerSeen) {
                    for (var i = 0; i < scan.ComponentIndices.Length; i++) {
                        state.Components[scan.ComponentIndices[i]].PrevDc = 0;
                    }
                    reader.RestartMarkerSeen = false;
                    eobRun = 0;
                }

                mcuIndex++;
            }
            overlap?.Publish(my + 1);
        }
    }

    private static int GetProgressiveScanRows(ScanHeader scan, JpegFrame frame, ProgressiveState state) =>
        scan.ComponentIndices.Length == 1
            ? GetNonInterleavedBlockCount(frame.Height, frame.Components[scan.ComponentIndices[0]].V, frame.MaxV)
            : state.McuRows;

    // DC scans decode one short symbol (first scan) or one bit (refinement) per block, so the
    // MCU walk itself dominates; this loop keeps the bit buffer in locals across blocks and
    // hands it back to the reader at restart boundaries and whenever a step needs the
    // reader's checked methods. Blocks, restart handling and coefficients match the general
    // loop in DecodeProgressiveScan.
    private static void DecodeProgressiveDcScan(
        ref JpegBitReader reader,
        ProgressiveBlockKind kind,
        ScanHeader scan,
        ProgressiveState state,
        HuffmanTable[] dcTables,
        int scanMcuCols,
        int scanMcuRows,
        int restartInterval,
        CancellationToken cancellationToken) {
        int[] componentIndices = scan.ComponentIndices;
        bool isSingle = componentIndices.Length == 1;
        int al = scan.Al;
        bool first = kind == ProgressiveBlockKind.DcFirst;
        ushort refineBit = unchecked((ushort)(1 << al));
        var mcuIndex = 0;
        reader.FillBuffered();
        ulong buffer = reader.BufferedBits;
        int count = reader.BufferedBitCount;
        for (var my = 0; my < scanMcuRows; my++) {
            cancellationToken.ThrowIfCancellationRequested();
            for (var mx = 0; mx < scanMcuCols; mx++) {
                if (restartInterval > 0 && mcuIndex > 0 && mcuIndex % restartInterval == 0) {
                    reader.BufferedBitCount = count;
                    reader.ExpectRestartMarker();
                    for (var i = 0; i < componentIndices.Length; i++) {
                        state.Components[componentIndices[i]].PrevDc = 0;
                    }
                    reader.FillBuffered();
                    buffer = reader.BufferedBits;
                    count = reader.BufferedBitCount;
                }
                if (count < BufferedSymbolBits) {
                    reader.BufferedBitCount = count;
                    reader.FillBuffered();
                    buffer = reader.BufferedBits;
                    count = reader.BufferedBitCount;
                }

                for (var ci = 0; ci < componentIndices.Length; ci++) {
                    var component = state.Components[componentIndices[ci]];
                    int horizontal = isSingle ? 1 : component.Component.H;
                    int vertical = isSingle ? 1 : component.Component.V;
                    short[] coeffs = component.Coeffs;
                    for (var v = 0; v < vertical; v++) {
                        for (var h = 0; h < horizontal; h++) {
                            int blockX = isSingle ? mx : mx * horizontal + h;
                            int blockY = isSingle ? my : my * vertical + v;
                            int baseIndex = (blockY * component.BlocksPerRow + blockX) * 64;
                            if (!first) {
                                if (ReadBufferedBits(ref reader, ref buffer, ref count, 1) != 0) {
                                    coeffs[baseIndex] = unchecked((short)((ushort)coeffs[baseIndex] | refineBit));
                                }
                                continue;
                            }
                            if (count < BufferedSymbolBits) {
                                reader.BufferedBitCount = count;
                                reader.FillBuffered();
                                buffer = reader.BufferedBits;
                                count = reader.BufferedBitCount;
                            }
                            ref readonly HuffmanTable table = ref dcTables[component.Component.DcTable];
                            int t;
                            int entry;
                            if (count >= BufferedSymbolBits && table.Fast is not null &&
                                (entry = table.Fast[(int)(buffer >> (count - HuffmanFastBits)) & ((1 << HuffmanFastBits) - 1)]) >= 0) {
                                count -= entry >> 8;
                                t = entry & 0xFF;
                            } else {
                                reader.BufferedBitCount = count;
                                t = DecodeHuffman(ref reader, table, useFast: true);
                                buffer = reader.BufferedBits;
                                count = reader.BufferedBitCount;
                            }
                            var diff = t == 0 ? 0 : Extend(ReadBufferedBits(ref reader, ref buffer, ref count, t), t);
                            var dc = component.PrevDc + (diff << al);
                            component.PrevDc = dc;
                            SetProgressiveCoefficient(coeffs, baseIndex, dc);
                        }
                    }
                }

                if (reader.RestartMarkerSeen) {
                    for (var i = 0; i < componentIndices.Length; i++) {
                        state.Components[componentIndices[i]].PrevDc = 0;
                    }
                    reader.RestartMarkerSeen = false;
                }

                mcuIndex++;
            }
        }
        reader.BufferedBitCount = count;
    }

    private enum ProgressiveBlockKind {
        General,
        DcFirst,
        DcRefine,
        AcFirst,
        AcRefine
    }

    [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.AggressiveInlining)]
    private static void DecodeProgressiveBlock(
        ref JpegBitReader reader,
        ProgressiveBlockKind kind,
        ScanHeader scan,
        ProgressiveComponentState state,
        HuffmanTable[] dcTables,
        HuffmanTable[] acTables,
        int[][] quantTables,
        int blockX,
        int blockY,
        ref int eobRun) {
        var baseIndex = (blockY * state.BlocksPerRow + blockX) * 64;
        switch (kind) {
            case ProgressiveBlockKind.DcFirst: {
                    var t = DecodeHuffman(ref reader, dcTables[state.Component.DcTable], useFast: true);
                    var diff = t == 0 ? 0 : Extend(reader.ReadBits(t), t);
                    var dc = state.PrevDc + (diff << scan.Al);
                    state.PrevDc = dc;
                    SetProgressiveCoefficient(state.Coeffs, baseIndex, dc);
                    return;
                }
            case ProgressiveBlockKind.DcRefine:
                if (reader.ReadBit() != 0) {
                    state.Coeffs[baseIndex] = unchecked((short)((ushort)state.Coeffs[baseIndex] | (ushort)(1 << scan.Al)));
                }
                return;
            case ProgressiveBlockKind.AcFirst:
                DecodeProgressiveAcFirst(ref reader, scan, state, acTables[state.Component.AcTable], baseIndex, ref eobRun);
                return;
            case ProgressiveBlockKind.AcRefine:
                DecodeProgressiveAcRefine(ref reader, scan, state, acTables[state.Component.AcTable], baseIndex, ref eobRun);
                return;
            default:
                DecodeProgressiveBlock(ref reader, scan, state, dcTables, acTables, quantTables, blockX, blockY, ref eobRun);
                return;
        }
    }

    private static void ValidateProgressiveScan(
        ScanHeader scan,
        JpegFrame frame,
        int[][] quantTables,
        HuffmanTable[] dcTables,
        HuffmanTable[] acTables) {
        if (frame.Precision == 8) EnsureStandardHuffmanTables(dcTables, acTables);
        if (scan.Ss > 0 && scan.ComponentIndices.Length != 1) {
            throw new FormatException("Progressive JPEG AC scans must contain exactly one component.");
        }
        foreach (int componentIndex in scan.ComponentIndices) {
            Component component = frame.Components[componentIndex];
            if (component.QuantId >= quantTables.Length || quantTables[component.QuantId] is null) {
                throw new FormatException("Missing JPEG quantization table.");
            }
            if (scan.Ss == 0 && (component.DcTable >= dcTables.Length || !dcTables[component.DcTable].IsValid)) {
                throw new FormatException("Missing JPEG DC Huffman table.");
            }
            if (scan.Se > 0 && (component.AcTable >= acTables.Length || !acTables[component.AcTable].IsValid)) {
                throw new FormatException("Missing JPEG AC Huffman table.");
            }
        }
    }

    private static void DecodeProgressiveBlock(
        ref JpegBitReader reader,
        ScanHeader scan,
        ProgressiveComponentState state,
        HuffmanTable[] dcTables,
        HuffmanTable[] acTables,
        int[][] quantTables,
        int blockX,
        int blockY,
        ref int eobRun) {
        if (quantTables[state.Component.QuantId] is null) throw new FormatException("Missing JPEG quantization table.");
        var baseIndex = (blockY * state.BlocksPerRow + blockX) * 64;

        if (scan.Ss == 0 && scan.Se == 0) {
            ref readonly HuffmanTable dcTable = ref dcTables[state.Component.DcTable];
            if (!dcTable.IsValid) throw new FormatException("Missing JPEG DC Huffman table.");
            if (scan.Ah == 0) {
                var t = DecodeHuffman(ref reader, dcTable, useFast: true);
                var diff = t == 0 ? 0 : Extend(reader.ReadBits(t), t);
                var dc = state.PrevDc + (diff << scan.Al);
                state.PrevDc = dc;
                SetProgressiveCoefficient(state.Coeffs, baseIndex, dc);
            } else {
                var bit = reader.ReadBit();
                if (bit != 0) {
                    var delta = 1 << scan.Al;
                    int refined = unchecked((short)((ushort)state.Coeffs[baseIndex] | (ushort)delta));
                    SetProgressiveCoefficient(state.Coeffs, baseIndex, refined);
                }
            }
            return;
        }

        if (scan.Ss > scan.Se || scan.Se > 63) throw new FormatException("Invalid JPEG spectral selection.");
        ref readonly HuffmanTable acTable = ref acTables[state.Component.AcTable];
        if (!acTable.IsValid) throw new FormatException("Missing JPEG AC Huffman table.");

        if (scan.Ah == 0) {
            DecodeProgressiveAcFirst(ref reader, scan, state, acTable, baseIndex, ref eobRun);
        } else {
            DecodeProgressiveAcRefine(ref reader, scan, state, acTable, baseIndex, ref eobRun);
        }
    }

    private static int GetNonInterleavedBlockCount(int pixels, int componentSampling, int maximumSampling) =>
        (pixels * componentSampling + (maximumSampling * 8) - 1) / (maximumSampling * 8);

    // Most symbols decode from bits already in the reader's buffer, held in locals so the
    // loops below do not round-trip the reader through memory. Whenever the buffered bits
    // are not enough for a step, the reader's checked methods take it, so markers, the end
    // of the data and malformed codes are handled exactly as a bit-by-bit decode handles them.
    private const int BufferedSymbolBits = 32;

    private static void DecodeProgressiveAcFirst(
        ref JpegBitReader reader,
        ScanHeader scan,
        ProgressiveComponentState state,
        in HuffmanTable acTable,
        int baseIndex,
        ref int eobRun) {
        if (eobRun > 0) {
            eobRun--;
            return;
        }

        var k = (int)scan.Ss;
        int end = scan.Se;
        int al = scan.Al;
        short[] coeffs = state.Coeffs;
        byte[] zigZag = ZigZag;
        ref ulong nonzeroPositions = ref state.NonzeroPositions[baseIndex >> 6];
        short[]? fastAc = acTable.FastAc;
        short[]? fast = acTable.Fast;
        reader.FillBuffered();
        ulong buffer = reader.BufferedBits;
        int count = reader.BufferedBitCount;
        while (k <= end) {
            if (count < BufferedSymbolBits) {
                reader.BufferedBitCount = count;
                reader.FillBuffered();
                buffer = reader.BufferedBits;
                count = reader.BufferedBitCount;
            }
            int rs;
            if (count >= BufferedSymbolBits && fastAc is not null && fast is not null) {
                int peek = (int)(buffer >> (count - HuffmanFastBits)) & ((1 << HuffmanFastBits) - 1);
                // Codes whose symbol and magnitude bits fit the lookup decode in one step.
                int entry = fastAc[peek];
                if (entry != 0) {
                    int coefficient = entry >> 8;
                    int run = (entry >> 4) & 15;
                    if (coefficient != 0) {
                        if (k + run <= end) {
                            count -= entry & 15;
                            k += run;
                            SetProgressiveCoefficient(coeffs, baseIndex + zigZag[k], coefficient << al);
                            nonzeroPositions |= 1UL << k;
                            k++;
                            continue;
                        }
                    } else {
                        count -= entry & 15;
                        if (run == 15) {
                            k += 16;
                            continue;
                        }
                        // At most fourteen run bits follow, and BufferedSymbolBits covers them.
                        eobRun = (1 << run) - 1;
                        if (run > 0) {
                            count -= run;
                            eobRun += (int)(buffer >> count) & ((1 << run) - 1);
                        }
                        break;
                    }
                }
                entry = fast[peek];
                if (entry >= 0) {
                    count -= entry >> 8;
                    rs = entry & 0xFF;
                } else {
                    reader.BufferedBitCount = count;
                    rs = DecodeHuffman(ref reader, acTable, useFast: false);
                    buffer = reader.BufferedBits;
                    count = reader.BufferedBitCount;
                }
            } else {
                reader.BufferedBitCount = count;
                rs = DecodeHuffman(ref reader, acTable, useFast: true);
                buffer = reader.BufferedBits;
                count = reader.BufferedBitCount;
            }
            if (rs == 0) {
                eobRun = 0;
                break;
            }
            var r = rs >> 4;
            var s = rs & 0x0F;
            if (s == 0) {
                if (r == 15) {
                    k += 16;
                    continue;
                }
                eobRun = (1 << r) - 1;
                if (r > 0) eobRun += ReadBufferedBits(ref reader, ref buffer, ref count, r);
                break;
            }

            k += r;
            if (k > end) break;
            var ac = Extend(ReadBufferedBits(ref reader, ref buffer, ref count, s), s);
            var zig = zigZag[k];
            SetProgressiveCoefficient(coeffs, baseIndex + zig, ac << al);
            nonzeroPositions |= 1UL << k;
            k++;
        }
        reader.BufferedBitCount = count;
    }

    // Reads up to sixteen bits from the locals, or through the reader when fewer are buffered.
    [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.AggressiveInlining)]
    private static int ReadBufferedBits(ref JpegBitReader reader, ref ulong buffer, ref int count, int bits) {
        if (bits <= count) {
            count -= bits;
            return (int)(buffer >> count) & ((1 << bits) - 1);
        }
        reader.BufferedBitCount = count;
        int value = reader.ReadBits(bits);
        buffer = reader.BufferedBits;
        count = reader.BufferedBitCount;
        return value;
    }

    // Successive-approximation AC refinement. Each block keeps which of its coefficients
    // are nonzero (bit j = zigzag position j); Huffman AC scans set a bit whenever they
    // store a coefficient, and stored AC coefficients are never zero. Zero runs
    // are then located in that mask, and every nonzero coefficient passed on the way takes
    // one correction bit, in zigzag order, exactly as a position-by-position walk would.
    private static void DecodeProgressiveAcRefine(
        ref JpegBitReader reader,
        ScanHeader scan,
        ProgressiveComponentState state,
        in HuffmanTable acTable,
        int baseIndex,
        ref int eobRun) {
        short[] coeffs = state.Coeffs;
        byte[] zigZag = ZigZag;
        int end = scan.Se;
        int al = scan.Al;
        var k = (int)scan.Ss;
        ref ulong nonzeroPositions = ref state.NonzeroPositions[baseIndex >> 6];
        ulong nonzero = nonzeroPositions & BandMask(k, end);
        // DC scans do not track position zero; a malformed band that includes it reads
        // the coefficient itself.
        if (k == 0) nonzero = (nonzero & ~1UL) | (coeffs[baseIndex] != 0 ? 1UL : 0UL);
        reader.FillBuffered();
        ulong buffer = reader.BufferedBits;
        int count = reader.BufferedBitCount;
        if (eobRun > 0) {
            count = RefineBuffered(ref reader, ref buffer, count, coeffs, zigZag, baseIndex, nonzero, al);
            reader.BufferedBitCount = count;
            eobRun--;
            return;
        }

        short[]? fast = acTable.Fast;
        while (k <= end) {
            if (count < BufferedSymbolBits) {
                reader.BufferedBitCount = count;
                reader.FillBuffered();
                buffer = reader.BufferedBits;
                count = reader.BufferedBitCount;
            }
            int rs;
            int entry;
            if (count >= BufferedSymbolBits && fast is not null &&
                (entry = fast[(int)(buffer >> (count - HuffmanFastBits)) & ((1 << HuffmanFastBits) - 1)]) >= 0) {
                count -= entry >> 8;
                rs = entry & 0xFF;
            } else {
                reader.BufferedBitCount = count;
                rs = DecodeHuffman(ref reader, acTable, useFast: true);
                buffer = reader.BufferedBits;
                count = reader.BufferedBitCount;
            }
            var r = rs >> 4;
            var s = rs & 0x0F;
            ulong remaining = BandMask(k, end);

            if (s == 0) {
                if (r == 15) {
                    // Sixteen zeros are skipped, or the rest of the band when it has fewer.
                    ulong zeros = ~nonzero & remaining;
                    for (int skipped = 0; skipped < 15 && zeros != 0; skipped++) zeros &= zeros - 1;
                    int stop = zeros == 0 ? end + 1 : LowestSetBit(zeros) + 1;
                    count = RefineBuffered(ref reader, ref buffer, count, coeffs, zigZag, baseIndex,
                        nonzero & remaining & BandPrefix(stop), al);
                    k = stop;
                    continue;
                }

                eobRun = (1 << r) - 1;
                if (r > 0) eobRun += ReadBufferedBits(ref reader, ref buffer, ref count, r);
                count = RefineBuffered(ref reader, ref buffer, count, coeffs, zigZag, baseIndex, nonzero & remaining, al);
                break;
            }

            if (s != 1) throw new FormatException("Invalid progressive JPEG AC refinement symbol.");
            var ac = ReadBufferedBits(ref reader, ref buffer, ref count, 1) == 1 ? 1 : -1;
            // The new coefficient lands on the zero that follows r skipped zeros.
            ulong candidates = ~nonzero & remaining;
            for (int skipped = 0; skipped < r && candidates != 0; skipped++) candidates &= candidates - 1;
            if (candidates == 0) {
                count = RefineBuffered(ref reader, ref buffer, count, coeffs, zigZag, baseIndex, nonzero & remaining, al);
                k = end + 1;
                break;
            }
            int position = LowestSetBit(candidates);
            count = RefineBuffered(ref reader, ref buffer, count, coeffs, zigZag, baseIndex,
                nonzero & remaining & BandPrefix(position), al);
            SetProgressiveCoefficient(coeffs, baseIndex + zigZag[position], ac << al);
            nonzeroPositions |= 1UL << position;
            k = position + 1;
        }
        reader.BufferedBitCount = count;
    }

    // Applies one correction bit to each coefficient in the mask from the locals when they
    // hold enough bits; otherwise the reader supplies them. Returns the buffered bit count.
    [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.AggressiveInlining)]
    private static int RefineBuffered(ref JpegBitReader reader, ref ulong buffer, int count, short[] coeffs,
        byte[] zigZag, int baseIndex, ulong mask, int al) {
        if (mask == 0) return count;
        if (PopCount(mask) <= count) {
            do {
                int index = baseIndex + zigZag[LowestSetBit(mask)];
                mask &= mask - 1;
                count--;
                coeffs[index] = ApplyCorrectionBit(coeffs[index], (int)(buffer >> count) & 1, al);
            } while (mask != 0);
            return count;
        }
        reader.BufferedBitCount = count;
        RefineMasked(ref reader, coeffs, zigZag, baseIndex, mask, al);
        buffer = reader.BufferedBits;
        return reader.BufferedBitCount;
    }

    private static int PopCount(ulong value) {
#if NET8_0_OR_GREATER
        return System.Numerics.BitOperations.PopCount(value);
#else
        value -= (value >> 1) & 0x5555555555555555UL;
        value = (value & 0x3333333333333333UL) + ((value >> 2) & 0x3333333333333333UL);
        value = (value + (value >> 4)) & 0x0F0F0F0F0F0F0F0FUL;
        return (int)((value * 0x0101010101010101UL) >> 56);
#endif
    }

    // Positions start..end (end at most 63).
    private static ulong BandMask(int start, int end) => (ulong.MaxValue << start) & (ulong.MaxValue >> (63 - end));

    // Positions below stop (stop at most 64).
    private static ulong BandPrefix(int stop) => stop >= 64 ? ulong.MaxValue : (1UL << stop) - 1;

    private static int LowestSetBit(ulong value) {
#if NET8_0_OR_GREATER
        return System.Numerics.BitOperations.TrailingZeroCount(value);
#else
        int position = 0;
        if ((value & 0xFFFFFFFFUL) == 0) { position += 32; value >>= 32; }
        if ((value & 0xFFFFUL) == 0) { position += 16; value >>= 16; }
        if ((value & 0xFFUL) == 0) { position += 8; value >>= 8; }
        if ((value & 0xFUL) == 0) { position += 4; value >>= 4; }
        if ((value & 0x3UL) == 0) { position += 2; value >>= 2; }
        if ((value & 0x1UL) == 0) position++;
        return position;
#endif
    }

    // Applies one correction bit to each coefficient in the mask, lowest position first.
    // The bit buffer stays in locals; bits are applied without branching on their values.
    private static void RefineMasked(ref JpegBitReader reader, short[] coeffs, byte[] zigZag,
        int baseIndex, ulong mask, int al) {
        if (mask == 0) return;
        ulong buffer = reader.BufferedBits;
        int count = reader.BufferedBitCount;
        do {
            int index = baseIndex + zigZag[LowestSetBit(mask)];
            mask &= mask - 1;
            if (count == 0) {
                reader.BufferedBitCount = count;
                reader.EnsureBits(1);
                buffer = reader.BufferedBits;
                count = reader.BufferedBitCount;
            }
            count--;
            coeffs[index] = ApplyCorrectionBit(coeffs[index], (int)(buffer >> count) & 1, al);
        } while (mask != 0);
        reader.BufferedBitCount = count;
    }

    // A set correction bit moves a nonzero coefficient one step of 2^al away from zero
    // unless that bit is already set; a zero coefficient always reads a zero bit.
    [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.AggressiveInlining)]
    private static short ApplyCorrectionBit(int value, int bit, int al) {
        int delta = 1 << al;
        int apply = bit & (((ushort)value & (ushort)delta) == 0 ? 1 : 0);
        int sign = value >> 31;
        int refined = value + (((delta ^ sign) - sign) & -apply);
        if ((uint)(refined - short.MinValue) > ushort.MaxValue) {
            throw new FormatException("Progressive JPEG coefficient exceeds the signed sixteen-bit range.");
        }
        return (short)refined;
    }

    [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.AggressiveInlining)]
    private static void SetProgressiveCoefficient(short[] coefficients, int index, int value) {
        if (value < short.MinValue || value > short.MaxValue) {
            throw new FormatException("Progressive JPEG coefficient exceeds the signed sixteen-bit range.");
        }
        coefficients[index] = (short)value;
    }

}
