using System;

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
        bool allowTruncated) {
        ValidateBaselineScan(scan, frame, quantTables, dcTables, acTables);

        var states = state.Components;
        for (var i = 0; i < scan.ComponentIndices.Length; i++) {
            var componentIndex = scan.ComponentIndices[i];
            var comp = frame.Components[componentIndex];
            states[componentIndex].Component = comp;
            states[componentIndex].PrevDc = 0;
        }

        var reader = new JpegBitReader(scanData, allowTruncated);
        var mcuIndex = 0;
        var isSingle = scan.ComponentIndices.Length == 1;
        var scanComponent = frame.Components[scan.ComponentIndices[0]];
        var scanMcuCols = isSingle ? GetNonInterleavedBlockCount(frame.Width, scanComponent.H, frame.MaxH) : state.McuCols;
        var scanMcuRows = isSingle ? GetNonInterleavedBlockCount(frame.Height, scanComponent.V, frame.MaxV) : state.McuRows;

        for (var my = 0; my < scanMcuRows; my++) {
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
                    DecodeBlock(
                        ref reader,
                        dcTables[componentState.Component.DcTable],
                        acTables[componentState.Component.AcTable],
                        quantTables[componentState.Component.QuantId],
                        ref componentState.PrevDc,
                        componentState.BlockCoeffs,
                        componentState.BlockPixels);
                    WriteBlock(componentState.Buffer, componentState.Stride, mx, my, componentState.BlockPixels);
                } else {
                    for (var ci = 0; ci < scan.ComponentIndices.Length; ci++) {
                        var compIndex = scan.ComponentIndices[ci];
                        var componentState = states[compIndex];
                        var blocks = componentState.Component.H * componentState.Component.V;
                        for (var b = 0; b < blocks; b++) {
                            DecodeBlock(
                                ref reader,
                                dcTables[componentState.Component.DcTable],
                                acTables[componentState.Component.AcTable],
                                quantTables[componentState.Component.QuantId],
                                ref componentState.PrevDc,
                                componentState.BlockCoeffs,
                                componentState.BlockPixels);

                            var blockX = mx * componentState.Component.H + (b % componentState.Component.H);
                            var blockY = my * componentState.Component.V + (b / componentState.Component.H);
                            WriteBlock(componentState.Buffer, componentState.Stride, blockX, blockY, componentState.BlockPixels);
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
        if (frame.ComponentCount != 1 && frame.ComponentCount != 3 && frame.ComponentCount != 4) {
            throw new FormatException("Unsupported JPEG component count.");
        }
        if (scan.Ss != 0 || scan.Se != 63 || scan.Ah != 0 || scan.Al != 0) {
            throw new FormatException("Invalid baseline JPEG scan parameters.");
        }

        EnsureStandardHuffmanTables(dcTables, acTables);

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
        bool allowTruncated) {
        ValidateProgressiveScan(scan, frame, quantTables, dcTables, acTables);
        // Progressive scans are lenient to match historical behavior.
        var reader = new JpegBitReader(scanData, allowTruncated);
        var mcuIndex = 0;
        var eobRun = 0;
        var isSingle = scan.ComponentIndices.Length == 1;
        var scanComponent = frame.Components[scan.ComponentIndices[0]];
        var scanMcuCols = isSingle ? GetNonInterleavedBlockCount(frame.Width, scanComponent.H, frame.MaxH) : state.McuCols;
        var scanMcuRows = isSingle ? GetNonInterleavedBlockCount(frame.Height, scanComponent.V, frame.MaxV) : state.McuRows;

        for (var i = 0; i < scan.ComponentIndices.Length; i++) {
            state.Components[scan.ComponentIndices[i]].PrevDc = 0;
        }

        for (var my = 0; my < scanMcuRows; my++) {
            for (var mx = 0; mx < scanMcuCols; mx++) {
                if (restartInterval > 0 && mcuIndex > 0 && mcuIndex % restartInterval == 0) {
                    reader.ExpectRestartMarker();
                    for (var i = 0; i < scan.ComponentIndices.Length; i++) {
                        state.Components[scan.ComponentIndices[i]].PrevDc = 0;
                    }
                    eobRun = 0;
                }

                if (isSingle) {
                    var compIndex = scan.ComponentIndices[0];
                    DecodeProgressiveBlock(
                        ref reader,
                        scan,
                        state.Components[compIndex],
                        dcTables,
                        acTables,
                        quantTables,
                        mx,
                        my,
                        ref eobRun);
                } else {
                    for (var ci = 0; ci < scan.ComponentIndices.Length; ci++) {
                        var compIndex = scan.ComponentIndices[ci];
                        var compState = state.Components[compIndex];
                        var blocks = compState.Component.H * compState.Component.V;
                        for (var b = 0; b < blocks; b++) {
                            var blockX = mx * compState.Component.H + (b % compState.Component.H);
                            var blockY = my * compState.Component.V + (b / compState.Component.H);
                            DecodeProgressiveBlock(
                                ref reader,
                                scan,
                                compState,
                                dcTables,
                                acTables,
                                quantTables,
                                blockX,
                                blockY,
                                ref eobRun);
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
        }
    }

    private static void ValidateProgressiveScan(
        ScanHeader scan,
        JpegFrame frame,
        int[][] quantTables,
        HuffmanTable[] dcTables,
        HuffmanTable[] acTables) {
        EnsureStandardHuffmanTables(dcTables, acTables);
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
        var quant = quantTables[state.Component.QuantId];
        if (quant is null) throw new FormatException("Missing JPEG quantization table.");
        var baseIndex = (blockY * state.BlocksPerRow + blockX) * 64;

        if (scan.Ss == 0 && scan.Se == 0) {
            var dcTable = dcTables[state.Component.DcTable];
            if (!dcTable.IsValid) throw new FormatException("Missing JPEG DC Huffman table.");
            if (scan.Ah == 0) {
                var t = DecodeHuffman(ref reader, dcTable, useFast: false);
                var diff = t == 0 ? 0 : Extend(reader.ReadBits(t), t);
                var dc = state.PrevDc + (diff << scan.Al);
                state.PrevDc = dc;
                state.Coeffs[baseIndex] = dc * quant[0];
            } else {
                var bit = reader.ReadBit();
                if (bit != 0) {
                    var delta = (1 << scan.Al) * quant[0];
                    var sign = state.Coeffs[baseIndex] >= 0 ? 1 : -1;
                    state.Coeffs[baseIndex] += sign * delta;
                }
            }
            return;
        }

        if (scan.Ss > scan.Se || scan.Se > 63) throw new FormatException("Invalid JPEG spectral selection.");
        var acTable = acTables[state.Component.AcTable];
        if (!acTable.IsValid) throw new FormatException("Missing JPEG AC Huffman table.");

        if (scan.Ah == 0) {
            DecodeProgressiveAcFirst(ref reader, scan, state, acTable, quant, baseIndex, ref eobRun);
        } else {
            DecodeProgressiveAcRefine(ref reader, scan, state, acTable, quant, baseIndex, ref eobRun);
        }
    }

    private static int GetNonInterleavedBlockCount(int pixels, int componentSampling, int maximumSampling) =>
        (pixels * componentSampling + (maximumSampling * 8) - 1) / (maximumSampling * 8);

    private static void DecodeProgressiveAcFirst(
        ref JpegBitReader reader,
        ScanHeader scan,
        ProgressiveComponentState state,
        HuffmanTable acTable,
        int[] quant,
        int baseIndex,
        ref int eobRun) {
        if (eobRun > 0) {
            eobRun--;
            return;
        }

        var k = (int)scan.Ss;
        while (k <= scan.Se) {
            var rs = DecodeHuffman(ref reader, acTable, useFast: false);
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
                if (r > 0) eobRun += reader.ReadBits(r);
                break;
            }

            k += r;
            if (k > scan.Se) break;
            var ac = Extend(reader.ReadBits(s), s);
            var zig = ZigZag[k];
            state.Coeffs[baseIndex + zig] = (ac << scan.Al) * quant[zig];
            k++;
        }
    }

    private static void DecodeProgressiveAcRefine(
        ref JpegBitReader reader,
        ScanHeader scan,
        ProgressiveComponentState state,
        HuffmanTable acTable,
        int[] quant,
        int baseIndex,
        ref int eobRun) {
        var k = (int)scan.Ss;
        if (eobRun > 0) {
            for (; k <= scan.Se; k++) {
                RefineCoefficient(ref reader, state.Coeffs, baseIndex + ZigZag[k], scan.Al, quant[ZigZag[k]]);
            }
            eobRun--;
            return;
        }

        while (k <= scan.Se) {
            var rs = DecodeHuffman(ref reader, acTable, useFast: false);
            var r = rs >> 4;
            var s = rs & 0x0F;

            if (s == 0) {
                if (r == 15) {
                    var zeros = 16;
                    while (zeros > 0 && k <= scan.Se) {
                        var index = baseIndex + ZigZag[k];
                        RefineCoefficient(ref reader, state.Coeffs, index, scan.Al, quant[ZigZag[k]]);
                        if (state.Coeffs[index] == 0) zeros--;
                        k++;
                    }
                    continue;
                }

                eobRun = (1 << r) - 1;
                if (r > 0) eobRun += reader.ReadBits(r);
                for (; k <= scan.Se; k++) {
                    RefineCoefficient(ref reader, state.Coeffs, baseIndex + ZigZag[k], scan.Al, quant[ZigZag[k]]);
                }
                break;
            }

            if (s != 1) throw new FormatException("Invalid progressive JPEG AC refinement symbol.");
            var ac = reader.ReadBit() == 1 ? 1 : -1;
            while (k <= scan.Se) {
                var zig = ZigZag[k];
                var index = baseIndex + zig;
                if (state.Coeffs[index] != 0) {
                    RefineCoefficient(ref reader, state.Coeffs, index, scan.Al, quant[zig]);
                    k++;
                    continue;
                }

                if (r > 0) {
                    r--;
                    k++;
                    continue;
                }

                state.Coeffs[index] = (ac << scan.Al) * quant[zig];
                k++;
                break;
            }
        }
    }

    private static void RefineCoefficient(ref JpegBitReader reader, int[] coeffs, int index, int al, int quant) {
        if (coeffs[index] == 0) return;
        var bit = reader.ReadBit();
        if (bit == 0) return;
        var delta = (1 << al) * quant;
        coeffs[index] += coeffs[index] > 0 ? delta : -delta;
    }

}
