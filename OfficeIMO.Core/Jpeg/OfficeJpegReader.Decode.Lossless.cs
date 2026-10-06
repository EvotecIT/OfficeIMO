using System;
using System.Threading;

namespace OfficeIMO.Drawing;

internal static partial class OfficeJpegReader {
    // T.81 Annex H: sample prediction, without DCT or quantization. The current
    // sample-plane renderer accepts two through sixteen-bit frames; point transforms retain the
    // original sample scale, with the discarded low bits restored as zero.
    private static void DecodeLosslessScan(OfficeByteView data, ScanHeader scan,
        JpegFrame frame, BaselineState state, HuffmanTable[] dcTables,
        int restartInterval, CancellationToken token) {
        if (scan.Ss < 1 || scan.Ss > 7 || scan.Se != 0 || scan.Ah != 0 || scan.Al >= frame.Precision)
            throw new FormatException("Invalid lossless JPEG scan parameters.");
        foreach (int index in scan.ComponentIndices) {
            var component = frame.Components[index];
            if (state.DecodedComponents[index] || component.QuantId != 0 || component.AcTable != 0 ||
                component.DcTable >= dcTables.Length || !dcTables[component.DcTable].IsValid)
                throw new FormatException("Invalid lossless JPEG component or Huffman table.");
        }

        bool single = scan.ComponentIndices.Length == 1;
        var first = frame.Components[scan.ComponentIndices[0]];
        int columns = single ? (frame.Width * first.H + frame.MaxH - 1) / frame.MaxH :
            (frame.Width + frame.MaxH - 1) / frame.MaxH;
        int rows = single ? (frame.Height * first.V + frame.MaxV - 1) / frame.MaxV :
            (frame.Height + frame.MaxV - 1) / frame.MaxV;
        if (restartInterval != 0 && restartInterval % columns != 0)
            throw new FormatException("Lossless JPEG restarts must align with MCU rows.");

        var reader = new JpegBitReader(data, allowTruncated: false, token);
        int mcu = 0, restartRow = 0, restartNumber = 0;
        int shift = scan.Al, mask = (1 << (frame.Precision - shift)) - 1, initial = 1 << (frame.Precision - 1 - shift);
        for (int my = 0; my < rows; my++) {
            token.ThrowIfCancellationRequested();
            for (int mx = 0; mx < columns; mx++, mcu++) {
                if ((mcu & 4095) == 0) token.ThrowIfCancellationRequested();
                if (restartInterval > 0 && mcu > 0 && mcu % restartInterval == 0) {
                    reader.ExpectRestartMarker(0xD0 + (restartNumber++ & 7));
                    restartRow = my;
                }
                foreach (int index in scan.ComponentIndices) {
                    var component = frame.Components[index];
                    var pixels = state.Components[index];
                    int h = single ? 1 : component.H, v = single ? 1 : component.V;
                    for (int dy = 0; dy < v; dy++) for (int dx = 0; dx < h; dx++) {
                        int x = mx * h + dx, y = my * v + dy, at = y * pixels.Stride + x;
                        int prediction;
                        if (y == restartRow * v) {
                            prediction = x == 0 ? initial : pixels.ReadSample(at - 1) >> shift;
                        } else if (x == 0) {
                            prediction = pixels.ReadSample(at - pixels.Stride) >> shift;
                        } else {
                            int a = pixels.ReadSample(at - 1) >> shift;
                            int b = pixels.ReadSample(at - pixels.Stride) >> shift;
                            int c = pixels.ReadSample(at - pixels.Stride - 1) >> shift;
                            prediction = scan.Ss switch {
                                1 => a, 2 => b, 3 => c, 4 => a + b - c,
                                5 => a + ((b - c) >> 1), 6 => b + ((a - c) >> 1),
                                _ => (a + b) >> 1
                            };
                        }
                        int category = DecodeHuffman(ref reader, dcTables[component.DcTable], useFast: true);
                        if (category > 16) throw new FormatException("Invalid lossless JPEG difference category.");
                        // Category 16 represents -32768 without additional bits.
                        int difference = category == 16 ? -32768 : category == 0 ? 0 :
                            Extend(reader.ReadBits(category), category);
                        if (reader.RestartMarkerSeen) throw new FormatException("Unexpected lossless JPEG restart marker.");
                        pixels.WriteSample(at, ((prediction + difference) & mask) << shift);
                    }
                }
            }
        }
        foreach (int index in scan.ComponentIndices) {
            ExtendLosslessEdges(frame, state.Components[index], token);
            state.DecodedComponents[index] = true;
        }
    }
    // Shared rendering planes are block-padded. Lossless entropy stores samples,
    // so extend the actual image edge before bilinear sampling touches padding.
    private static void ExtendLosslessEdges(JpegFrame frame, BaselineComponentState state, CancellationToken token) {
        int width = (frame.Width * state.Component.H + frame.MaxH - 1) / frame.MaxH;
        int height = (frame.Height * state.Component.V + frame.MaxV - 1) / frame.MaxV;
        for (int y = 0; y < height; y++) {
            token.ThrowIfCancellationRequested();
            int row = y * state.Stride;
            int edge = state.ReadSample(row + width - 1);
            for (int x = width; x < state.Stride; x++) state.WriteSample(row + x, edge);
        }
        for (int y = height; y < state.SampleCount / state.Stride; y++) {
            token.ThrowIfCancellationRequested();
            state.CopySamples((height - 1) * state.Stride, y * state.Stride, state.Stride);
        }
    }

}
