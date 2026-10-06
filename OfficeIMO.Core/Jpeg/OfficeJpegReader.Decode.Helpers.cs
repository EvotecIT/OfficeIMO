using System;
using System.Threading;

namespace OfficeIMO.Drawing;

internal static partial class OfficeJpegReader {
    private const int HuffmanFastBits = 9;
    private const int ConstBits = 13;
    private const int Pass1Bits = 2;
    private const string JpegDimensionsLimitMessage = "JPEG dimensions exceed limits.";
    private static readonly int[] CrToR = new int[256];
    private static readonly int[] CrToG = new int[256];
    private static readonly int[] CbToG = new int[256];
    private static readonly int[] CbToB = new int[256];

    // Fixed-point constants from the IJG islow integer IDCT implementation.
    private const long Fix0_298631336 = 2446;
    private const long Fix0_390180644 = 3196;
    private const long Fix0_541196100 = 4433;
    private const long Fix0_765366865 = 6270;
    private const long Fix0_899976223 = 7373;
    private const long Fix1_175875602 = 9633;
    private const long Fix1_501321110 = 12299;
    private const long Fix1_847759065 = 15137;
    private const long Fix1_961570560 = 16069;
    private const long Fix2_053119869 = 16819;
    private const long Fix2_562915447 = 20995;
    private const long Fix3_072711026 = 25172;

    static OfficeJpegReader() {
        for (var i = 0; i < 256; i++) {
            var d = i - 128;
            CrToR[i] = (91881 * d + 32768) >> 16;
            CrToG[i] = (46802 * d + 32768) >> 16;
            CbToG[i] = (22554 * d + 32768) >> 16;
            CbToB[i] = (116130 * d + 32768) >> 16;
        }
    }

    private static byte[] ComposeRgba(
        JpegFrame frame,
        BaselineComponentState[] states,
        int? adobeTransform,
        bool highQualityChroma,
        CancellationToken cancellationToken) {
        return ComposeColorComponents(
            frame,
            states,
            adobeTransform,
            requestedColorTransform: null,
            usePdfColorTransformDefault: false,
            highQualityChroma,
            outputRgba: true,
            cancellationToken,
            out _);
    }

    private static byte[] ComposeColorComponents(
        JpegFrame frame,
        BaselineComponentState[] states,
        int? adobeTransform,
        int? requestedColorTransform,
        bool usePdfColorTransformDefault,
        bool highQualityChroma,
        bool outputRgba,
        CancellationToken cancellationToken,
        out int componentCount) {
        componentCount = frame.ComponentCount;
        if (componentCount < 1 || componentCount > 255) {
            throw new FormatException("Unsupported JPEG component count.");
        }
        if (outputRgba && componentCount is not (1 or 3 or 4)) {
            throw new FormatException("Unsupported JPEG component count.");
        }
        byte[] components = outputRgba
            ? OfficeRasterGuards.AllocateRgba32(frame.Width, frame.Height, JpegDimensionsLimitMessage)
            : new byte[checked(frame.Width * frame.Height * componentCount)];
        var maxH = frame.MaxH;
        var maxV = frame.MaxV;

        if (!outputRgba && (componentCount == 2 || componentCount > 4 && requestedColorTransform == 0)) {
            CopyRawComponents(frame, states, highQualityChroma, cancellationToken, components);
            return components;
        }

        if (componentCount == 2 || componentCount > 4) throw new FormatException("A raw component transform is required for this JPEG component count.");

        if (frame.ComponentCount == 4) {
            // PDF Decode arrays own polarity. APP14 still selects the color transform,
            // but only standalone/ICC-image callers need Adobe's inverted CMYK normalized.
            bool invertAdobe = adobeTransform.HasValue && (!usePdfColorTransformDefault || outputRgba);
            var cIndex = FindComponentIndex(frame.Components, (byte)'C');
            var mIndex = FindComponentIndex(frame.Components, (byte)'M');
            var yIndex = FindComponentIndex(frame.Components, (byte)'Y');
            var kIndex = FindComponentIndex(frame.Components, (byte)'K');
            if (cIndex < 0 || mIndex < 0 || yIndex < 0 || kIndex < 0) {
                cIndex = FindComponentIndex(frame.Components, 1);
                mIndex = FindComponentIndex(frame.Components, 2);
                yIndex = FindComponentIndex(frame.Components, 3);
                kIndex = FindComponentIndex(frame.Components, 4);
                if (cIndex < 0 || mIndex < 0 || yIndex < 0 || kIndex < 0) {
                    cIndex = 0;
                    mIndex = 1;
                    yIndex = 2;
                    kIndex = 3;
                }
            }

            var isYcck = requestedColorTransform.HasValue
                ? requestedColorTransform.Value == 1
                : adobeTransform == 2;
            var ycckY = FindComponentIndex(frame.Components, 1);
            var ycckCb = FindComponentIndex(frame.Components, 2);
            var ycckCr = FindComponentIndex(frame.Components, 3);
            var ycckK = FindComponentIndex(frame.Components, 4);
            if (isYcck && (ycckY < 0 || ycckCb < 0 || ycckCr < 0 || ycckK < 0)) {
                ycckY = cIndex;
                ycckCb = mIndex;
                ycckCr = yIndex;
                ycckK = kIndex;
            }

            for (var y = 0; y < frame.Height; y++) {
                cancellationToken.ThrowIfCancellationRequested();
                for (var x = 0; x < frame.Width; x++) {
                    byte c;
                    byte m;
                    byte y0;
                    var kVal = SampleComponent(states, isYcck ? ycckK : kIndex, x, y, maxH, maxV, 0, highQualityChroma);

                    if (isYcck) {
                        SampleYccToRgb(frame, states, ycckY, ycckCb, ycckCr, x, y,
                            highQualityChroma, out byte r, out byte g, out byte b);
                        if (invertAdobe) {
                            // YCCK conversion already complements the three native CMY
                            // samples. Adobe inversion cancels that complement; K still inverts.
                            c = r;
                            m = g;
                            y0 = b;
                            kVal = 255 - kVal;
                        } else {
                            c = (byte)(255 - r);
                            m = (byte)(255 - g);
                            y0 = (byte)(255 - b);
                        }
                    } else {
                        c = (byte)SampleComponent(states, cIndex, x, y, maxH, maxV, 0, highQualityChroma);
                        m = (byte)SampleComponent(states, mIndex, x, y, maxH, maxV, 0, highQualityChroma);
                        y0 = (byte)SampleComponent(states, yIndex, x, y, maxH, maxV, 0, highQualityChroma);
                        if (invertAdobe) {
                            c = (byte)(255 - c);
                            m = (byte)(255 - m);
                            y0 = (byte)(255 - y0);
                            kVal = 255 - kVal;
                        }
                    }

                    WriteCmykPixel(components, y * frame.Width + x, c, m, y0, (byte)kVal, outputRgba);
                }
            }

            return components;
        }

        if (frame.ComponentCount == 1) {
            var grayIndex = FindComponentIndex(frame.Components, 1);
            if (grayIndex < 0) grayIndex = 0;
            for (var y = 0; y < frame.Height; y++) {
                cancellationToken.ThrowIfCancellationRequested();
                for (var x = 0; x < frame.Width; x++) {
                    var v = SampleComponent(states, grayIndex, x, y, maxH, maxV, 0, highQualityChroma);
                    WriteGrayPixel(components, y * frame.Width + x, (byte)v, outputRgba);
                }
            }
            return components;
        }

        var rIndex = FindComponentIndex(frame.Components, (byte)'R');
        var gIndex = FindComponentIndex(frame.Components, (byte)'G');
        var bIndex = FindComponentIndex(frame.Components, (byte)'B');
        var hasRgbComponentIds = rIndex >= 0 && gIndex >= 0 && bIndex >= 0;
        bool transformToRgb = requestedColorTransform.HasValue
            ? requestedColorTransform.Value == 1
            : adobeTransform.HasValue
                ? adobeTransform.Value == 1
                : usePdfColorTransformDefault || !hasRgbComponentIds;

        var yIndex2 = FindComponentIndex(frame.Components, 1);
        var cbIndex = frame.ComponentCount > 1 ? FindComponentIndex(frame.Components, 2) : -1;
        var crIndex = frame.ComponentCount > 1 ? FindComponentIndex(frame.Components, 3) : -1;
        if (frame.ComponentCount == 3) {
            bool hasConventionalYccIds = yIndex2 >= 0 && cbIndex >= 0 && crIndex >= 0;
            if (!hasConventionalYccIds) {
                yIndex2 = 0;
                cbIndex = 1;
                crIndex = 2;
            }

            if (!highQualityChroma && (!transformToRgb || frame.Precision == 8)) {
                int firstIndex = transformToRgb ? yIndex2 : hasRgbComponentIds ? rIndex : 0;
                int secondIndex = transformToRgb ? cbIndex : hasRgbComponentIds ? gIndex : 1;
                int thirdIndex = transformToRgb ? crIndex : hasRgbComponentIds ? bIndex : 2;
                ComposeThreeComponentNearest(
                    components,
                    frame.Width,
                    frame.Height,
                    states[firstIndex],
                    states[secondIndex],
                    states[thirdIndex],
                    maxH,
                    maxV,
                    transformToRgb,
                    outputRgba,
                    cancellationToken);
                return components;
            }
        }

        for (var y = 0; y < frame.Height; y++) {
            cancellationToken.ThrowIfCancellationRequested();
            for (var x = 0; x < frame.Width; x++) {
                if (frame.ComponentCount == 3 && !transformToRgb) {
                    int firstIndex = hasRgbComponentIds ? rIndex : 0;
                    int secondIndex = hasRgbComponentIds ? gIndex : 1;
                    int thirdIndex = hasRgbComponentIds ? bIndex : 2;
                    WriteRgbPixel(
                        components,
                        y * frame.Width + x,
                        (byte)SampleComponent(states, firstIndex, x, y, maxH, maxV, 0, highQualityChroma),
                        (byte)SampleComponent(states, secondIndex, x, y, maxH, maxV, 0, highQualityChroma),
                        (byte)SampleComponent(states, thirdIndex, x, y, maxH, maxV, 0, highQualityChroma),
                        outputRgba);
                } else if (frame.ComponentCount == 3) {
                    SampleYccToRgb(frame, states, yIndex2, cbIndex, crIndex, x, y,
                        highQualityChroma, out byte r, out byte g, out byte b);
                    WriteRgbPixel(components, y * frame.Width + x, r, g, b, outputRgba);
                } else {
                    int p = (y * frame.Width + x) * componentCount;
                    for (int component = 0; component < componentCount; component++) {
                        components[p + component] = (byte)SampleComponent(
                            states,
                            component,
                            x,
                            y,
                            maxH,
                            maxV,
                            0,
                            highQualityChroma);
                    }
                }
            }
        }

        return components;
    }

    private static void ComposeThreeComponentNearest(
        byte[] output,
        int width,
        int height,
        BaselineComponentState first,
        BaselineComponentState second,
        BaselineComponentState third,
        int maximumHorizontalSampling,
        int maximumVerticalSampling,
        bool transformYccToRgb,
        bool outputRgba,
        CancellationToken cancellationToken) {
        int firstY = 0;
        int secondY = 0;
        int thirdY = 0;
        int firstYAccumulator = 0;
        int secondYAccumulator = 0;
        int thirdYAccumulator = 0;
        int target = 0;

        for (int y = 0; y < height; y++) {
            cancellationToken.ThrowIfCancellationRequested();
            int firstRow = firstY * first.Stride;
            int secondRow = secondY * second.Stride;
            int thirdRow = thirdY * third.Stride;
            int firstX = 0;
            int secondX = 0;
            int thirdX = 0;
            int firstXAccumulator = 0;
            int secondXAccumulator = 0;
            int thirdXAccumulator = 0;

            for (int x = 0; x < width; x++) {
                byte firstValue = ProjectSampleToByte(first, first.ReadSample(firstRow + firstX));
                byte secondValue = ProjectSampleToByte(second, second.ReadSample(secondRow + secondX));
                byte thirdValue = ProjectSampleToByte(third, third.ReadSample(thirdRow + thirdX));
                if (transformYccToRgb) {
                    int red = firstValue + CrToR[thirdValue];
                    int green = firstValue - CbToG[secondValue] - CrToG[thirdValue];
                    int blue = firstValue + CbToB[secondValue];
                    output[target++] = ClampToByte(red);
                    output[target++] = ClampToByte(green);
                    output[target++] = ClampToByte(blue);
                } else {
                    output[target++] = firstValue;
                    output[target++] = secondValue;
                    output[target++] = thirdValue;
                }
                if (outputRgba) output[target++] = byte.MaxValue;

                firstXAccumulator += first.Component.H;
                if (firstXAccumulator >= maximumHorizontalSampling) {
                    firstX++;
                    firstXAccumulator -= maximumHorizontalSampling;
                }
                secondXAccumulator += second.Component.H;
                if (secondXAccumulator >= maximumHorizontalSampling) {
                    secondX++;
                    secondXAccumulator -= maximumHorizontalSampling;
                }
                thirdXAccumulator += third.Component.H;
                if (thirdXAccumulator >= maximumHorizontalSampling) {
                    thirdX++;
                    thirdXAccumulator -= maximumHorizontalSampling;
                }
            }

            firstYAccumulator += first.Component.V;
            if (firstYAccumulator >= maximumVerticalSampling) {
                firstY++;
                firstYAccumulator -= maximumVerticalSampling;
            }
            secondYAccumulator += second.Component.V;
            if (secondYAccumulator >= maximumVerticalSampling) {
                secondY++;
                secondYAccumulator -= maximumVerticalSampling;
            }
            thirdYAccumulator += third.Component.V;
            if (thirdYAccumulator >= maximumVerticalSampling) {
                thirdY++;
                thirdYAccumulator -= maximumVerticalSampling;
            }
        }
    }

    private static void WriteGrayPixel(byte[] output, int pixel, byte gray, bool outputRgba) {
        int target = pixel * (outputRgba ? 4 : 1);
        output[target] = gray;
        if (!outputRgba) return;
        output[target + 1] = gray;
        output[target + 2] = gray;
        output[target + 3] = 255;
    }

    private static void WriteRgbPixel(byte[] output, int pixel, byte red, byte green, byte blue, bool outputRgba) {
        int target = pixel * (outputRgba ? 4 : 3);
        output[target] = red;
        output[target + 1] = green;
        output[target + 2] = blue;
        if (outputRgba) output[target + 3] = 255;
    }

    private static void WriteCmykPixel(byte[] output, int pixel, byte cyan, byte magenta, byte yellow, byte black, bool outputRgba) {
        int target = pixel * 4;
        if (outputRgba) {
            output[target] = ApplyCmyk(cyan, black);
            output[target + 1] = ApplyCmyk(magenta, black);
            output[target + 2] = ApplyCmyk(yellow, black);
            output[target + 3] = 255;
            return;
        }
        output[target] = cyan;
        output[target + 1] = magenta;
        output[target + 2] = yellow;
        output[target + 3] = black;
    }

    private static void YccToRgb(int y, int cb, int cr, out byte r, out byte g, out byte b) {
        var rVal = y + CrToR[cr];
        var gVal = y - CbToG[cb] - CrToG[cr];
        var bVal = y + CbToB[cb];
        r = ClampToByte(rVal);
        g = ClampToByte(gVal);
        b = ClampToByte(bVal);
    }

    private static int SampleComponent(
        BaselineComponentState[] states,
        int index,
        int x,
        int y,
        int maxH,
        int maxV,
        int fallback,
        bool highQualityChroma,
        bool preserveRaw16 = false) {
        if (index < 0 || index >= states.Length) return fallback;
        var state = states[index];
        if (!highQualityChroma || (state.Component.H == maxH && state.Component.V == maxV)) {
            var sx = x * state.Component.H / maxH;
            var sy = y * state.Component.V / maxV;
            var stride = state.Stride;
            int sample = state.ReadSample(sy * stride + sx);
            return preserveRaw16 ? sample : ProjectSampleToByte(state, sample);
        }

        int interpolated = SampleComponentBilinear(state, x, y, maxH, maxV);
        return preserveRaw16 ? interpolated : ProjectSampleToByte(state, interpolated);
    }

    private static byte ProjectSampleToByte(BaselineComponentState state, int sample) =>
        state.SampleMaximum == 255 ? (byte)sample :
            (byte)((sample * 255L + state.SampleMaximum / 2) / state.SampleMaximum);

    private static int SampleComponentBilinear(BaselineComponentState state, int x, int y, int maxH, int maxV) {
        var stride = state.Stride;
        var height = state.SampleCount / stride;

        var fx = (x + 0.5) * state.Component.H / maxH - 0.5;
        var fy = (y + 0.5) * state.Component.V / maxV - 0.5;

        var x0 = (int)Math.Floor(fx);
        var y0 = (int)Math.Floor(fy);
        var x1 = x0 + 1;
        var y1 = y0 + 1;

        if (x0 < 0) x0 = 0;
        if (y0 < 0) y0 = 0;
        if (x1 >= stride) x1 = stride - 1;
        if (y1 >= height) y1 = height - 1;

        var dx = fx - x0;
        var dy = fy - y0;

        var p00 = state.ReadSample(y0 * stride + x0);
        var p10 = state.ReadSample(y0 * stride + x1);
        var p01 = state.ReadSample(y1 * stride + x0);
        var p11 = state.ReadSample(y1 * stride + x1);

        var top = p00 + (p10 - p00) * dx;
        var bottom = p01 + (p11 - p01) * dx;
        var value = top + (bottom - top) * dy;
        return (int)Math.Round(value);
    }

    private static void DecodeBlockCoefficients(
        ref JpegBitReader reader,
        HuffmanTable dcTable,
        HuffmanTable acTable,
        int[] quant,
        ref int prevDc,
        int[] coeffs) {
        Array.Clear(coeffs, 0, 64);

        var t = DecodeHuffman(ref reader, dcTable, useFast: true);
        var diff = t == 0 ? 0 : Extend(reader.ReadBits(t), t);
        var dc = checked(prevDc + diff);
        prevDc = dc;
        coeffs[0] = checked(dc * quant[0]);

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
            coeffs[zig] = checked(ac * quant[zig]);
            k++;
        }

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

    private static void InverseDct(int[] input, byte[] output, int[] workspace) {

        for (int i = 0; i < 64; i++) {
            if (input[i] < -8191 || input[i] > 8191) { InverseDctWide(input, output); return; }
        }

        // Pass 1: process columns into the workspace (scaled by Pass1Bits).
        for (var ctr = 0; ctr < 8; ctr++) {
            var c0 = input[ctr];
            var c1 = input[ctr + 8];
            var c2 = input[ctr + 16];
            var c3 = input[ctr + 24];
            var c4 = input[ctr + 32];
            var c5 = input[ctr + 40];
            var c6 = input[ctr + 48];
            var c7 = input[ctr + 56];

            if (c1 == 0 && c2 == 0 && c3 == 0 && c4 == 0 && c5 == 0 && c6 == 0 && c7 == 0) {
                var dc = c0 << Pass1Bits;
                workspace[ctr] = dc;
                workspace[ctr + 8] = dc;
                workspace[ctr + 16] = dc;
                workspace[ctr + 24] = dc;
                workspace[ctr + 32] = dc;
                workspace[ctr + 40] = dc;
                workspace[ctr + 48] = dc;
                workspace[ctr + 56] = dc;
                continue;
            }

            long tmp0;
            long tmp1;
            long tmp2;
            long tmp3;
            long tmp10;
            long tmp11;
            long tmp12;
            long tmp13;
            long z1;
            long z2;
            long z3;
            long z4;
            long z5;

            // Even part.
            z2 = c2;
            z3 = c6;
            z1 = (z2 + z3) * Fix0_541196100;
            tmp2 = z1 + z3 * -Fix1_847759065;
            tmp3 = z1 + z2 * Fix0_765366865;

            tmp0 = ((long)c0 + c4) << ConstBits;
            tmp1 = ((long)c0 - c4) << ConstBits;

            tmp10 = tmp0 + tmp3;
            tmp13 = tmp0 - tmp3;
            tmp11 = tmp1 + tmp2;
            tmp12 = tmp1 - tmp2;

            // Odd part.
            tmp0 = c7;
            tmp1 = c5;
            tmp2 = c3;
            tmp3 = c1;

            z1 = tmp0 + tmp3;
            z2 = tmp1 + tmp2;
            z3 = tmp0 + tmp2;
            z4 = tmp1 + tmp3;
            z5 = (z3 + z4) * Fix1_175875602;

            tmp0 *= Fix0_298631336;
            tmp1 *= Fix2_053119869;
            tmp2 *= Fix3_072711026;
            tmp3 *= Fix1_501321110;
            z1 *= -Fix0_899976223;
            z2 *= -Fix2_562915447;
            z3 *= -Fix1_961570560;
            z4 *= -Fix0_390180644;

            z3 += z5;
            z4 += z5;

            tmp0 += z1 + z3;
            tmp1 += z2 + z4;
            tmp2 += z2 + z3;
            tmp3 += z1 + z4;

            workspace[ctr] = Descale(tmp10 + tmp3, ConstBits - Pass1Bits);
            workspace[ctr + 56] = Descale(tmp10 - tmp3, ConstBits - Pass1Bits);
            workspace[ctr + 8] = Descale(tmp11 + tmp2, ConstBits - Pass1Bits);
            workspace[ctr + 48] = Descale(tmp11 - tmp2, ConstBits - Pass1Bits);
            workspace[ctr + 16] = Descale(tmp12 + tmp1, ConstBits - Pass1Bits);
            workspace[ctr + 40] = Descale(tmp12 - tmp1, ConstBits - Pass1Bits);
            workspace[ctr + 24] = Descale(tmp13 + tmp0, ConstBits - Pass1Bits);
            workspace[ctr + 32] = Descale(tmp13 - tmp0, ConstBits - Pass1Bits);
        }

        // Pass 2: process rows from the workspace into final pixels.
        for (var ctr = 0; ctr < 8; ctr++) {
            var row = ctr * 8;
            var w0 = workspace[row];
            var w1 = workspace[row + 1];
            var w2 = workspace[row + 2];
            var w3 = workspace[row + 3];
            var w4 = workspace[row + 4];
            var w5 = workspace[row + 5];
            var w6 = workspace[row + 6];
            var w7 = workspace[row + 7];

            if (w1 == 0 && w2 == 0 && w3 == 0 && w4 == 0 && w5 == 0 && w6 == 0 && w7 == 0) {
                var dc = Descale(w0, Pass1Bits + 3) + 128;
                var clamped = ClampToByte(dc);
                output[row] = clamped;
                output[row + 1] = clamped;
                output[row + 2] = clamped;
                output[row + 3] = clamped;
                output[row + 4] = clamped;
                output[row + 5] = clamped;
                output[row + 6] = clamped;
                output[row + 7] = clamped;
                continue;
            }

            long tmp0;
            long tmp1;
            long tmp2;
            long tmp3;
            long tmp10;
            long tmp11;
            long tmp12;
            long tmp13;
            long z1;
            long z2;
            long z3;
            long z4;
            long z5;

            // Even part.
            z2 = w2;
            z3 = w6;
            z1 = (z2 + z3) * Fix0_541196100;
            tmp2 = z1 + z3 * -Fix1_847759065;
            tmp3 = z1 + z2 * Fix0_765366865;

            tmp0 = ((long)w0 + w4) << ConstBits;
            tmp1 = ((long)w0 - w4) << ConstBits;

            tmp10 = tmp0 + tmp3;
            tmp13 = tmp0 - tmp3;
            tmp11 = tmp1 + tmp2;
            tmp12 = tmp1 - tmp2;

            // Odd part.
            tmp0 = w7;
            tmp1 = w5;
            tmp2 = w3;
            tmp3 = w1;

            z1 = tmp0 + tmp3;
            z2 = tmp1 + tmp2;
            z3 = tmp0 + tmp2;
            z4 = tmp1 + tmp3;
            z5 = (z3 + z4) * Fix1_175875602;

            tmp0 *= Fix0_298631336;
            tmp1 *= Fix2_053119869;
            tmp2 *= Fix3_072711026;
            tmp3 *= Fix1_501321110;
            z1 *= -Fix0_899976223;
            z2 *= -Fix2_562915447;
            z3 *= -Fix1_961570560;
            z4 *= -Fix0_390180644;

            z3 += z5;
            z4 += z5;

            tmp0 += z1 + z3;
            tmp1 += z2 + z4;
            tmp2 += z2 + z3;
            tmp3 += z1 + z4;

            var shift = ConstBits + Pass1Bits + 3;
            output[row] = ClampToByte(Descale(tmp10 + tmp3, shift) + 128);
            output[row + 7] = ClampToByte(Descale(tmp10 - tmp3, shift) + 128);
            output[row + 1] = ClampToByte(Descale(tmp11 + tmp2, shift) + 128);
            output[row + 6] = ClampToByte(Descale(tmp11 - tmp2, shift) + 128);
            output[row + 2] = ClampToByte(Descale(tmp12 + tmp1, shift) + 128);
            output[row + 5] = ClampToByte(Descale(tmp12 - tmp1, shift) + 128);
            output[row + 3] = ClampToByte(Descale(tmp13 + tmp0, shift) + 128);
            output[row + 4] = ClampToByte(Descale(tmp13 - tmp0, shift) + 128);
        }
    }

    private static byte ClampToByte(int value) {
        if (value <= 0) return 0;
        if (value >= 255) return 255;
        return (byte)value;
    }

    private static int Descale(long value, int shift) {
        if (shift <= 0) return (int)value;
        var round = 1L << (shift - 1);
        if (value >= 0) {
            return (int)((value + round) >> shift);
        }
        return (int)(-(((-value) + round) >> shift));
    }

    private static int FindScanEnd(OfficeByteView data, int start, CancellationToken cancellationToken) {
        // Stuffed bytes and restart markers advance by two, so offsets can skip checkpoint boundaries.
        var i = start;
        int scanSteps = 0;
        while (i + 1 < data.Length) {
            if ((scanSteps++ & 0x3FFF) == 0) cancellationToken.ThrowIfCancellationRequested();
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

    private static byte ApplyCmyk(int c, int k) {
        var v = c + k;
        if (v > 255) v = 255;
        return (byte)(255 - v);
    }

}
