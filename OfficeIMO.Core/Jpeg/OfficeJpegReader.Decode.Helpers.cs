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
        if (componentCount < 1 || componentCount > 4) {
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

        if (frame.ComponentCount == 4) {
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
                        var yVal = SampleComponent(states, ycckY, x, y, maxH, maxV, 128, highQualityChroma);
                        var cbVal = SampleComponent(states, ycckCb, x, y, maxH, maxV, 128, highQualityChroma);
                        var crVal = SampleComponent(states, ycckCr, x, y, maxH, maxV, 128, highQualityChroma);
                        YccToRgb(yVal, cbVal, crVal, out byte r, out byte g, out byte b);
                        if (adobeTransform.HasValue) {
                            c = (byte)(255 - r);
                            m = (byte)(255 - g);
                            y0 = (byte)(255 - b);
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
                        if (adobeTransform.HasValue) {
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

            if (!highQualityChroma) {
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
                    byte r;
                    byte g;
                    byte b;
                    var yVal = SampleComponent(states, yIndex2, x, y, maxH, maxV, 128, highQualityChroma);
                    var cbVal = SampleComponent(states, cbIndex, x, y, maxH, maxV, 128, highQualityChroma);
                    var crVal = SampleComponent(states, crIndex, x, y, maxH, maxV, 128, highQualityChroma);
                    YccToRgb(yVal, cbVal, crVal, out r, out g, out b);
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
                byte firstValue = first.Buffer[firstRow + firstX];
                byte secondValue = second.Buffer[secondRow + secondX];
                byte thirdValue = third.Buffer[thirdRow + thirdX];
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
        bool highQualityChroma) {
        if (index < 0 || index >= states.Length) return fallback;
        var state = states[index];
        if (!highQualityChroma || (state.Component.H == maxH && state.Component.V == maxV)) {
            var sx = x * state.Component.H / maxH;
            var sy = y * state.Component.V / maxV;
            var stride = state.Stride;
            return state.Buffer[sy * stride + sx];
        }

        return SampleComponentBilinear(state, x, y, maxH, maxV);
    }

    private static int SampleComponentBilinear(BaselineComponentState state, int x, int y, int maxH, int maxV) {
        var stride = state.Stride;
        var height = state.Buffer.Length / stride;

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

        var p00 = state.Buffer[y0 * stride + x0];
        var p10 = state.Buffer[y0 * stride + x1];
        var p01 = state.Buffer[y1 * stride + x0];
        var p11 = state.Buffer[y1 * stride + x1];

        var top = p00 + (p10 - p00) * dx;
        var bottom = p01 + (p11 - p01) * dx;
        var value = top + (bottom - top) * dy;
        return (int)Math.Round(value);
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

    private static void InverseDct(int[] input, byte[] output, int[] workspace) {

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

            tmp0 = (c0 + c4) << ConstBits;
            tmp1 = (c0 - c4) << ConstBits;

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

            tmp0 = (w0 + w4) << ConstBits;
            tmp1 = (w0 - w4) << ConstBits;

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
                while (j < data.Length && data[j] == 0xFF) j++;
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

    private static byte[] ApplyOrientation(
        byte[] rgba,
        ref int width,
        ref int height,
        int orientation,
        CancellationToken cancellationToken) {
        if (orientation <= 1) return rgba;
        var srcWidth = width;
        var srcHeight = height;
        var destWidth = (orientation >= 5 && orientation <= 8) ? srcHeight : srcWidth;
        var destHeight = (orientation >= 5 && orientation <= 8) ? srcWidth : srcHeight;
        var result = OfficeRasterGuards.AllocateRgba32(destWidth, destHeight, JpegDimensionsLimitMessage);

        for (var y = 0; y < destHeight; y++) {
            cancellationToken.ThrowIfCancellationRequested();
            for (var x = 0; x < destWidth; x++) {
                int sx;
                int sy;
                switch (orientation) {
                    case 2:
                        sx = srcWidth - 1 - x;
                        sy = y;
                        break;
                    case 3:
                        sx = srcWidth - 1 - x;
                        sy = srcHeight - 1 - y;
                        break;
                    case 4:
                        sx = x;
                        sy = srcHeight - 1 - y;
                        break;
                    case 5:
                        sx = y;
                        sy = x;
                        break;
                    case 6:
                        sx = y;
                        sy = srcHeight - 1 - x;
                        break;
                    case 7:
                        sx = srcWidth - 1 - y;
                        sy = srcHeight - 1 - x;
                        break;
                    case 8:
                        sx = srcWidth - 1 - y;
                        sy = x;
                        break;
                    default:
                        sx = x;
                        sy = y;
                        break;
                }

                var srcIndex = (sy * srcWidth + sx) * 4;
                var dstIndex = (y * destWidth + x) * 4;
                result[dstIndex + 0] = rgba[srcIndex + 0];
                result[dstIndex + 1] = rgba[srcIndex + 1];
                result[dstIndex + 2] = rgba[srcIndex + 2];
                result[dstIndex + 3] = rgba[srcIndex + 3];
            }
        }

        width = destWidth;
        height = destHeight;
        return result;
    }

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
        private int _pos;
        private int _bitBuffer;
        private int _bitCount;

        public bool RestartMarkerSeen;

        public JpegBitReader(OfficeByteView data, bool allowTruncated = false) {
            _data = data;
            _allowTruncated = allowTruncated;
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
                while (_pos < _data.Length && _data[_pos] == 0xFF) _pos++;
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
                while (_pos < _data.Length && _data[_pos] == 0xFF) _pos++;
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
