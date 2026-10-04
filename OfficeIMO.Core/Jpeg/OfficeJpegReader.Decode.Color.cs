using System;
using System.Threading;
using System.Threading.Tasks;
#if NET8_0_OR_GREATER
using System.Runtime.Intrinsics.X86;
#endif

namespace OfficeIMO.Drawing;

internal static partial class OfficeJpegReader {
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
        if ((long)width * height >= 262_144 && height >= 16 && Environment.ProcessorCount > 1) {
            Parallel.For(0, height, new ParallelOptions {
                CancellationToken = cancellationToken,
                MaxDegreeOfParallelism = Math.Min(8, Environment.ProcessorCount)
            }, y => ComposeThreeComponentNearestRow(output, width, y, first, second, third,
                maximumHorizontalSampling, maximumVerticalSampling, transformYccToRgb, outputRgba, cancellationToken));
        } else {
            for (int y = 0; y < height; y++) {
                ComposeThreeComponentNearestRow(output, width, y, first, second, third,
                    maximumHorizontalSampling, maximumVerticalSampling, transformYccToRgb, outputRgba, cancellationToken);
            }
        }
    }

    private static void ComposeThreeComponentNearestRow(
        byte[] output, int width, int y,
        BaselineComponentState first, BaselineComponentState second, BaselineComponentState third,
        int maximumHorizontalSampling, int maximumVerticalSampling,
        bool transformYccToRgb, bool outputRgba, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        int firstRow = (y * first.Component.V / maximumVerticalSampling) * first.Stride;
        int secondRow = (y * second.Component.V / maximumVerticalSampling) * second.Stride;
        int thirdRow = (y * third.Component.V / maximumVerticalSampling) * third.Stride;
        int target = y * width * (outputRgba ? 4 : 3);
        int firstX = 0;
        int secondX = 0;
        int thirdX = 0;
        int firstXAccumulator = 0;
        int secondXAccumulator = 0;
        int thirdXAccumulator = 0;
        int startX = 0;

#if NET8_0_OR_GREATER
        if (transformYccToRgb && outputRgba && width >= 8 && Avx2.IsSupported && Ssse3.IsSupported &&
            first.Component.H == maximumHorizontalSampling &&
            second.Component.H * 2 == maximumHorizontalSampling &&
            third.Component.H * 2 == maximumHorizontalSampling) {
            startX = ComposeYccToRgbaHalfChromaVector(first.Buffer, firstRow,
                second.Buffer, secondRow, third.Buffer, thirdRow, output, target, width);
            firstX = startX;
            secondX = thirdX = startX / 2;
            target += startX * 4;
        }
#endif

        for (int x = startX; x < width; x++) {
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

}
