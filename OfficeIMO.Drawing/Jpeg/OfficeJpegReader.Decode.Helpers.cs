using System;

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

    private static byte[] ComposeRgba(JpegFrame frame, BaselineComponentState[] states, int? adobeTransform, bool highQualityChroma) {
        var rgba = OfficeRasterGuards.AllocateRgba32(frame.Width, frame.Height, JpegDimensionsLimitMessage);
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

            var isYcck = adobeTransform == 2;
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
                for (var x = 0; x < frame.Width; x++) {
                    byte r;
                    byte g;
                    byte b;
                    var kVal = SampleComponent(states, isYcck ? ycckK : kIndex, x, y, maxH, maxV, 0, highQualityChroma);

                    if (isYcck) {
                        var yVal = SampleComponent(states, ycckY, x, y, maxH, maxV, 128, highQualityChroma);
                        var cbVal = SampleComponent(states, ycckCb, x, y, maxH, maxV, 128, highQualityChroma);
                        var crVal = SampleComponent(states, ycckCr, x, y, maxH, maxV, 128, highQualityChroma);
                        YccToRgb(yVal, cbVal, crVal, out r, out g, out b);
                    } else {
                        var c = SampleComponent(states, cIndex, x, y, maxH, maxV, 0, highQualityChroma);
                        var m = SampleComponent(states, mIndex, x, y, maxH, maxV, 0, highQualityChroma);
                        var y0 = SampleComponent(states, yIndex, x, y, maxH, maxV, 0, highQualityChroma);
                        r = ApplyCmyk(c, kVal);
                        g = ApplyCmyk(m, kVal);
                        b = ApplyCmyk(y0, kVal);
                    }

                    if (isYcck) {
                        var c = (byte)(255 - r);
                        var m = (byte)(255 - g);
                        var y0 = (byte)(255 - b);
                        r = ApplyCmyk(c, kVal);
                        g = ApplyCmyk(m, kVal);
                        b = ApplyCmyk(y0, kVal);
                    }

                    var p = (y * frame.Width + x) * 4;
                    rgba[p + 0] = r;
                    rgba[p + 1] = g;
                    rgba[p + 2] = b;
                    rgba[p + 3] = 255;
                }
            }

            return rgba;
        }

        if (frame.ComponentCount == 1) {
            var grayIndex = FindComponentIndex(frame.Components, 1);
            if (grayIndex < 0) grayIndex = 0;
            for (var y = 0; y < frame.Height; y++) {
                for (var x = 0; x < frame.Width; x++) {
                    var v = SampleComponent(states, grayIndex, x, y, maxH, maxV, 0, highQualityChroma);
                    var p = (y * frame.Width + x) * 4;
                    rgba[p + 0] = (byte)v;
                    rgba[p + 1] = (byte)v;
                    rgba[p + 2] = (byte)v;
                    rgba[p + 3] = 255;
                }
            }
            return rgba;
        }

        var rIndex = FindComponentIndex(frame.Components, (byte)'R');
        var gIndex = FindComponentIndex(frame.Components, (byte)'G');
        var bIndex = FindComponentIndex(frame.Components, (byte)'B');
        var rgb = rIndex >= 0 && gIndex >= 0 && bIndex >= 0;

        var yIndex2 = FindComponentIndex(frame.Components, 1);
        if (yIndex2 < 0) yIndex2 = 0;
        var cbIndex = frame.ComponentCount > 1 ? FindComponentIndex(frame.Components, 2) : -1;
        var crIndex = frame.ComponentCount > 1 ? FindComponentIndex(frame.Components, 3) : -1;
        if (frame.ComponentCount == 3) {
            if (cbIndex < 0) cbIndex = yIndex2 == 0 ? 1 : 0;
            if (crIndex < 0) crIndex = yIndex2 == 2 ? 1 : 2;
        }

        for (var y = 0; y < frame.Height; y++) {
            for (var x = 0; x < frame.Width; x++) {
                byte r;
                byte g;
                byte b;
                if (rgb) {
                    r = (byte)SampleComponent(states, rIndex, x, y, maxH, maxV, 0, highQualityChroma);
                    g = (byte)SampleComponent(states, gIndex, x, y, maxH, maxV, 0, highQualityChroma);
                    b = (byte)SampleComponent(states, bIndex, x, y, maxH, maxV, 0, highQualityChroma);
                } else {
                    var yVal = SampleComponent(states, yIndex2, x, y, maxH, maxV, 128, highQualityChroma);
                    var cbVal = SampleComponent(states, cbIndex, x, y, maxH, maxV, 128, highQualityChroma);
                    var crVal = SampleComponent(states, crIndex, x, y, maxH, maxV, 128, highQualityChroma);

                    YccToRgb(yVal, cbVal, crVal, out r, out g, out b);
                }

                var p = (y * frame.Width + x) * 4;
                rgba[p + 0] = r;
                rgba[p + 1] = g;
                rgba[p + 2] = b;
                rgba[p + 3] = 255;
            }
        }

        return rgba;
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
        int[] pixels) {
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

        InverseDct(coeffs, pixels);
    }

    private static int DecodeHuffman(ref JpegBitReader reader, HuffmanTable table, bool useFast) {
        if (useFast && table.Fast is not null && reader.HasBits(HuffmanFastBits)) {
            var peek = reader.PeekBits(HuffmanFastBits);
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

    private static void WriteBlock(int[] buffer, int stride, int blockX, int blockY, int[] pixels) {
        var baseX = blockX * 8;
        var baseY = blockY * 8;
        for (var y = 0; y < 8; y++) {
            var row = (baseY + y) * stride + baseX;
            var src = y * 8;
            for (var x = 0; x < 8; x++) {
                buffer[row + x] = pixels[src + x];
            }
        }
    }

    private static void InverseDct(int[] input, int[] output) {
        int[] workspace = new int[64];

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
        if (components != 1 && components != 3 && components != 4) {
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

    private static ScanHeader ParseScanHeader(OfficeByteView data, ref JpegFrame frame) {
        var components = data[0];
        if (components == 0) throw new FormatException("Invalid JPEG scan component count.");
        if (data.Length < 1 + components * 2 + 3) throw new FormatException("Invalid JPEG scan header.");

        var indices = new int[components];
        var offset = 1;
        for (var i = 0; i < components; i++) {
            var id = data[offset++];
            var table = data[offset++];
            var dc = table >> 4;
            var ac = table & 0x0F;
            var index = FindComponentIndex(frame.Components, id);
            if (index < 0) throw new FormatException("Unknown JPEG component in scan.");
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

    private static int FindScanEnd(OfficeByteView data, int start) {
        var i = start;
        while (i + 1 < data.Length) {
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

    private static bool TryReadExifOrientation(OfficeByteView data, out int orientation) {
        orientation = 1;
        if (data.Length < 6) return false;
        if (data[0] != (byte)'E' || data[1] != (byte)'x' || data[2] != (byte)'i' || data[3] != (byte)'f' || data[4] != 0 || data[5] != 0) {
            return false;
        }

        var tiff = data.Slice(6);
        if (tiff.Length < 8) return false;
        var little = tiff[0] == (byte)'I' && tiff[1] == (byte)'I';
        var big = tiff[0] == (byte)'M' && tiff[1] == (byte)'M';
        if (!little && !big) return false;
        if (ReadUInt16(tiff, 2, little) != 0x2A) return false;
        var ifdOffset = ReadUInt32(tiff, 4, little);
        if (ifdOffset > (uint)(tiff.Length - 2)) return false;
        var ifd = tiff.Slice((int)ifdOffset);
        var count = ReadUInt16(ifd, 0, little);
        var entriesOffset = 2;
        for (var i = 0; i < count; i++) {
            var entryOffset = entriesOffset + i * 12;
            if (entryOffset + 12 > ifd.Length) break;
            var tag = ReadUInt16(ifd, entryOffset, little);
            if (tag != 0x0112) continue;
            var type = ReadUInt16(ifd, entryOffset + 2, little);
            var entryCount = ReadUInt32(ifd, entryOffset + 4, little);
            if (type != 3 || entryCount != 1) break;
            var value = ReadUInt16(ifd, entryOffset + 8, little);
            if (value is >= 1 and <= 8) {
                orientation = value;
                return true;
            }
            break;
        }

        return false;
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

    private static ushort ReadUInt16(OfficeByteView data, int offset, bool little) {
        return little
            ? (ushort)(data[offset] | (data[offset + 1] << 8))
            : (ushort)((data[offset] << 8) | data[offset + 1]);
    }

    private static uint ReadUInt32(OfficeByteView data, int offset, bool little) {
        return little
            ? (uint)(data[offset]
                     | (data[offset + 1] << 8)
                     | (data[offset + 2] << 16)
                     | (data[offset + 3] << 24))
            : (uint)((data[offset] << 24)
                     | (data[offset + 1] << 16)
                     | (data[offset + 2] << 8)
                     | data[offset + 3]);
    }

    private static byte[] ApplyOrientation(byte[] rgba, ref int width, ref int height, int orientation) {
        if (orientation <= 1) return rgba;
        var srcWidth = width;
        var srcHeight = height;
        var destWidth = (orientation >= 5 && orientation <= 8) ? srcHeight : srcWidth;
        var destHeight = (orientation >= 5 && orientation <= 8) ? srcWidth : srcHeight;
        var result = OfficeRasterGuards.AllocateRgba32(destWidth, destHeight, JpegDimensionsLimitMessage);

        for (var y = 0; y < destHeight; y++) {
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

    private sealed class BaselineState {
        public BaselineComponentState[] Components = Array.Empty<BaselineComponentState>();
        public bool[] DecodedComponents = Array.Empty<bool>();
        public int McuCols;
        public int McuRows;

        public static BaselineState Create(JpegFrame frame) {
            var mcuWidth = frame.MaxH * 8;
            var mcuHeight = frame.MaxV * 8;
            var mcuCols = (frame.Width + mcuWidth - 1) / mcuWidth;
            var mcuRows = (frame.Height + mcuHeight - 1) / mcuHeight;
            var components = new BaselineComponentState[frame.ComponentCount];
            long aggregateBytes = checked((long)frame.Width * frame.Height * 4L);
            if (aggregateBytes > OfficeRasterGuards.MaximumDecodedBytes) {
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
                McuRows = mcuRows
            };
        }

        public byte[] RenderRgba(JpegFrame frame, int? adobeTransform, bool highQualityChroma) {
            for (var i = 0; i < DecodedComponents.Length; i++) {
                if (!DecodedComponents[i]) throw new FormatException("Missing JPEG component scan.");
            }

            return ComposeRgba(frame, Components, adobeTransform, highQualityChroma);
        }
    }

    private sealed class BaselineComponentState {
        public Component Component;
        public int[] Buffer;
        public int[] BlockCoeffs;
        public int[] BlockPixels;
        public int Stride;
        public int BlocksPerRow;
        public int BlocksPerCol;
        public int PrevDc;

        public BaselineComponentState(Component component, int blocksPerRow, int blocksPerCol, ref long aggregateBytes) {
            Component = component;
            BlocksPerRow = blocksPerRow;
            BlocksPerCol = blocksPerCol;
            Stride = OfficeRasterGuards.EnsureByteCount((long)blocksPerRow * 8, JpegDimensionsLimitMessage);
            var bufferLength = OfficeRasterGuards.EnsureInt32ArrayLength((long)Stride * blocksPerCol * 8, ref aggregateBytes, JpegDimensionsLimitMessage);
            Buffer = new int[bufferLength];
            BlockCoeffs = new int[OfficeRasterGuards.EnsureInt32ArrayLength(64, ref aggregateBytes, JpegDimensionsLimitMessage)];
            BlockPixels = new int[OfficeRasterGuards.EnsureInt32ArrayLength(64, ref aggregateBytes, JpegDimensionsLimitMessage)];
            PrevDc = 0;
        }

        public static BaselineComponentState FromDecodedBuffer(
            Component component,
            int blocksPerRow,
            int blocksPerCol,
            int stride,
            int[] buffer) {
            return new BaselineComponentState {
                Component = component,
                BlocksPerRow = blocksPerRow,
                BlocksPerCol = blocksPerCol,
                Stride = stride,
                Buffer = buffer,
                BlockCoeffs = Array.Empty<int>(),
                BlockPixels = Array.Empty<int>()
            };
        }

        private BaselineComponentState() {
            Buffer = Array.Empty<int>();
            BlockCoeffs = Array.Empty<int>();
            BlockPixels = Array.Empty<int>();
        }
    }

    private sealed class ProgressiveState {
        public ProgressiveComponentState[] Components = Array.Empty<ProgressiveComponentState>();
        public int McuCols;
        public int McuRows;

        public static ProgressiveState Create(JpegFrame frame, int[][] quantTables) {
            var maxH = frame.MaxH;
            var maxV = frame.MaxV;
            var mcuWidth = maxH * 8;
            var mcuHeight = maxV * 8;
            var mcuCols = (frame.Width + mcuWidth - 1) / mcuWidth;
            var mcuRows = (frame.Height + mcuHeight - 1) / mcuHeight;

            var components = new ProgressiveComponentState[frame.ComponentCount];
            long aggregateBytes = checked((long)frame.Width * frame.Height * 4L);
            if (aggregateBytes > OfficeRasterGuards.MaximumDecodedBytes) {
                throw new FormatException(JpegDimensionsLimitMessage);
            }
            for (var i = 0; i < frame.ComponentCount; i++) {
                var comp = frame.Components[i];
                if (comp.QuantId >= quantTables.Length || quantTables[comp.QuantId] is null) {
                    throw new FormatException("Missing JPEG quantization table.");
                }
                var blocksPerRow = OfficeRasterGuards.EnsureByteCount((long)mcuCols * comp.H, JpegDimensionsLimitMessage);
                var blocksPerCol = OfficeRasterGuards.EnsureByteCount((long)mcuRows * comp.V, JpegDimensionsLimitMessage);
                components[i] = new ProgressiveComponentState(comp, blocksPerRow, blocksPerCol, ref aggregateBytes);
            }

            return new ProgressiveState {
                Components = components,
                McuCols = mcuCols,
                McuRows = mcuRows
            };
        }

        public byte[] RenderRgba(JpegFrame frame, int? adobeTransform, bool highQualityChroma) {
            for (var i = 0; i < Components.Length; i++) {
                var compState = Components[i];
                for (var by = 0; by < compState.BlocksPerCol; by++) {
                    for (var bx = 0; bx < compState.BlocksPerRow; bx++) {
                        var baseIndex = (by * compState.BlocksPerRow + bx) * 64;
                        Array.Copy(compState.Coeffs, baseIndex, compState.BlockCoeffs, 0, 64);
                        InverseDct(compState.BlockCoeffs, compState.BlockPixels);
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

            return ComposeRgba(frame, baselineStates, adobeTransform, highQualityChroma);
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
        public int[] Coeffs;
        public int[] Buffer;
        public int[] BlockCoeffs;
        public int[] BlockPixels;
        public int Stride;
        public int PrevDc;

        public ProgressiveComponentState(Component component, int blocksPerRow, int blocksPerCol, ref long aggregateBytes) {
            Component = component;
            BlocksPerRow = blocksPerRow;
            BlocksPerCol = blocksPerCol;
            Stride = OfficeRasterGuards.EnsureByteCount((long)blocksPerRow * 8, JpegDimensionsLimitMessage);
            var coeffLength = OfficeRasterGuards.EnsureInt32ArrayLength((long)BlocksPerRow * BlocksPerCol * 64, ref aggregateBytes, JpegDimensionsLimitMessage);
            var bufferLength = OfficeRasterGuards.EnsureInt32ArrayLength((long)Stride * blocksPerCol * 8, ref aggregateBytes, JpegDimensionsLimitMessage);
            Coeffs = new int[coeffLength];
            Buffer = new int[bufferLength];
            BlockCoeffs = new int[OfficeRasterGuards.EnsureInt32ArrayLength(64, ref aggregateBytes, JpegDimensionsLimitMessage)];
            BlockPixels = new int[OfficeRasterGuards.EnsureInt32ArrayLength(64, ref aggregateBytes, JpegDimensionsLimitMessage)];
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

        public bool HasBits(int count) {
            return _bitCount >= count;
        }

        public int PeekBits(int count) {
            if (count == 0) return 0;
            EnsureBits(count);
            return (_bitBuffer >> (_bitCount - count)) & ((1 << count) - 1);
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
            while (_pos < _data.Length) {
                var b = _data[_pos++];
                if (b != 0xFF) return b;
                while (_pos < _data.Length && _data[_pos] == 0xFF) _pos++;
                if (_pos >= _data.Length) {
                    if (_allowTruncated) return 0;
                    throw new FormatException("Unexpected JPEG end.");
                }
                var marker = _data[_pos++];
                if (marker == 0x00) return 0xFF;
                if (marker >= 0xD0 && marker <= 0xD7) {
                    RestartMarkerSeen = true;
                    continue;
                }
                throw new FormatException("Unexpected JPEG marker in scan.");
            }
            if (_allowTruncated) return 0;
            throw new FormatException("Unexpected JPEG end.");
        }
    }

}
