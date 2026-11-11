namespace OfficeIMO.Drawing;

internal static partial class OfficeJpeg2000Decoder {
    private const float IrreversibleAlpha = -1.586134342059924F;
    private const float IrreversibleBeta = -0.052980118572961F;
    private const float IrreversibleGamma = 0.882911075530934F;
    private const float IrreversibleDelta = 0.443506852043971F;
    private const float IrreversibleK = 1.230174104914001F;

    /// <summary>Reconstructed samples of one component at its own (possibly subsampled) resolution.</summary>
    private sealed class ComponentPlane {
        internal int Width, Height;
        internal int OriginX, OriginY; // component coordinates of the first sample
        internal int Precision;
        internal bool Signed;
        internal byte[]? Bytes;
        internal ushort[]? Words;

        internal int Get(int index) => Bytes != null ? Bytes[index] : Words![index];
    }

    private static ComponentPlane[] DecodeTiles(byte[] bytes, Codestream codestream, Budget budget) {
        var planes = new ComponentPlane[codestream.Components];
        for (int index = 0; index < planes.Length; index++) {
            int dx = codestream.SubsamplingX[index], dy = codestream.SubsamplingY[index];
            int originX = (int)CeilDiv(codestream.XOsiz, dx), originY = (int)CeilDiv(codestream.YOsiz, dy);
            var plane = new ComponentPlane {
                OriginX = originX,
                OriginY = originY,
                Width = (int)CeilDiv(codestream.Xsiz, dx) - originX,
                Height = (int)CeilDiv(codestream.Ysiz, dy) - originY,
                Precision = codestream.Precisions[index],
                Signed = codestream.Signed[index]
            };
            long samples = (long)plane.Width * plane.Height;
            budget.Reserve(plane.Precision > 8 ? samples * 2 : samples);
            if (plane.Precision > 8) plane.Words = new ushort[samples];
            else plane.Bytes = new byte[samples];
            // Missing tiles keep the level-shifted zero, the mid-grey of an unsigned component.
            if (!plane.Signed) {
                if (plane.Bytes != null) Fill(plane.Bytes, (byte)(1 << (plane.Precision - 1)), budget.Token);
                else Fill(plane.Words!, (ushort)(1 << (plane.Precision - 1)), budget.Token);
            } else {
                if (plane.Bytes != null) Fill(plane.Bytes, (byte)(1 << (plane.Precision - 1)), budget.Token);
                else Fill(plane.Words!, (ushort)(1 << (plane.Precision - 1)), budget.Token);
            }
            planes[index] = plane;
        }
        var blocks = new CodeBlockDecoder(budget);
        var tiles = new List<int>(codestream.Tiles.Keys);
        tiles.Sort();
        foreach (int tileIndex in tiles) {
            budget.Token.ThrowIfCancellationRequested();
            DecodeTile(bytes, codestream, codestream.Tiles[tileIndex], tileIndex, planes, blocks, budget);
        }
        return planes;
    }

    private static void Fill<T>(T[] array, T value, CancellationToken token) {
        for (int index = 0; index < array.Length; index++) {
            if ((index & 0xFFFF) == 0) token.ThrowIfCancellationRequested();
            array[index] = value;
        }
    }

    private static void DecodeTile(byte[] bytes, Codestream codestream, TileData tile, int tileIndex, ComponentPlane[] planes,
        CodeBlockDecoder blocks, Budget budget) {
        TileComponent[] components = CreateTileLayout(codestream, tile, tileIndex, budget,
            out GlobalCoding coding, out List<ProgressionChange> changes, out _, out _);
        TileStream stream = CreateTileStream(bytes, codestream, tile, coding, budget);
        DecodePackets(stream, components, coding, changes, budget);
        long reserved = 0;
        try {
            foreach (TileComponent component in components) {
                long samples = (long)component.Width * component.Height;
                if (samples <= 0) continue;
                budget.Reserve(samples * 4);
                reserved += samples * 4;
                if (component.Coding.Reversible) component.Integers = new int[samples];
                else component.Floats = new float[samples];
                DecodeCodeBlocks(stream.Data, component, blocks, budget);
                InverseWavelet(component, budget);
            }
            bool transform = coding.MultipleComponentTransform && components.Length >= 3;
            if (transform) ApplyInverseComponentTransform(components, budget);
            for (int index = 0; index < components.Length; index++)
                StoreComponent(components[index], planes[index], budget.Token);
        } finally {
            foreach (TileComponent component in components) {
                component.Integers = null;
                component.Floats = null;
            }
            budget.Release(reserved);
        }
    }

    private static void DecodeCodeBlocks(byte[] source, TileComponent component, CodeBlockDecoder blocks, Budget budget) {
        int stride = component.Width;
        foreach (Resolution resolution in component.Resolutions) {
            foreach (Precinct precinct in resolution.Precincts) {
                budget.Token.ThrowIfCancellationRequested();
                for (int b = 0; b < precinct.Bands.Length; b++) {
                    SubBand band = resolution.Bands[b];
                    foreach (CodeBlock block in precinct.Bands[b].Blocks) {
                        if (!blocks.Decode(block, band, component.Coding.Style, source)) continue;
                        int left = band.OffsetX + block.X0 - band.X0, top = band.OffsetY + block.Y0 - band.Y0;
                        if (component.Integers != null) blocks.WriteIntegers(component.Integers, stride, left, top, component.RegionShift);
                        else blocks.WriteFloats(component.Floats!, stride, left, top, component.RegionShift, band.StepSize);
                    }
                }
            }
        }
    }

    // Inverse discrete wavelet transformation (F.3): each level runs HOR_SR on rows, then VER_SR on columns.
    private static void InverseWavelet(TileComponent component, Budget budget) {
        int stride = component.Width;
        int longest = Math.Max(component.Width, component.Height);
        int[]? integerLine = component.Integers != null ? new int[longest + 8] : null;
        float[]? floatLine = component.Floats != null ? new float[longest + 8] : null;
        for (int r = 1; r < component.Resolutions.Length; r++) {
            Resolution resolution = component.Resolutions[r], lower = component.Resolutions[r - 1];
            int width = resolution.Width, height = resolution.Height;
            if (width <= 0 || height <= 0) continue;
            budget.Charge((long)width * height * 2);
            for (int y = 0; y < height; y++) {
                if ((y & 63) == 0) budget.Token.ThrowIfCancellationRequested();
                int row = y * stride;
                if (integerLine != null) {
                    Interleave(component.Integers!, row, 1, width, lower.Width, resolution.X0, integerLine);
                    Synthesize53(integerLine, width, resolution.X0);
                    Array.Copy(integerLine, 4, component.Integers!, row, width);
                } else {
                    Interleave(component.Floats!, row, 1, width, lower.Width, resolution.X0, floatLine!);
                    Synthesize97(floatLine!, width, resolution.X0);
                    Array.Copy(floatLine!, 4, component.Floats!, row, width);
                }
            }
            for (int x = 0; x < width; x++) {
                if ((x & 63) == 0) budget.Token.ThrowIfCancellationRequested();
                if (integerLine != null) {
                    Interleave(component.Integers!, x, stride, height, lower.Height, resolution.Y0, integerLine);
                    Synthesize53(integerLine, height, resolution.Y0);
                    for (int y = 0; y < height; y++) component.Integers![y * stride + x] = integerLine[4 + y];
                } else {
                    Interleave(component.Floats!, x, stride, height, lower.Height, resolution.Y0, floatLine!);
                    Synthesize97(floatLine!, height, resolution.Y0);
                    for (int y = 0; y < height; y++) component.Floats![y * stride + x] = floatLine![4 + y];
                }
            }
        }
    }

    // 2D_INTERLEAVE along one line: low-pass samples sit at even absolute coordinates.
    private static void Interleave<T>(T[] source, int start, int step, int count, int lowCount, int origin, T[] line) {
        int low = 0, high = 0;
        for (int index = 0; index < count; index++) {
            bool even = ((origin + index) & 1) == 0;
            line[4 + index] = even ? source[start + (low++) * step] : source[start + (lowCount + high++) * step];
        }
    }

    // Periodic symmetric extension (F.3.7).
    private static int Reflect(int index, int count) {
        int period = 2 * (count - 1);
        index %= period;
        if (index < 0) index += period;
        return index >= count ? period - index : index;
    }

    private static void Extend<T>(T[] line, int count, int padding) {
        for (int k = 1; k <= padding; k++) {
            line[4 - k] = line[4 + Reflect(-k, count)];
            line[4 + count - 1 + k] = line[4 + Reflect(count - 1 + k, count)];
        }
    }

    // 1D_SR for the reversible 5/3 filter (F.3.8.1).
    private static void Synthesize53(int[] line, int count, int origin) {
        if (count == 1) {
            if ((origin & 1) != 0) line[4] /= 2;
            return;
        }
        Extend(line, count, 2);
        int first = ((origin - 1) & 1) == 0 ? -1 : 0;
        for (int j = first; j <= count; j += 2) line[4 + j] -= (line[3 + j] + line[5 + j] + 2) >> 2;
        first = (origin & 1) != 0 ? 0 : 1;
        for (int j = first; j < count; j += 2) line[4 + j] += (line[3 + j] + line[5 + j]) >> 1;
    }

    // 1D_SR for the irreversible 9/7 filter (F.3.8.2).
    private static void Synthesize97(float[] line, int count, int origin) {
        if (count == 1) {
            if ((origin & 1) != 0) line[4] *= 0.5F;
            return;
        }
        Extend(line, count, 4);
        int evenStart = ((origin - 4) & 1) == 0 ? -4 : -3;
        int oddStart = evenStart == -4 ? -3 : -4;
        const float inverseK = 1F / IrreversibleK;
        for (int j = evenStart; j < count + 4; j += 2) line[4 + j] *= IrreversibleK;
        for (int j = oddStart; j < count + 4; j += 2) line[4 + j] *= inverseK;
        int even = ((origin - 3) & 1) == 0 ? -3 : -2;
        for (int j = even; j < count + 3; j += 2) line[4 + j] -= IrreversibleDelta * (line[3 + j] + line[5 + j]);
        int odd = ((origin - 2) & 1) != 0 ? -2 : -1;
        for (int j = odd; j < count + 2; j += 2) line[4 + j] -= IrreversibleGamma * (line[3 + j] + line[5 + j]);
        even = ((origin - 1) & 1) == 0 ? -1 : 0;
        for (int j = even; j < count + 1; j += 2) line[4 + j] -= IrreversibleBeta * (line[3 + j] + line[5 + j]);
        odd = (origin & 1) != 0 ? 0 : 1;
        for (int j = odd; j < count; j += 2) line[4 + j] -= IrreversibleAlpha * (line[3 + j] + line[5 + j]);
    }

    // Inverse RCT and ICT (G.2, G.3) on the first three components.
    private static void ApplyInverseComponentTransform(TileComponent[] components, Budget budget) {
        TileComponent y = components[0], cb = components[1], cr = components[2];
        if (y.Width != cb.Width || y.Width != cr.Width || y.Height != cb.Height || y.Height != cr.Height)
            throw new Jpeg2000Exception("The JPEG 2000 component transform requires equally sized components.");
        int count = y.Width * y.Height;
        budget.Charge(count);
        if (y.Integers != null && cb.Integers != null && cr.Integers != null) {
            int[] c0 = y.Integers, c1 = cb.Integers, c2 = cr.Integers;
            for (int index = 0; index < count; index++) {
                int green = c0[index] - ((c1[index] + c2[index]) >> 2);
                int red = c2[index] + green;
                int blue = c1[index] + green;
                c0[index] = red;
                c1[index] = green;
                c2[index] = blue;
            }
        } else if (y.Floats != null && cb.Floats != null && cr.Floats != null) {
            float[] c0 = y.Floats, c1 = cb.Floats, c2 = cr.Floats;
            for (int index = 0; index < count; index++) {
                float luma = c0[index], blueDifference = c1[index], redDifference = c2[index];
                c0[index] = luma + 1.402F * redDifference;
                c1[index] = luma - 0.34413F * blueDifference - 0.71414F * redDifference;
                c2[index] = luma + 1.772F * blueDifference;
            }
        } else if (count > 0) {
            throw new Jpeg2000Exception("The JPEG 2000 component transform mixes reversible and irreversible components.");
        }
    }

    // DC level shifting (G.1) and clamping into the component plane.
    private static void StoreComponent(TileComponent component, ComponentPlane plane, CancellationToken token) {
        int width = component.Width, height = component.Height;
        if (width <= 0 || height <= 0) return;
        int precision = plane.Precision;
        int offset = 1 << (precision - 1);
        int minimum = plane.Signed ? -offset : 0;
        int maximum = plane.Signed ? offset - 1 : (1 << precision) - 1;
        int shift = plane.Signed ? 0 : offset;
        int storageOffset = plane.Signed ? offset : 0;
        int left = component.X0 - plane.OriginX, top = component.Y0 - plane.OriginY;
        for (int y = 0; y < height; y++) {
            if ((y & 63) == 0) token.ThrowIfCancellationRequested();
            int planeY = top + y;
            if ((uint)planeY >= (uint)plane.Height) continue;
            int source = y * width;
            int target = planeY * plane.Width + left;
            for (int x = 0; x < width; x++) {
                if ((uint)(left + x) >= (uint)plane.Width) continue;
                int value = component.Integers != null
                    ? component.Integers[source + x]
                    : (int)Math.Floor(component.Floats![source + x] + 0.5);
                value += shift;
                value = value < minimum ? minimum : value > maximum ? maximum : value;
                value += storageOffset;
                if (plane.Bytes != null) plane.Bytes[target + x] = (byte)value;
                else plane.Words![target + x] = (ushort)value;
            }
        }
    }
}
