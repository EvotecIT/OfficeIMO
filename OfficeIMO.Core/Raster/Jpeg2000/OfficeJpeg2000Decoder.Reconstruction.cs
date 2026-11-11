namespace OfficeIMO.Drawing;

internal static partial class OfficeJpeg2000Decoder {
    /// <summary>Full-resolution rows reconstructed, colour transformed and stored per step.</summary>
    private const int StripRows = 32;

    /// <summary>Reconstructed samples of one component at its own (possibly subsampled) resolution.</summary>
    private sealed class ComponentPlane {
        internal int Width, Height;
        internal int OriginX, OriginY; // component coordinates of the first sample
        internal int Precision;
        internal bool Signed;
        internal byte[]? Bytes;
        internal ushort[]? Words;

        // Direct output: stored values go straight through these lookups into interleaved output channels.
        internal byte[]? Output;
        internal int OutputChannels;
        internal int[] OutputIndexes = Array.Empty<int>();
        internal byte[][] OutputLookups = Array.Empty<byte[]>();
    }

    private static ComponentPlane[] CreatePlanes(Codestream codestream) {
        var planes = new ComponentPlane[codestream.Components];
        for (int index = 0; index < planes.Length; index++) {
            int dx = codestream.SubsamplingX[index], dy = codestream.SubsamplingY[index];
            int originX = (int)CeilDiv(codestream.XOsiz, dx), originY = (int)CeilDiv(codestream.YOsiz, dy);
            planes[index] = new ComponentPlane {
                OriginX = originX,
                OriginY = originY,
                Width = (int)CeilDiv(codestream.Xsiz, dx) - originX,
                Height = (int)CeilDiv(codestream.Ysiz, dy) - originY,
                Precision = codestream.Precisions[index],
                Signed = codestream.Signed[index]
            };
        }
        return planes;
    }

    private static void AllocatePlane(ComponentPlane plane, Budget budget) {
        long samples = (long)plane.Width * plane.Height;
        budget.Reserve(plane.Precision > 8 ? samples * 2 : samples);
        // Missing tiles keep the level-shifted zero, the mid-grey of an unsigned component.
        if (plane.Precision > 8) {
            plane.Words = new ushort[samples];
            Fill(plane.Words, (ushort)(1 << (plane.Precision - 1)), budget.Token);
        } else {
            plane.Bytes = new byte[samples];
            Fill(plane.Bytes, (byte)(1 << (plane.Precision - 1)), budget.Token);
        }
    }

    private static void Fill<T>(T[] array, T value, CancellationToken token) {
        for (int index = 0; index < array.Length; index++) {
            if ((index & 0xFFFF) == 0) token.ThrowIfCancellationRequested();
            array[index] = value;
        }
    }

    private static void DecodeTiles(byte[] bytes, Codestream codestream, ComponentPlane[] planes, Budget budget) {
        var workspace = new Workspace(budget, new CodeBlockDecoder(budget));
        var tiles = new List<int>(codestream.Tiles.Keys);
        tiles.Sort();
        foreach (int tileIndex in tiles) {
            budget.Token.ThrowIfCancellationRequested();
            DecodeTile(bytes, codestream, codestream.Tiles[tileIndex], tileIndex, planes, workspace);
        }
    }

    // Each tile-component is reconstructed in strips of full-resolution rows; only the code-block rows and wavelet
    // rows a strip needs are held, never a whole coefficient plane.
    private static void DecodeTile(byte[] bytes, Codestream codestream, TileData tile, int tileIndex, ComponentPlane[] planes,
        Workspace workspace) {
        Budget budget = workspace.Budget;
        TileComponent[] components = CreateTileLayout(codestream, tile, tileIndex, budget,
            out GlobalCoding coding, out List<ProgressionChange> changes, out _, out _);
        TileStream stream = CreateTileStream(bytes, codestream, tile, coding, budget);
        DecodePackets(stream, components, coding, changes, budget);
        workspace.Source = stream.Data;
        var sources = new RowSource?[components.Length];
        try {
            for (int index = 0; index < components.Length; index++) {
                TileComponent component = components[index];
                if ((long)component.Width * component.Height > 0) sources[index] = CreateComponentSource(component, workspace);
            }
            int first = 0;
            if (coding.MultipleComponentTransform && components.Length >= 3) {
                ReconstructTransformed(components, sources, planes, workspace);
                first = 3;
            }
            for (int index = first; index < components.Length; index++) {
                RowSource? source = sources[index];
                if (source == null) continue;
                for (int row = 0; row < source.Height; row += StripRows) {
                    int end = Math.Min(source.Height, row + StripRows);
                    source.Ensure(row, end - 1, workspace);
                    StoreRows(components[index], source, row, end, planes[index], budget.Token);
                }
            }
        } finally {
            foreach (RowSource? source in sources) source?.Return(workspace.Pool);
            workspace.Source = Array.Empty<byte>();
        }
    }

    // Inverse RCT and ICT (G.2, G.3) on the first three components, strip by strip.
    private static void ReconstructTransformed(TileComponent[] components, RowSource?[] sources, ComponentPlane[] planes, Workspace workspace) {
        Budget budget = workspace.Budget;
        TileComponent y = components[0], cb = components[1], cr = components[2];
        bool equalSizes = y.Width == cb.Width && y.Width == cr.Width && y.Height == cb.Height && y.Height == cr.Height;
        bool mixed = equalSizes && (long)y.Width * y.Height > 0 &&
            (y.Coding.Reversible != cb.Coding.Reversible || y.Coding.Reversible != cr.Coding.Reversible);
        if (!equalSizes || mixed) {
            // Reconstruct the components first so that coding failures and budgets take precedence, in the same
            // order as a decoder that transforms whole tiles.
            for (int index = 0; index < sources.Length; index++) {
                RowSource? source = sources[index];
                if (source == null) continue;
                for (int row = 0; row < source.Height; row += StripRows) source.Ensure(row, Math.Min(source.Height, row + StripRows) - 1, workspace);
            }
            if (!equalSizes) throw new Jpeg2000Exception("The JPEG 2000 component transform requires equally sized components.");
            budget.Charge(y.Width * y.Height);
            throw new Jpeg2000Exception("The JPEG 2000 component transform mixes reversible and irreversible components.");
        }
        int height = y.Height, width = y.Width;
        budget.Charge(width * height);
        if ((long)width * height <= 0) return;
        RowSource s0 = sources[0]!, s1 = sources[1]!, s2 = sources[2]!;
        for (int row = 0; row < height; row += StripRows) {
            int end = Math.Min(height, row + StripRows);
            s0.Ensure(row, end - 1, workspace);
            s1.Ensure(row, end - 1, workspace);
            s2.Ensure(row, end - 1, workspace);
            for (int line = row; line < end; line++) {
                int o0 = s0.Offset(line), o1 = s1.Offset(line), o2 = s2.Offset(line);
                if (s0.Ints != null) {
                    int[] c0 = s0.Ints, c1 = s1.Ints!, c2 = s2.Ints!;
                    for (int x = 0; x < width; x++) {
                        int green = c0[o0 + x] - ((c1[o1 + x] + c2[o2 + x]) >> 2);
                        int red = c2[o2 + x] + green;
                        int blue = c1[o1 + x] + green;
                        c0[o0 + x] = red;
                        c1[o1 + x] = green;
                        c2[o2 + x] = blue;
                    }
                } else {
                    float[] c0 = s0.Floats!, c1 = s1.Floats!, c2 = s2.Floats!;
                    for (int x = 0; x < width; x++) {
                        float luma = c0[o0 + x], blueDifference = c1[o1 + x], redDifference = c2[o2 + x];
                        c0[o0 + x] = luma + 1.402F * redDifference;
                        c1[o1 + x] = luma - 0.34413F * blueDifference - 0.71414F * redDifference;
                        c2[o2 + x] = luma + 1.772F * blueDifference;
                    }
                }
            }
            StoreRows(y, s0, row, end, planes[0], budget.Token);
            StoreRows(cb, s1, row, end, planes[1], budget.Token);
            StoreRows(cr, s2, row, end, planes[2], budget.Token);
        }
    }

    // DC level shifting (G.1) and clamping into the component plane or straight into the output channels.
    private static void StoreRows(TileComponent component, RowSource source, int first, int end, ComponentPlane plane, CancellationToken token) {
        int width = component.Width;
        if (width <= 0) return;
        token.ThrowIfCancellationRequested();
        if (plane.Output == null && plane.Bytes == null && plane.Words == null) return;
        int precision = plane.Precision;
        int offset = 1 << (precision - 1);
        int minimum = plane.Signed ? -offset : 0;
        int maximum = plane.Signed ? offset - 1 : (1 << precision) - 1;
        int shift = plane.Signed ? 0 : offset;
        int storageOffset = plane.Signed ? offset : 0;
        int left = component.X0 - plane.OriginX, top = component.Y0 - plane.OriginY;
        int from = Math.Max(0, -left), to = Math.Min(width, plane.Width - left);
        if (from >= to) return;
        int[]? ints = source.Ints;
        float[]? floats = source.Floats;
        byte[]? output = plane.Output;
        int channels = plane.OutputChannels;
        int[] indexes = plane.OutputIndexes;
        byte[][] lookups = plane.OutputLookups;
        for (int y = first; y < end; y++) {
            int planeY = top + y;
            if ((uint)planeY >= (uint)plane.Height) continue;
            int sourceRow = source.Offset(y);
            int target = planeY * plane.Width + left;
            for (int x = from; x < to; x++) {
                int value = ints != null ? ints[sourceRow + x] : (int)Math.Floor(floats![sourceRow + x] + 0.5);
                value += shift;
                value = value < minimum ? minimum : value > maximum ? maximum : value;
                value += storageOffset;
                if (output != null) {
                    int pixel = (target + x) * channels;
                    for (int t = 0; t < indexes.Length; t++) output[pixel + indexes[t]] = lookups[t][value];
                } else if (plane.Bytes != null) {
                    plane.Bytes[target + x] = (byte)value;
                } else {
                    plane.Words![target + x] = (ushort)value;
                }
            }
        }
    }
}
