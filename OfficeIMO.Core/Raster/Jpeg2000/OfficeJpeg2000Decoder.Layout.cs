namespace OfficeIMO.Drawing;

internal static partial class OfficeJpeg2000Decoder {
    private const int CodeBlockBypass = 0x01, CodeBlockReset = 0x02, CodeBlockTerminateAll = 0x04,
        CodeBlockVerticallyCausal = 0x08, CodeBlockSegmentationSymbols = 0x20;

    private sealed class TileComponent {
        internal int Component;
        internal int X0, Y0, X1, Y1; // component tile coordinates (B.3)
        internal ComponentCoding Coding = new();
        internal Quantization Quantization = new();
        internal int RegionShift;
        internal Resolution[] Resolutions = Array.Empty<Resolution>();
        internal int Width => X1 - X0;
        internal int Height => Y1 - Y0;
    }

    private sealed class Resolution {
        internal int Level;
        internal int X0, Y0, X1, Y1;
        internal int PrecinctWidth, PrecinctHeight; // exponents
        internal int PrecinctOriginX, PrecinctOriginY; // first precinct index along each axis
        internal int PrecinctsWide, PrecinctsHigh;
        internal SubBand[] Bands = Array.Empty<SubBand>();
        internal Precinct[] Precincts = Array.Empty<Precinct>();
        internal int Width => X1 - X0;
        internal int Height => Y1 - Y0;
    }

    private sealed class SubBand {
        internal int Orientation; // 0 LL, 1 HL, 2 LH, 3 HH
        internal int X0, Y0, X1, Y1;
        internal int OffsetX, OffsetY; // placement in the tile-component coefficient buffer
        internal int CodeBlockWidth, CodeBlockHeight; // exponents
        internal int MagnitudeBits;
        internal float StepSize = 1F;
    }

    private sealed class Precinct {
        internal PrecinctBand[] Bands = Array.Empty<PrecinctBand>();
        internal int NextLayer;
        internal long PositionX, PositionY; // reference-grid position used by position-driven progressions
    }

    private sealed class PrecinctBand {
        internal int BlocksWide, BlocksHigh;
        internal TagTree? Inclusion;
        internal TagTree? ZeroPlanes;
        internal CodeBlock[] Blocks = Array.Empty<CodeBlock>();
    }

    private sealed class CodeBlock {
        internal int X0, Y0, X1, Y1; // band coordinates
        internal bool Included;
        internal int LengthBits = 3;
        internal int ZeroPlanes;
        internal int Passes;
        internal List<CodeSegment>? Segments;
    }

    private sealed class CodeSegment {
        internal int Passes;
        internal int MaximumPasses;
        internal int Length;
        internal List<(int Start, int Length)> Chunks = new(1);
    }

    private static TileComponent[] CreateTileLayout(Codestream codestream, TileData tile, int tileIndex, Budget budget,
        out GlobalCoding coding, out List<ProgressionChange> changes, out long tileX0, out long tileY0) {
        HeaderScope main = codestream.Main;
        HeaderScope? local = tile.Scope;
        coding = local?.Coding ?? main.Coding!;
        changes = local != null && local.ProgressionChanges.Count > 0 ? local.ProgressionChanges : main.ProgressionChanges;
        int column = tileIndex % codestream.TilesWide, row = tileIndex / codestream.TilesWide;
        long tx0 = Math.Max((long)codestream.XTOsiz + (long)column * codestream.XTsiz, codestream.XOsiz);
        long ty0 = Math.Max((long)codestream.YTOsiz + (long)row * codestream.YTsiz, codestream.YOsiz);
        long tx1 = Math.Min((long)codestream.XTOsiz + (long)(column + 1) * codestream.XTsiz, codestream.Xsiz);
        long ty1 = Math.Min((long)codestream.YTOsiz + (long)(row + 1) * codestream.YTsiz, codestream.Ysiz);
        tileX0 = tx0;
        tileY0 = ty0;
        var components = new TileComponent[codestream.Components];
        for (int index = 0; index < components.Length; index++) {
            ComponentCoding componentCoding =
                local != null && local.ComponentCodings.TryGetValue(index, out ComponentCoding? tileComponent) ? tileComponent :
                local?.Coding?.Component ??
                (main.ComponentCodings.TryGetValue(index, out ComponentCoding? mainComponent) ? mainComponent : main.Coding!.Component);
            Quantization quantization =
                local != null && local.ComponentQuantizations.TryGetValue(index, out Quantization? tileQuantization) ? tileQuantization :
                local?.Quantization ??
                (main.ComponentQuantizations.TryGetValue(index, out Quantization? mainQuantization) ? mainQuantization : main.Quantization!);
            int shift = local != null && local.RegionShifts.TryGetValue(index, out int tileShift) ? tileShift :
                main.RegionShifts.TryGetValue(index, out int mainShift) ? mainShift : 0;
            int dx = codestream.SubsamplingX[index], dy = codestream.SubsamplingY[index];
            var component = new TileComponent {
                Component = index,
                X0 = (int)CeilDiv(tx0, dx), Y0 = (int)CeilDiv(ty0, dy),
                X1 = (int)CeilDiv(tx1, dx), Y1 = (int)CeilDiv(ty1, dy),
                Coding = componentCoding,
                Quantization = quantization,
                RegionShift = shift
            };
            CreateResolutions(component, codestream.Precisions[index], dx, dy, tx0, ty0, budget);
            components[index] = component;
        }
        return components;
    }

    private static void CreateResolutions(TileComponent component, int precision, int dx, int dy, long tileX0, long tileY0, Budget budget) {
        int levels = component.Coding.Levels;
        component.Resolutions = new Resolution[levels + 1];
        for (int r = 0; r <= levels; r++) {
            int reduction = levels - r;
            var resolution = new Resolution {
                Level = r,
                X0 = (int)CeilDiv(component.X0, 1L << reduction), Y0 = (int)CeilDiv(component.Y0, 1L << reduction),
                X1 = (int)CeilDiv(component.X1, 1L << reduction), Y1 = (int)CeilDiv(component.Y1, 1L << reduction),
                PrecinctWidth = component.Coding.PrecinctWidths[r],
                PrecinctHeight = component.Coding.PrecinctHeights[r]
            };
            Resolution? lower = r == 0 ? null : component.Resolutions[r - 1];
            if (r == 0) {
                resolution.Bands = new[] { CreateBand(component, precision, 0, levels, resolution.X0, resolution.Y0, resolution.X1, resolution.Y1, 0, 0) };
            } else {
                int decomposition = levels - r + 1;
                resolution.Bands = new SubBand[3];
                for (int orientation = 1; orientation <= 3; orientation++) {
                    int xo = orientation & 1, yo = orientation >> 1;
                    long half = 1L << (decomposition - 1);
                    long scale = 1L << decomposition;
                    long x0 = CeilDiv(component.X0 - half * xo, scale), x1 = CeilDiv(component.X1 - half * xo, scale);
                    long y0 = CeilDiv(component.Y0 - half * yo, scale), y1 = CeilDiv(component.Y1 - half * yo, scale);
                    resolution.Bands[orientation - 1] = CreateBand(component, precision, orientation, decomposition,
                        (int)x0, (int)y0, (int)x1, (int)y1, xo == 1 ? lower!.Width : 0, yo == 1 ? lower!.Height : 0);
                }
                if (lower!.Width + resolution.Bands[0].X1 - resolution.Bands[0].X0 != resolution.Width ||
                    lower.Height + resolution.Bands[1].Y1 - resolution.Bands[1].Y0 != resolution.Height)
                    throw new Jpeg2000Exception("The JPEG 2000 subband geometry is inconsistent.");
            }
            CreatePrecincts(component, resolution, dx, dy, tileX0, tileY0, reduction, budget);
            component.Resolutions[r] = resolution;
        }
    }

    // Subband geometry (B.5) and quantization (E.1).
    private static SubBand CreateBand(TileComponent component, int precision, int orientation, int decomposition,
        int x0, int y0, int x1, int y1, int offsetX, int offsetY) {
        int levels = component.Coding.Levels;
        int bandIndex = orientation == 0 ? 0 : 3 * (levels - decomposition) + orientation;
        Quantization quantization = component.Quantization;
        int exponent, mantissa;
        if (quantization.Style == 1) {
            exponent = quantization.Exponents[0] - levels + decomposition;
            if (orientation == 0) exponent = quantization.Exponents[0];
            mantissa = quantization.Mantissas[0];
        } else {
            if (bandIndex >= quantization.Exponents.Length) throw new Jpeg2000Exception("The JPEG 2000 quantization marker lacks a subband.");
            exponent = quantization.Exponents[bandIndex];
            mantissa = quantization.Mantissas[bandIndex];
        }
        int gain = orientation == 0 ? 0 : orientation == 3 ? 2 : 1;
        int magnitudeBits = quantization.GuardBits + exponent - 1 + component.RegionShift;
        if (exponent < 0 || magnitudeBits > 30) throw new Jpeg2000Exception("JPEG 2000 coefficient magnitudes exceed 30 bits.");
        var band = new SubBand {
            Orientation = orientation,
            X0 = x0, Y0 = y0, X1 = x1, Y1 = y1,
            OffsetX = offsetX, OffsetY = offsetY,
            MagnitudeBits = magnitudeBits
        };
        if (!component.Coding.Reversible && quantization.Style != 0)
            band.StepSize = (float)(Math.Pow(2.0, precision + gain - exponent) * (1.0 + mantissa / 2048.0));
        return band;
    }

    // Precinct partition (B.6) and code-blocks (B.7).
    private static void CreatePrecincts(TileComponent component, Resolution resolution, int dx, int dy,
        long tileX0, long tileY0, int reduction, Budget budget) {
        int ppx = resolution.PrecinctWidth, ppy = resolution.PrecinctHeight;
        if (resolution.Width <= 0 || resolution.Height <= 0) return;
        resolution.PrecinctOriginX = (int)(resolution.X0 >> ppx);
        resolution.PrecinctOriginY = (int)(resolution.Y0 >> ppy);
        resolution.PrecinctsWide = (int)(CeilDiv(resolution.X1, 1L << ppx) - resolution.PrecinctOriginX);
        resolution.PrecinctsHigh = (int)(CeilDiv(resolution.Y1, 1L << ppy) - resolution.PrecinctOriginY);
        long count = (long)resolution.PrecinctsWide * resolution.PrecinctsHigh;
        budget.Reserve(count * 96L);
        resolution.Precincts = new Precinct[count];
        int bandShiftX = resolution.Level == 0 ? 0 : 1, bandShiftY = resolution.Level == 0 ? 0 : 1;
        int blockWidth = Math.Min(component.Coding.CodeBlockWidth, ppx - bandShiftX);
        int blockHeight = Math.Min(component.Coding.CodeBlockHeight, ppy - bandShiftY);
        if (blockWidth < 0 || blockHeight < 0) throw new Jpeg2000Exception("The JPEG 2000 precinct is smaller than a code-block.");
        foreach (SubBand band in resolution.Bands) {
            band.CodeBlockWidth = blockWidth;
            band.CodeBlockHeight = blockHeight;
        }
        for (int py = 0; py < resolution.PrecinctsHigh; py++) {
            for (int px = 0; px < resolution.PrecinctsWide; px++) {
                budget.Token.ThrowIfCancellationRequested();
                long resolutionX = (long)(resolution.PrecinctOriginX + px) << ppx;
                long resolutionY = (long)(resolution.PrecinctOriginY + py) << ppy;
                var precinct = new Precinct {
                    Bands = new PrecinctBand[resolution.Bands.Length],
                    PositionX = Math.Max(tileX0, dx * (resolutionX << reduction)),
                    PositionY = Math.Max(tileY0, dy * (resolutionY << reduction))
                };
                for (int b = 0; b < resolution.Bands.Length; b++) {
                    SubBand band = resolution.Bands[b];
                    long x0 = Math.Max(band.X0, resolutionX >> bandShiftX);
                    long y0 = Math.Max(band.Y0, resolutionY >> bandShiftY);
                    long x1 = Math.Min(band.X1, (resolutionX + (1L << ppx)) >> bandShiftX);
                    long y1 = Math.Min(band.Y1, (resolutionY + (1L << ppy)) >> bandShiftY);
                    var precinctBand = new PrecinctBand();
                    if (x1 > x0 && y1 > y0) {
                        long firstX = x0 >> blockWidth, firstY = y0 >> blockHeight;
                        int wide = (int)(CeilDiv(x1, 1L << blockWidth) - firstX);
                        int high = (int)(CeilDiv(y1, 1L << blockHeight) - firstY);
                        long blocks = (long)wide * high;
                        budget.Reserve(blocks * 128L);
                        budget.Charge(blocks);
                        precinctBand.BlocksWide = wide;
                        precinctBand.BlocksHigh = high;
                        precinctBand.Blocks = new CodeBlock[blocks];
                        for (int by = 0; by < high; by++) {
                            for (int bx = 0; bx < wide; bx++) {
                                long cx = (firstX + bx) << blockWidth, cy = (firstY + by) << blockHeight;
                                precinctBand.Blocks[by * wide + bx] = new CodeBlock {
                                    X0 = (int)Math.Max(cx, x0), Y0 = (int)Math.Max(cy, y0),
                                    X1 = (int)Math.Min(cx + (1L << blockWidth), x1), Y1 = (int)Math.Min(cy + (1L << blockHeight), y1)
                                };
                            }
                        }
                        precinctBand.Inclusion = new TagTree(wide, high);
                        precinctBand.ZeroPlanes = new TagTree(wide, high);
                    }
                    precinct.Bands[b] = precinctBand;
                }
                resolution.Precincts[py * resolution.PrecinctsWide + px] = precinct;
            }
        }
    }

    private static long CeilDiv(long value, long divisor) => value >= 0 ? (value + divisor - 1) / divisor : -(-value / divisor);

    /// <summary>Tag tree (B.10.2) with lazily discovered node values.</summary>
    private sealed class TagTree {
        private readonly int[] _values;
        private readonly int[] _lows;
        private readonly int[] _parents;
        private readonly int[] _path = new int[34];

        internal TagTree(int width, int height) {
            int total = 0;
            int w = width, h = height;
            while (true) {
                total += w * h;
                if (w == 1 && h == 1) break;
                w = (w + 1) >> 1;
                h = (h + 1) >> 1;
            }
            _values = new int[total];
            _lows = new int[total];
            _parents = new int[total];
            for (int index = 0; index < total; index++) _values[index] = int.MaxValue;
            int levelStart = 0;
            w = width;
            h = height;
            while (!(w == 1 && h == 1)) {
                int parentWidth = (w + 1) >> 1, parentHeight = (h + 1) >> 1;
                int parentStart = levelStart + w * h;
                for (int y = 0; y < h; y++) {
                    for (int x = 0; x < w; x++) _parents[levelStart + y * w + x] = parentStart + (y >> 1) * parentWidth + (x >> 1);
                }
                levelStart = parentStart;
                w = parentWidth;
                h = parentHeight;
            }
            _parents[total - 1] = -1;
        }

        /// <summary>Returns whether the leaf value is below <paramref name="threshold"/>.</summary>
        internal bool Decode(PacketHeaderReader reader, int leaf, int threshold) {
            int depth = 0;
            for (int node = leaf; node >= 0; node = _parents[node]) _path[depth++] = node;
            int low = 0;
            for (int index = depth - 1; index >= 0; index--) {
                int node = _path[index];
                if (low > _lows[node]) _lows[node] = low;
                else low = _lows[node];
                while (low < threshold && low < _values[node]) {
                    if (reader.ReadBit() != 0) _values[node] = low;
                    else low++;
                }
                _lows[node] = low;
            }
            return _values[leaf] < threshold;
        }

        internal int DecodeValue(PacketHeaderReader reader, int leaf, int maximum) {
            int threshold = 1;
            while (!Decode(reader, leaf, threshold)) {
                if (++threshold > maximum) throw new Jpeg2000Exception("A JPEG 2000 tag-tree value is out of range.");
            }
            return _values[leaf];
        }
    }
}
