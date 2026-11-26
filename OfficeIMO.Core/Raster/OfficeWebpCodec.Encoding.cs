namespace OfficeIMO.Drawing;

public static partial class OfficeWebpCodec {
    private const int Vp8lCompressionMaximumPixels = 4_000_000;
    private static readonly int[] Vp8lPredictorModes = { 11, 12, 13, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 0 };

    /// <summary>How much search each effort spends; every setting changes only the size of the stream.</summary>
    private sealed class Vp8lEncoderSettings {
        internal int ColorBits { get; private set; }
        internal int ColorSampleStep { get; private set; } = 1;
        internal int MatchAttempts { get; private set; }
        // Zero keeps the greedy parse; otherwise the number of cost-based passes.
        internal int ParsePasses { get; private set; }
        internal bool TryBothParseStarts { get; private set; }
        internal int MaximumCacheBits { get; private set; }
        // Fast compares only no cache against the largest one.
        internal bool QuickCacheChoice { get; private set; }
        internal int EntropyTiles { get; private set; }
        internal int ExactClusters { get; private set; } = 64;
        // Layouts compared over the whole image; with a row sample, its best single predictor joins them.
        internal Vp8lLayout[] Layouts { get; private set; } = { Vp8lLayout.PerTile, Vp8lLayout.Untransformed };
        internal int TrialPixels { get; private set; }
        internal int TrialAttempts { get; private set; } = 8;
        internal bool CompareOnRowSample { get; private set; }

        internal static Vp8lEncoderSettings For(OfficeImageEncodingEffort effort) => effort switch {
            OfficeImageEncodingEffort.Fast => new Vp8lEncoderSettings {
                MatchAttempts = 4, MaximumCacheBits = 8, QuickCacheChoice = true,
                Layouts = Vp8lFastLayouts, TrialAttempts = 1, CompareOnRowSample = true
            },
            OfficeImageEncodingEffort.Smallest => new Vp8lEncoderSettings {
                ColorBits = 6, MatchAttempts = 256, ParsePasses = 2, TryBothParseStarts = true,
                MaximumCacheBits = Vp8lMaximumCacheBits, EntropyTiles = 2048, ExactClusters = 192,
                TrialPixels = 1 << 20
            },
            _ => new Vp8lEncoderSettings {
                ColorBits = 6, ColorSampleStep = 2, MatchAttempts = 16, ParsePasses = 2,
                MaximumCacheBits = Vp8lMaximumCacheBits, EntropyTiles = 512, TrialPixels = 1 << 17
            }
        };
    }

    /// <summary>How the main image is predicted: one mode everywhere (the value), a mode per tile, or not at all.</summary>
    private enum Vp8lLayout {
        Left = 1,
        Top = 2,
        TopRight = 3,
        TopLeft = 4,
        Select = 11,
        Gradient = 12,
        HalfGradient = 13,
        PerTile = 100,
        Untransformed = 101
    }

    // Single-mode predictors compared on a row sample; the best of them meets the full-image candidates.
    private static readonly Vp8lLayout[] Vp8lSampleLayouts = {
        Vp8lLayout.Left, Vp8lLayout.Top, Vp8lLayout.TopRight, Vp8lLayout.TopLeft,
        Vp8lLayout.Select, Vp8lLayout.Gradient, Vp8lLayout.HalfGradient
    };
    private static readonly Vp8lLayout[] Vp8lFastLayouts = { Vp8lLayout.Select, Vp8lLayout.Untransformed };

    /// <summary>The residuals of one layout over the whole image, with its predictor map and subtract-green choice.</summary>
    private sealed class Vp8lPrediction {
        internal Vp8lPrediction(Vp8lLayout layout, uint[] residuals, uint[] map, bool subtractGreen) {
            Layout = layout;
            Residuals = residuals;
            Map = map;
            SubtractGreen = subtractGreen;
        }

        internal Vp8lLayout Layout { get; }
        internal uint[] Residuals { get; }
        internal uint[] Map { get; }
        internal bool SubtractGreen { get; }
        // Map tile size; one predictor everywhere uses the largest tiles, so its map is a handful of pixels.
        internal int Bits => Layout == Vp8lLayout.PerTile ? Vp8lPerTileBits : 9;
    }

    private const int Vp8lPerTileBits = 3;

    private static Vp8lPrediction PredictVp8lLayout(uint[] pixels, int width, int height, Vp8lLayout layout,
        CancellationToken cancellationToken) {
        var residuals = new uint[pixels.Length];
        if (layout == Vp8lLayout.Untransformed) {
            bool rawGreen = ShouldSubtractVp8lGreenFromPixels(pixels, cancellationToken);
            for (int index = 0; index < pixels.Length; index++) {
                if ((index & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
                residuals[index] = rawGreen ? SubtractGreenFromResidual(pixels[index]) : pixels[index];
            }
            return new Vp8lPrediction(layout, residuals, Array.Empty<uint>(), rawGreen);
        }
        bool subtractGreen = ShouldSubtractVp8lGreen(pixels, width, height, cancellationToken);
        uint[] map = layout == Vp8lLayout.PerTile
            ? SelectVp8lPredictors(pixels, residuals, width, height, Vp8lPerTileBits, Vp8lPredictorModes,
                subtractGreen, cancellationToken)
            : ApplyVp8lUniformPredictor(pixels, residuals, width, height, 9, (int)layout, subtractGreen,
                cancellationToken);
        return new Vp8lPrediction(layout, residuals, map, subtractGreen);
    }

    // A quick parse's estimate. A costed parse can fall back to coding everything as literals, so for it
    // the all-literal estimate counts too when that is cheaper; the greedy parse keeps its copies.
    private static double EstimateVp8lLayoutBits(uint[] residuals, int width, int attempts, int maximumCacheBits,
        bool costedParse, CancellationToken cancellationToken) {
        Vp8lTokens tokens = CreateVp8lTokens(residuals, width, attempts, cancellationToken);
        SelectVp8lCacheBits(tokens, residuals, maximumCacheBits, cancellationToken, out double parsedBits, onlyLargest: true);
        if (!costedParse) return parsedBits;
        SelectVp8lCacheBits(null, residuals, maximumCacheBits, cancellationToken, out double literalBits, onlyLargest: true);
        return Math.Min(parsedBits, literalBits);
    }

    /// <summary>
    /// Chooses how to predict the image by the estimated size of a quick parse. Residual entropy alone favours
    /// the mode with the smallest residuals, but simple modes often keep repeats intact that copies then code
    /// almost for free, and screenshots and document pages often code best unpredicted. Single-mode predictors
    /// are ranked on bands of rows; the best of them, a predictor per tile and the unpredicted pixels are then
    /// compared over the whole image, because a page repeats its glyphs and shapes far beyond any sample.
    /// </summary>
    private static Vp8lPrediction ChooseVp8lPrediction(uint[] pixels, int width, int height, Vp8lEncoderSettings settings,
        CancellationToken cancellationToken) {
        // Unpredicted pixels first, then the best single predictor, then a predictor per tile.
        var candidates = new List<Vp8lLayout>();
        if (Array.IndexOf(settings.Layouts, Vp8lLayout.Untransformed) >= 0) candidates.Add(Vp8lLayout.Untransformed);
        foreach (Vp8lLayout layout in settings.Layouts) {
            if (layout != Vp8lLayout.Untransformed && layout != Vp8lLayout.PerTile) candidates.Add(layout);
        }
        if (settings.TrialPixels > 0) candidates.Add(ChooseVp8lSampleLayout(pixels, width, height, settings, cancellationToken));
        bool perTile = Array.IndexOf(settings.Layouts, Vp8lLayout.PerTile) >= 0;
        if (perTile) candidates.Add(Vp8lLayout.PerTile);
        if (candidates.Count == 1) return PredictVp8lLayout(pixels, width, height, candidates[0], cancellationToken);

        // Fast compares on every third band of rows and predicts only the winner over the whole image.
        uint[] compared = pixels;
        int comparedRows = height;
        if (settings.CompareOnRowSample && height >= 96) {
            const int bandRows = 16;
            const int period = 3;
            int bands = height / (bandRows * period);
            compared = new uint[bands * bandRows * width];
            for (int band = 0; band < bands; band++) {
                Array.Copy(pixels, band * bandRows * period * width, compared, band * bandRows * width, bandRows * width);
            }
            comparedRows = bands * bandRows;
        }
        Vp8lPrediction? best = null;
        double bestBits = double.MaxValue, untransformedBits = double.MaxValue, singleBits = double.MaxValue;
        foreach (Vp8lLayout layout in candidates) {
            cancellationToken.ThrowIfCancellationRequested();
            if (layout == Vp8lLayout.PerTile && untransformedBits <= singleBits * 0.85) {
                // Per-tile prediction gains a few percent over the best single predictor, rarely the
                // 15% it would need to beat the unpredicted pixels here.
                break;
            }
            Vp8lPrediction prediction = PredictVp8lLayout(compared, width, comparedRows, layout, cancellationToken);
            double bits = EstimateVp8lLayoutBits(prediction.Residuals, width, settings.TrialAttempts,
                settings.MaximumCacheBits, settings.ParsePasses > 0, cancellationToken);
            if (layout == Vp8lLayout.Untransformed) untransformedBits = bits;
            else if (layout != Vp8lLayout.PerTile) singleBits = Math.Min(singleBits, bits);
            // Per-tile prediction also gets per-tile colour multipliers, which the estimate leaves out, so
            // another layout has to win clearly.
            if (layout != Vp8lLayout.PerTile && perTile) bits *= 1.03;
            if (bits < bestBits) {
                bestBits = bits;
                best = prediction;
            }
        }
        return compared == pixels ? best! : PredictVp8lLayout(pixels, width, height, best!.Layout, cancellationToken);
    }

    private static Vp8lLayout ChooseVp8lSampleLayout(uint[] pixels, int width, int height, Vp8lEncoderSettings settings,
        CancellationToken cancellationToken) {
        int sampleRows = Math.Max(1, Math.Min(height, settings.TrialPixels / width));
        uint[] sample;
        if (sampleRows >= height) {
            sample = pixels;
            sampleRows = height;
        } else {
            // Two contiguous bands, at a third and two thirds of the height, keep copies between rows as
            // they are in the whole image.
            int rowsPerBand = Math.Max(1, sampleRows / 2);
            sampleRows = rowsPerBand * 2;
            sample = new uint[sampleRows * width];
            for (int band = 0; band < 2; band++) {
                int start = Math.Max(0, Math.Min(height - rowsPerBand, (height * (band + 1) / 3) - rowsPerBand / 2));
                Array.Copy(pixels, start * width, sample, band * rowsPerBand * width, rowsPerBand * width);
            }
        }
        var residuals = new uint[sample.Length];
        bool subtractGreen = ShouldSubtractVp8lGreen(sample, width, sampleRows, cancellationToken);
        Vp8lLayout best = Vp8lSampleLayouts[0];
        double bestBits = double.MaxValue;
        foreach (Vp8lLayout layout in Vp8lSampleLayouts) {
            cancellationToken.ThrowIfCancellationRequested();
            ApplyVp8lUniformPredictor(sample, residuals, width, sampleRows, 9, (int)layout, subtractGreen, cancellationToken);
            double bits = EstimateVp8lLayoutBits(residuals, width, 8, settings.MaximumCacheBits, settings.ParsePasses > 0,
                cancellationToken);
            if (bits < bestBits) {
                bestBits = bits;
                best = layout;
            }
        }
        return best;
    }

    private static byte[]? TryEncodeCompressedVp8l(
        int width,
        int height,
        bool hasAlpha,
        byte[] rgba,
        int literalFileLength,
        long metadataBytes,
        long additionalRetainedManagedBytes,
        CancellationToken cancellationToken,
        Action<OfficeRasterEncodingCheckpoint>? checkpointObserver = null,
        OfficeImageEncodingEffort effort = OfficeImageEncodingEffort.Balanced) {
        try {
            int pixelCount = rgba.Length / 4;
            if (pixelCount > Vp8lCompressionMaximumPixels) return null;
            cancellationToken.ThrowIfCancellationRequested();
            Vp8lEncoderSettings settings = Vp8lEncoderSettings.For(effort);
            // A stream no smaller than the literal one is useless, so that size bounds the writer.
            int maximumBytes = checked(rgba.Length + 1024);
            // Peak, during the cost-based parse: source and residual pixels (8 bytes a pixel), per-pixel
            // matches (6), path costs and steps (10), the parse (8) and a finished candidate (4), with
            // margin; the stream writer later needs less. Tile histograms and their clusters, the hash
            // chain heads and the distance-code table come on top.
            long entropyTiles = settings.EntropyTiles <= 1 ? 1 :
                Math.Min(settings.EntropyTiles, (long)DivideRoundUp(width, 4) * DivideRoundUp(height, 4));
            long compressionBytes = checked(40L * pixelCount + entropyTiles * 17_500L +
                (4L << GetVp8lMatchHashBits(pixelCount)) + 32L * width + 256L * 1024L);
            if (!IsEncodingWorkingSetWithinLimit(rgba.LongLength, literalFileLength, compressionBytes,
                    metadataBytes, additionalRetainedManagedBytes)) return null;

            var pixels = new uint[pixelCount];
            FillVp8lSourcePixels(rgba, pixels, cancellationToken, checkpointObserver);
            Vp8lPrediction? prediction = ChooseVp8lPrediction(pixels, width, height, settings, cancellationToken);
            Vp8lLayout layout = prediction.Layout;
            byte[]? best = EncodeVp8lCandidate(width, height, hasAlpha, settings, prediction, maximumBytes, cancellationToken);
            prediction = null;
            if (effort == OfficeImageEncodingEffort.Balanced && pixelCount <= 1 << 16) {
                // Small images cost little to encode, and their map and header overhead is easy to misjudge,
                // so the other whole-image candidates are encoded too and the smallest stream kept.
                foreach (Vp8lLayout other in settings.Layouts) {
                    if (other == layout) continue;
                    Keep(EncodeVp8lCandidate(width, height, hasAlpha, settings,
                        PredictVp8lLayout(pixels, width, height, other, cancellationToken), best?.Length ?? maximumBytes,
                        cancellationToken));
                }
            }
            if (effort == OfficeImageEncodingEffort.Smallest) {
                // The wider search is estimated, not exact, so Smallest also keeps the unpredicted and the
                // Balanced streams when they are smaller.
                if (layout != Vp8lLayout.Untransformed) {
                    Keep(EncodeVp8lCandidate(width, height, hasAlpha, settings,
                        PredictVp8lLayout(pixels, width, height, Vp8lLayout.Untransformed, cancellationToken),
                        best?.Length ?? maximumBytes, cancellationToken));
                }
                Vp8lEncoderSettings balanced = Vp8lEncoderSettings.For(OfficeImageEncodingEffort.Balanced);
                Keep(EncodeVp8lCandidate(width, height, hasAlpha, balanced,
                    ChooseVp8lPrediction(pixels, width, height, balanced, cancellationToken),
                    best?.Length ?? maximumBytes, cancellationToken));
            }
            return best;

            void Keep(byte[]? candidate) {
                if (candidate != null && (best == null || candidate.Length < best.Length)) best = candidate;
            }
        } catch (OverflowException) {
            return null;
        }
    }

    private static byte[]? EncodeVp8lCandidate(int width, int height, bool hasAlpha, Vp8lEncoderSettings settings,
        Vp8lPrediction prediction, int maximumBytes, CancellationToken cancellationToken) {
        Vp8lLayout layout = prediction.Layout;
        bool transformed = layout != Vp8lLayout.Untransformed;
        int predictorBits = prediction.Bits;
        // The colour transform rewrites the residuals in place, so a prediction is encoded once.
        uint[] residuals = prediction.Residuals;
        uint[] predictors = prediction.Map, colors = Array.Empty<uint>();
        bool subtractGreen = prediction.SubtractGreen;
        // One predictor everywhere keeps repeats intact; per-tile colour multipliers would break them again.
        int colorBits = !transformed || settings.ColorBits == 0 ? 0 : layout == Vp8lLayout.PerTile ? settings.ColorBits : 9;
        if (colorBits > 0) {
            colors = SelectVp8lColorTransform(residuals, width, height, colorBits, settings.ColorSampleStep,
                cancellationToken);
        }

        Vp8lTokens tokens;
        int cacheBits;
        if (settings.ParsePasses > 0) {
            tokens = CreateVp8lCostedTokens(residuals, width, settings.MatchAttempts, settings.MaximumCacheBits,
                settings.ParsePasses, settings.TryBothParseStarts, cancellationToken, out cacheBits);
        } else {
            tokens = CreateVp8lTokens(residuals, width, settings.MatchAttempts, cancellationToken);
            cacheBits = SelectVp8lCacheBits(tokens, residuals, settings.MaximumCacheBits, cancellationToken, out _,
                settings.QuickCacheChoice);
        }

        Vp8lEntropyPlan plan = settings.EntropyTiles > 1
            ? CreateVp8lEntropyPlan(tokens, residuals, width, height, cacheBits, settings.EntropyTiles,
                settings.ExactClusters, cancellationToken)
            : CreateVp8lSingleGroupPlan(tokens, residuals, cacheBits, cancellationToken);

        byte[]? stream = WriteStream(plan, maximumBytes);
        // On small images the group headers are a large share of the stream and their estimate is rough, so
        // the single-group stream is written as well and the smaller one kept.
        if (plan.TileGroups != null && residuals.Length <= 1 << 16) {
            byte[]? single = WriteStream(CreateVp8lSingleGroupPlan(tokens, residuals, cacheBits, cancellationToken),
                stream?.Length ?? maximumBytes);
            if (single != null && (stream == null || single.Length < stream.Length)) stream = single;
        }
        return stream;

        byte[]? WriteStream(Vp8lEntropyPlan streamPlan, int limit) {
            try {
                var writer = new Vp8lBitWriter(Math.Min(limit, Math.Max(4096, tokens.Count * 2)), limit);
                writer.WriteBits(0x2F, 8);
                writer.WriteBits((uint)(width - 1), 14);
                writer.WriteBits((uint)(height - 1), 14);
                writer.WriteBits(hasAlpha ? 1U : 0U, 1);
                writer.WriteBits(0, 3);
                if (transformed) {
                    // Transforms are listed in encoding order; the decoder undoes them in reverse.
                    writer.WriteBits(1, 1);
                    writer.WriteBits(0, 2);
                    writer.WriteBits((uint)(predictorBits - 2), 3);
                    WriteVp8lSubImage(writer, predictors, DivideRoundUp(width, 1 << predictorBits), cancellationToken);
                }
                if (subtractGreen) {
                    writer.WriteBits(1, 1);
                    writer.WriteBits(2, 2);
                }
                if (colors.Length != 0) {
                    writer.WriteBits(1, 1);
                    writer.WriteBits(1, 2);
                    writer.WriteBits((uint)(colorBits - 2), 3);
                    WriteVp8lSubImage(writer, colors, DivideRoundUp(width, 1 << colorBits), cancellationToken);
                }
                writer.WriteBits(0, 1);
                WriteVp8lImage(writer, residuals, width, tokens, cacheBits, streamPlan, mainImage: true, cancellationToken);
                return writer.ToArray();
            } catch (OverflowException) {
                return null;
            }
        }
    }

    private static bool ShouldSubtractVp8lGreenFromPixels(uint[] pixels, CancellationToken cancellationToken) {
        var counts = new int[4 * 256];
        int step = Math.Max(1, pixels.Length / 65536);
        for (int index = 0; index < pixels.Length; index += step) {
            if ((index & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            uint color = pixels[index];
            int red = (int)(color >> 16) & 255, green = (int)(color >> 8) & 255, blue = (int)color & 255;
            counts[red]++;
            counts[256 + blue]++;
            counts[512 + ((red - green) & 255)]++;
            counts[768 + ((blue - green) & 255)]++;
        }
        return Vp8lShannonBits(counts, 512, 256) + Vp8lShannonBits(counts, 768, 256) + 16 <
            Vp8lShannonBits(counts, 0, 256) + Vp8lShannonBits(counts, 256, 256);
    }

    // Transform maps are small images in their own right: one group, a short parse and a cache if it pays.
    private static void WriteVp8lSubImage(ILsbBitWriter writer, uint[] pixels, int width, CancellationToken cancellationToken) {
        Vp8lTokens tokens = CreateVp8lTokens(pixels, width, 16, cancellationToken);
        int cacheBits = SelectVp8lCacheBits(tokens, pixels, 6, cancellationToken, out _);
        Vp8lEntropyPlan plan = CreateVp8lSingleGroupPlan(tokens, pixels, cacheBits, cancellationToken);
        WriteVp8lImage(writer, pixels, width, tokens, cacheBits, plan, mainImage: false, cancellationToken);
    }

    private static void WriteVp8lImage(ILsbBitWriter writer, uint[] pixels, int width, Vp8lTokens tokens, int cacheBits,
        Vp8lEntropyPlan plan, bool mainImage, CancellationToken cancellationToken) {
        writer.WriteBits(cacheBits == 0 ? 0U : 1U, 1);
        if (cacheBits != 0) writer.WriteBits((uint)cacheBits, 4);
        if (mainImage) {
            if (plan.TileGroups == null) {
                writer.WriteBits(0, 1);
            } else {
                writer.WriteBits(1, 1);
                writer.WriteBits((uint)(plan.Bits - 2), 3);
                var groupImage = new uint[plan.TileGroups.Length];
                for (int index = 0; index < groupImage.Length; index++) {
                    int group = plan.TileGroups[index];
                    groupImage[index] = 0xFF000000U | (uint)(group >> 8) << 16 | (uint)(group & 255) << 8;
                }
                WriteVp8lSubImage(writer, groupImage, plan.Columns, cancellationToken);
            }
        }

        var codes = new Vp8lCodebook[plan.Groups.Length * 5];
        for (int group = 0; group < plan.Groups.Length; group++) {
            cancellationToken.ThrowIfCancellationRequested();
            for (int alphabet = 0; alphabet < 5; alphabet++) {
                byte[] lengths = OfficeHuffmanCodeLengths.Create(plan.Groups[group].Alphabet(alphabet), 15);
                WriteVp8lHuffmanTree(writer, lengths);
                codes[group * 5 + alphabet] = new Vp8lCodebook(lengths);
            }
        }

        ReplayVp8lTokens(tokens, pixels, cacheBits, cancellationToken, (position, token, cacheIndex) => {
            int baseCode = plan.GroupAt(position, width) * 5;
            if (Vp8lTokens.IsCopy(token)) {
                GetVp8lPrefix(Vp8lTokens.Length(token), Vp8lLengthPrefixCount, out int prefix, out int extra, out int value);
                codes[baseCode].Write(writer, 256 + prefix);
                writer.WriteBits((uint)value, extra);
                GetVp8lPrefix(Vp8lTokens.DistanceCode(token), Vp8lDistancePrefixCount,
                    out int distancePrefix, out int distanceExtra, out int distanceValue);
                codes[baseCode + 4].Write(writer, distancePrefix);
                writer.WriteBits((uint)distanceValue, distanceExtra);
            } else if (cacheIndex >= 0) {
                codes[baseCode].Write(writer, 256 + Vp8lLengthPrefixCount + cacheIndex);
            } else {
                uint color = Vp8lTokens.Color(token);
                codes[baseCode].Write(writer, (int)(color >> 8) & 255);
                codes[baseCode + 1].Write(writer, (int)(color >> 16) & 255);
                codes[baseCode + 2].Write(writer, (int)color & 255);
                codes[baseCode + 3].Write(writer, (int)(color >> 24));
            }
        });
    }

    private static void WriteSingleSymbolTree(ILsbBitWriter writer, int symbol) {
        writer.WriteBits(1, 1);
        writer.WriteBits(0, 1);
        writer.WriteBits(symbol > 1 ? 1U : 0U, 1);
        writer.WriteBits((uint)symbol, symbol > 1 ? 8 : 1);
    }

    private sealed class Vp8lCodebook {
        private readonly uint[] _codes;
        private readonly byte[] _lengths;
        private readonly int _singleSymbol;

        internal Vp8lCodebook(byte[] lengths) {
            _lengths = lengths;
            _codes = new uint[lengths.Length];
            var counts = new int[16];
            for (int index = 0; index < lengths.Length; index++) counts[lengths[index]]++;
            int symbols = lengths.Length - counts[0];
            counts[0] = 0;
            _singleSymbol = -1;
            var next = new int[16];
            int code = 0;
            for (int length = 1; length <= 15; length++) {
                code = (code + counts[length - 1]) << 1;
                next[length] = code;
            }
            for (int symbol = 0; symbol < lengths.Length; symbol++) {
                int length = lengths[symbol];
                if (length == 0) continue;
                if (symbols == 1) _singleSymbol = symbol;
                _codes[symbol] = ReverseBits((uint)next[length]++, length);
            }
        }

        internal void Write(ILsbBitWriter writer, int symbol) {
            if (symbol == _singleSymbol) return;
            int length = _lengths[symbol];
            if (length == 0) throw new InvalidOperationException("Cannot emit an unused WebP Huffman symbol.");
            writer.WriteBits(_codes[symbol], length);
        }
    }

    private static uint ReverseBits(uint value, int length) {
        uint reversed = 0;
        for (int index = 0; index < length; index++) {
            reversed = (reversed << 1) | (value & 1U);
            value >>= 1;
        }
        return reversed;
    }

    /// <summary>Least-significant-bit-first writer into one growing array, capped at the useful size.</summary>
    private sealed class Vp8lBitWriter : ILsbBitWriter {
        private readonly int _maximumBytes;
        private byte[] _bytes;
        private int _count;
        private ulong _buffer;
        private int _bitCount;

        internal Vp8lBitWriter(int capacity, int maximumBytes) {
            _bytes = new byte[Math.Max(16, capacity)];
            _maximumBytes = maximumBytes;
        }

        public void WriteBits(uint value, int count) {
            ulong mask = count == 32 ? uint.MaxValue : (1UL << count) - 1UL;
            _buffer |= ((ulong)value & mask) << _bitCount;
            _bitCount += count;
            while (_bitCount >= 8) {
                AddByte((byte)_buffer);
                _buffer >>= 8;
                _bitCount -= 8;
            }
        }

        public void Flush() {
            if (_bitCount == 0) return;
            AddByte((byte)_buffer);
            _buffer = 0;
            _bitCount = 0;
        }

        internal byte[] ToArray() {
            Flush();
            var result = new byte[_count];
            Buffer.BlockCopy(_bytes, 0, result, 0, _count);
            return result;
        }

        private void AddByte(byte value) {
            if (_count == _bytes.Length) {
                if (_count >= _maximumBytes) throw new OverflowException("The WebP stream is no smaller than the literal one.");
                var grown = new byte[Math.Min(_maximumBytes, checked(_bytes.Length * 2))];
                Buffer.BlockCopy(_bytes, 0, grown, 0, _count);
                _bytes = grown;
            }
            _bytes[_count++] = value;
        }
    }
}
