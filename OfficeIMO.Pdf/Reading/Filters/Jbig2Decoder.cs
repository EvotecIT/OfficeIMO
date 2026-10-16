using OfficeIMO.Drawing;
using System.Threading;

namespace OfficeIMO.Pdf.Filters;

/// <summary>
/// Bounded JBIG2 decoder for embedded PDF streams (ITU-T T.88, ISO 32000-1 7.4.7): page information, generic regions
/// (MQ arithmetic with templates 0 to 3, typical prediction and adaptive pixels, or MMR), refinement regions, and
/// arithmetic-coded symbol dictionaries and text regions. Huffman-coded dictionaries and regions and halftone regions
/// are rejected with <see cref="NotSupportedException"/>.
/// </summary>
internal static partial class Jbig2Decoder {
    private const int MaximumSegments = 100_000;
    private const int MaximumSymbols = 1_000_000;
    private const long MaximumBitmapPixels = 256L * 1024 * 1024;

    /// <summary>Decodes page 1 to packed 1-bit rows where 0 is black, the PDF JBIG2Decode convention.</summary>
    internal static byte[] Decode(byte[] data, byte[]? globals, int width, int height, int maximumBytes, CancellationToken cancellationToken) {
        if (width <= 0 || height <= 0) throw new InvalidDataException("JBIG2 images require positive dimensions.");
        long stride = ((long)width + 7) / 8;
        if (stride * height > maximumBytes || (long)width * height > MaximumBitmapPixels) throw new InvalidDataException("JBIG2 image exceeds the decoded byte limit.");
        var context = new DecodingContext(width, height, maximumBytes, cancellationToken);
        if (globals != null && globals.Length > 0) context.ProcessSegments(globals);
        context.ProcessSegments(data);
        Bitmap page = context.Page ?? context.CreateBitmap(width, height);
        var output = new byte[stride * height];
        for (int y = 0; y < height; y++) {
            cancellationToken.ThrowIfCancellationRequested();
            for (int x = 0; x < width; x++) {
                    if ((x & 1023) == 0) cancellationToken.ThrowIfCancellationRequested();
                bool black = x < page.Width && y < page.Height && page.Get(x, y) != 0;
                if (!black) output[y * stride + x / 8] |= (byte)(0x80 >> (x & 7));
            }
        }
        return output;
    }

    /// <summary>
    /// Scans segment headers without decoding and names the first feature <see cref="Decode"/> rejects:
    /// "halftone", "huffman", "malformed" or "segment-N"; null when every segment is supported.
    /// </summary>
    internal static string? FindUnsupportedFeature(byte[] data, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try {
            int position = 0, count = 0;
            while (position < data.Length) {
                cancellationToken.ThrowIfCancellationRequested();
                if (++count > MaximumSegments) return "malformed";
                Segment segment = DecodingContext.ReadSegmentHeader(data, ref position, cancellationToken: cancellationToken);
                switch (segment.Type) {
                    case 51: return null;
                    case 16: case 20: case 22: case 23: return "halftone";
                    case 0:
                        if (HasHuffmanFlag(segment, 0)) return "huffman";
                        break;
                    case 4: case 6: case 7:
                        if (HasHuffmanFlag(segment, 17)) return "huffman";
                        break;
                    case 36: case 38: case 39: case 40: case 42: case 43: case 48: case 49: case 50: case 52: case 53: case 62: break;
                    default: return "segment-" + segment.Type.ToString(System.Globalization.CultureInfo.InvariantCulture);
                }
            }
            return null;
        } catch (InvalidDataException) {
            return "malformed";
        }
    }

    // SDHUFF and SBHUFF are bit 0 of the big-endian two-byte flags that follow the optional region information.
    private static bool HasHuffmanFlag(Segment segment, int flagsOffset) =>
        segment.DataEnd - segment.DataStart >= flagsOffset + 2 && (segment.Data[segment.DataStart + flagsOffset + 1] & 1) != 0;

    private sealed class Segment {
        internal int Number;
        internal int Type;
        internal int[] Referred = Array.Empty<int>();
        internal int DataStart;
        internal int DataEnd;
        internal byte[] Data = Array.Empty<byte>();
    }

    private sealed class DecodingContext {
        private readonly int _width, _height;
        private readonly CancellationToken _cancellationToken;
        private readonly Dictionary<int, List<Bitmap>> _symbols = new();
        private readonly Dictionary<int, Bitmap> _regions = new();
        private int _segmentCount;
        private int _pageDefaultCombination;
        private long _work;
        private long _allocated;
        private readonly int _maximumBytes;
        private readonly long _maximumWork;

        internal Bitmap CreateBitmap(int width, int height) {
            long pixels = (long)width * height;
            if (width < 0 || height < 0) throw new InvalidDataException("JBIG2 bitmap dimensions are invalid.");
            ReserveBytes(pixels);
            Charge(pixels);
            return new Bitmap(width, height, Charge, _cancellationToken);
        }

        private void ReserveBytes(long bytes) {
            _cancellationToken.ThrowIfCancellationRequested();
            if (bytes < 0 || bytes > _maximumBytes - _allocated)
                throw new InvalidDataException("JBIG2 storage exceeds the decoded byte budget.");
            _allocated += bytes;
        }

        private byte[] CreateContexts(int count) {
            ReserveBytes(count);
            return new byte[count];
        }

        private List<Bitmap> CreateSymbolList(int count) {
            if ((uint)count > MaximumSymbols) throw new InvalidDataException("JBIG2 symbol list is too large.");
            ReserveBytes(64L + (long)count * IntPtr.Size);
            return new List<Bitmap>(count);
        }

        internal DecodingContext(int width, int height, int maximumBytes, CancellationToken cancellationToken) {
            _maximumBytes = maximumBytes;
            _maximumWork = Math.Min(MaximumBitmapPixels * 4, Math.Max(1_000_000L, maximumBytes * 32L));
            _width = width;
            _height = height;
            _cancellationToken = cancellationToken;
        }

        internal Bitmap? Page { get; private set; }

        internal void Charge(long pixels) {
            _cancellationToken.ThrowIfCancellationRequested();
            if (pixels < 0 || pixels > _maximumWork - _work) throw new InvalidDataException("JBIG2 decoding exceeds its work budget.");
            _work += pixels;
        }

        internal void ProcessSegments(byte[] data) {
            int position = 0;
            while (position < data.Length) {
                _cancellationToken.ThrowIfCancellationRequested();
                if (++_segmentCount > MaximumSegments) throw new InvalidDataException("JBIG2 stream has too many segments.");
                Segment segment = ReadSegmentHeader(data, ref position, ReserveBytes, _cancellationToken);
                if (segment.Type == 51) break; // end of file
                ProcessSegment(segment);
            }
        }

        // Segment header syntax (7.2).
        internal static Segment ReadSegmentHeader(byte[] data, ref int position, Action<long>? reserveBytes = null,
            CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            var segment = new Segment { Data = data };
            segment.Number = (int)ReadUInt32(data, ref position);
            int flags = ReadByte(data, ref position);
            segment.Type = flags & 0x3F;
            bool largePageAssociation = (flags & 0x40) != 0;
            int referredByte = ReadByte(data, ref position);
            int referredCount = referredByte >> 5;
            if (referredCount == 7) {
                position--;
                referredCount = (int)(ReadUInt32(data, ref position) & 0x1FFFFFFF);
                if (referredCount > MaximumSegments) throw new InvalidDataException("JBIG2 segment refers to too many segments.");
                position += (referredCount + 8) / 8;
            }
            int referenceSize = segment.Number <= 256 ? 1 : segment.Number <= 65536 ? 2 : 4;
            reserveBytes?.Invoke((long)referredCount * sizeof(int));
            segment.Referred = new int[referredCount];
            for (int index = 0; index < referredCount; index++) {
                if ((index & 1023) == 0) cancellationToken.ThrowIfCancellationRequested();
                segment.Referred[index] = referenceSize == 1 ? ReadByte(data, ref position)
                    : referenceSize == 2 ? (ReadByte(data, ref position) << 8) | ReadByte(data, ref position)
                    : (int)ReadUInt32(data, ref position);
            }
            position += largePageAssociation ? 4 : 1;
            uint length = ReadUInt32(data, ref position);
            if (position > data.Length) throw new InvalidDataException("Truncated JBIG2 segment header.");
            segment.DataStart = position;
            if (length == uint.MaxValue && segment.Type is not (38 or 39))
                throw new InvalidDataException("Unknown JBIG2 segment lengths require an immediate generic region.");
            segment.DataEnd = length == 0xFFFFFFFF ? FindUnknownLengthEnd(data, position, cancellationToken) : (int)Math.Min(data.Length, position + (long)length);
            if (length != 0xFFFFFFFF && position + (long)length > data.Length) throw new InvalidDataException("Truncated JBIG2 segment data.");
            position = segment.DataEnd;
            return segment;
        }

        // An immediate generic region of unknown length ends with 0xFFAC or 0x0000 followed by a four-byte row count (7.2.7).
        private static int FindUnknownLengthEnd(byte[] data, int start, CancellationToken cancellationToken) {
            if (start > data.Length - 18) throw new InvalidDataException("Truncated JBIG2 generic region.");
            int flags = data[start + 17];
            bool mmr = (flags & 1) != 0;
            int template = (flags >> 1) & 3;
            int encodedStart = start + 18 + (mmr ? 0 : template == 0 ? 8 : 2);
            for (int index = encodedStart; index + 5 < data.Length; index++) {
                if ((index & 1023) == 0) cancellationToken.ThrowIfCancellationRequested();
                if (mmr ? data[index] == 0 && data[index + 1] == 0 : data[index] == 0xFF && data[index + 1] == 0xAC) return index + 6;
            }
            throw new InvalidDataException("JBIG2 unknown-length region has no complete terminator and row count.");
        }

        private void ProcessSegment(Segment segment) {
            switch (segment.Type) {
                case 48: ProcessPageInformation(segment); break;
                case 0: ProcessSymbolDictionary(segment); break;
                case 4: case 6: case 7: ProcessTextRegion(segment); break;
                case 36: case 38: case 39: ProcessGenericRegion(segment); break;
                case 40: case 42: case 43: ProcessRefinementRegion(segment); break;
                case 16: case 20: case 22: case 23: throw new NotSupportedException("JBIG2 pattern and halftone regions are not supported.");
                case 49: case 50: case 52: case 62: break; // end of page, end of stripe, profiles, extensions
                case 53: break; // custom Huffman tables are only used by unsupported Huffman coding
                default: throw new NotSupportedException("Unsupported JBIG2 segment type " + segment.Type + ".");
            }
        }

        private void ProcessPageInformation(Segment segment) {
            int position = segment.DataStart;
            int width = (int)ReadUInt32(segment.Data, ref position, segment.DataEnd);
            uint heightValue = ReadUInt32(segment.Data, ref position, segment.DataEnd);
            position += 8;
            int flags = ReadByte(segment.Data, ref position, segment.DataEnd);
            int height = heightValue == 0xFFFFFFFF ? _height : (int)Math.Min(heightValue, (uint)_height * 2);
            Page = CreateBitmap(Math.Min(Math.Max(width, 0), _width * 2), Math.Max(height, 0));
            Charge((long)Page.Width * Page.Height);
            Page.Fill((flags >> 2) & 1);
            _pageDefaultCombination = (flags >> 3) & 3;
        }

        private Bitmap EnsurePage() {
            if (Page == null) {
                Page = CreateBitmap(_width, _height);
                Charge((long)_width * _height);
            }
            return Page;
        }

        private static (int Width, int Height, int X, int Y, int Operator) ReadRegionInfo(Segment segment, ref int position) {
            int width = (int)ReadUInt32(segment.Data, ref position, segment.DataEnd);
            uint height = ReadUInt32(segment.Data, ref position, segment.DataEnd);
            int x = (int)ReadUInt32(segment.Data, ref position, segment.DataEnd);
            int y = (int)ReadUInt32(segment.Data, ref position, segment.DataEnd);
            int flags = ReadByte(segment.Data, ref position, segment.DataEnd);
            if (height == 0xFFFFFFFF) {
                int end = segment.DataEnd - 4;
                height = ReadUInt32(segment.Data, ref end, segment.DataEnd);
            }
            if (width < 0 || height > int.MaxValue || (long)width * height > MaximumBitmapPixels) throw new InvalidDataException("JBIG2 region is too large.");
            return (width, (int)height, x, y, flags & 7);
        }

        private void Place(Segment segment, Bitmap region, (int Width, int Height, int X, int Y, int Operator) info) {
            if (segment.Type is 4 or 36 or 40) {
                _regions[segment.Number] = region;
                return;
            }
            EnsurePage().Compose(region, info.X, info.Y, info.Operator);
        }

        private void ProcessGenericRegion(Segment segment) {
            int position = segment.DataStart;
            var info = ReadRegionInfo(segment, ref position);
            int flags = ReadByte(segment.Data, ref position, segment.DataEnd);
            bool mmr = (flags & 1) != 0;
            int template = (flags >> 1) & 3;
            bool typicalPrediction = (flags & 8) != 0;
            sbyte[] at = Array.Empty<sbyte>();
            if (!mmr) at = ReadAt(segment.Data, ref position, template == 0 ? 4 : 1, segment.DataEnd);
            Charge((long)info.Width * info.Height);
            Bitmap region;
            if (mmr) {
                int end = segment.DataEnd;
                ReserveBytes(end - (long)position);
                var encoded = new byte[Math.Max(0, end - position)];
                Buffer.BlockCopy(segment.Data, position, encoded, 0, encoded.Length);
                region = DecodeMmr(encoded, info.Width, info.Height);
            } else {
                var decoder = new ArithmeticDecoder(segment.Data, position, segment.DataEnd, Charge);
                region = DecodeGeneric(decoder, CreateContexts(1 << 16), info.Width, info.Height, template, typicalPrediction, at);
            }
            Place(segment, region, info);
        }

        private Bitmap DecodeMmr(byte[] encoded, int width, int height) {
            var region = CreateBitmap(width, height);
            if (width == 0 || height == 0) return region;
            ReserveBytes(((long)width + 7) / 8 * height);
            byte[] rows = OfficeFaxDecoder.Decode(encoded, width, height, -1, endOfLine: false, byteAligned: false, blackIsOne: true,
                endOfBlock: false, maximumBytes: checked((int)(((long)width + 7) / 8 * height)), _cancellationToken);
            int stride = (width + 7) / 8;
            for (int y = 0; y < height; y++) {
                for (int x = 0; x < width; x++) region.Set(x, y, (rows[y * stride + x / 8] >> (7 - (x & 7))) & 1);
            }
            return region;
        }

        private static sbyte[] ReadAt(byte[] data, ref int position, int pairs, int end) {
            var at = new sbyte[pairs * 2];
            for (int index = 0; index < at.Length; index++) at[index] = (sbyte)ReadByte(data, ref position, end);
            return at;
        }

        private static readonly (int X, int Y)[][] GenericTemplates = {
            new[] { (-1, -2), (0, -2), (1, -2), (-2, -1), (-1, -1), (0, -1), (1, -1), (2, -1), (-4, 0), (-3, 0), (-2, 0), (-1, 0) },
            new[] { (-1, -2), (0, -2), (1, -2), (2, -2), (-2, -1), (-1, -1), (0, -1), (1, -1), (2, -1), (-3, 0), (-2, 0), (-1, 0) },
            new[] { (-1, -2), (0, -2), (1, -2), (-2, -1), (-1, -1), (0, -1), (1, -1), (-2, 0), (-1, 0) },
            new[] { (-3, -1), (-2, -1), (-1, -1), (0, -1), (1, -1), (-4, 0), (-3, 0), (-2, 0), (-1, 0) }
        };
        private static readonly int[] TypicalPredictionContexts = { 0x9B25, 0x0795, 0x00E5, 0x0195 };
        private static readonly (int X, int Y)[][] DefaultGenericAt = {
            new[] { (3, -1), (-3, -1), (2, -2), (-2, -2) },
            new[] { (3, -1) },
            new[] { (2, -1) },
            new[] { (2, -1) }
        };

        // Generic region decoding procedure (6.2.5). The context lists template and adaptive pixels sorted by row,
        // then column, most significant first.
        private Bitmap DecodeGeneric(ArithmeticDecoder decoder, byte[] contexts, int width, int height, int template, bool typicalPrediction, sbyte[] at) {
            var region = CreateBitmap(width, height);
            // Adaptive pixels keep the context bit positions of their nominal locations (6.2.5.3).
            var pixels = new List<(int X, int Y, int At)>();
            foreach ((int X, int Y) pixel in GenericTemplates[template]) pixels.Add((pixel.X, pixel.Y, -1));
            (int X, int Y)[] nominal = DefaultGenericAt[template];
            for (int index = 0; index < nominal.Length; index++) pixels.Add((nominal[index].X, nominal[index].Y, index));
            pixels.Sort(static (left, right) => left.Y != right.Y ? left.Y.CompareTo(right.Y) : left.X.CompareTo(right.X));
            Charge((long)width * height * pixels.Count);
            var offsets = new (int X, int Y)[pixels.Count];
            for (int index = 0; index < offsets.Length; index++) {
                int slot = pixels[index].At;
                offsets[index] = slot >= 0 && 2 * slot + 1 < at.Length ? (at[2 * slot], at[2 * slot + 1]) : (pixels[index].X, pixels[index].Y);
            }
            bool lineTypical = false;
            for (int y = 0; y < height; y++) {
                _cancellationToken.ThrowIfCancellationRequested();
                if (typicalPrediction) {
                    lineTypical ^= decoder.Decode(contexts, TypicalPredictionContexts[template]) != 0;
                    if (lineTypical) {
                        if (y > 0) region.CopyRow(y - 1, y);
                        continue;
                    }
                }
                for (int x = 0; x < width; x++) {
                    if ((x & 1023) == 0) _cancellationToken.ThrowIfCancellationRequested();
                    int context = 0;
                    for (int index = 0; index < offsets.Length; index++) context = (context << 1) | region.Get(x + offsets[index].X, y + offsets[index].Y);
                    if (decoder.Decode(contexts, context) != 0) region.Set(x, y, 1);
                }
            }
            return region;
        }

        private static readonly (int X, int Y)[][] RefinementCoding = {
            new[] { (0, -1), (1, -1), (-1, 0) },
            new[] { (-1, -1), (0, -1), (1, -1), (-1, 0) }
        };
        private static readonly (int X, int Y)[][] RefinementReference = {
            new[] { (0, -1), (1, -1), (-1, 0), (0, 0), (1, 0), (-1, 1), (0, 1), (1, 1) },
            new[] { (0, -1), (-1, 0), (0, 0), (1, 0), (0, 1), (1, 1) }
        };

        // Generic refinement region decoding procedure (6.3.5.6).
        private Bitmap DecodeRefinement(ArithmeticDecoder decoder, byte[] contexts, int width, int height, int template, Bitmap reference,
            int referenceDx, int referenceDy, bool typicalPrediction, sbyte[] at) {
            Charge((long)width * height);
            var region = CreateBitmap(width, height);
            var coding = new List<(int X, int Y)>(RefinementCoding[template]);
            var referencePixels = new List<(int X, int Y)>(RefinementReference[template]);
            if (template == 0 && at.Length >= 4) {
                coding.Add((at[0], at[1]));
                referencePixels.Add((at[2], at[3]));
            }
            Charge((long)width * height * (coding.Count + referencePixels.Count));
            // The typical-prediction context is the one whose only set pixel is the reference centre (6.3.5.6).
            int typicalContext = 1 << (referencePixels.Count - 1 - RefinementReference[template].ToList().IndexOf((0, 0)));
            bool lineTypical = false;
            for (int y = 0; y < height; y++) {
                _cancellationToken.ThrowIfCancellationRequested();
                if (typicalPrediction) lineTypical ^= decoder.Decode(contexts, typicalContext) != 0;
                for (int x = 0; x < width; x++) {
                    if ((x & 1023) == 0) _cancellationToken.ThrowIfCancellationRequested();
                    int rx = x - referenceDx, ry = y - referenceDy;
                    if (lineTypical) {
                        int value = reference.Get(rx, ry);
                        bool uniform = true;
                        for (int dy = -1; dy <= 1 && uniform; dy++) {
                            for (int dx = -1; dx <= 1 && uniform; dx++) uniform = reference.Get(rx + dx, ry + dy) == value;
                        }
                        if (uniform) {
                            if (value != 0) region.Set(x, y, 1);
                            continue;
                        }
                    }
                    int context = 0;
                    foreach ((int X, int Y) offset in coding) context = (context << 1) | region.Get(x + offset.X, y + offset.Y);
                    foreach ((int X, int Y) offset in referencePixels) context = (context << 1) | reference.Get(rx + offset.X, ry + offset.Y);
                    if (decoder.Decode(contexts, context) != 0) region.Set(x, y, 1);
                }
            }
            return region;
        }

        private void ProcessRefinementRegion(Segment segment) {
            int position = segment.DataStart;
            var info = ReadRegionInfo(segment, ref position);
            int flags = ReadByte(segment.Data, ref position, segment.DataEnd);
            int template = flags & 1;
            bool typicalPrediction = (flags & 2) != 0;
            sbyte[] at = template == 0 ? ReadAt(segment.Data, ref position, 2, segment.DataEnd) : Array.Empty<sbyte>();
            Bitmap reference;
            if (segment.Referred.Length > 0 && _regions.TryGetValue(segment.Referred[0], out Bitmap? referred)) {
                reference = referred;
            } else {
                // Without a referred region the reference is the page area the region covers (7.4.7.5).
                Bitmap page = EnsurePage();
                reference = CreateBitmap(info.Width, info.Height);
                for (int y = 0; y < info.Height; y++) {
                    _cancellationToken.ThrowIfCancellationRequested();
                    for (int x = 0; x < info.Width; x++) {
                        if ((x & 1023) == 0) _cancellationToken.ThrowIfCancellationRequested();
                        reference.Set(x, y, page.Get(info.X + x, info.Y + y));
                    }
                }
            }
            var decoder = new ArithmeticDecoder(segment.Data, position, segment.DataEnd, Charge);
            Bitmap region = DecodeRefinement(decoder, CreateContexts(1 << 14), info.Width, info.Height, template, reference, 0, 0, typicalPrediction, at);
            Place(segment, region, info);
        }

        private List<Bitmap> ReferredSymbols(Segment segment) {
            int count = 0;
            foreach (int referred in segment.Referred) {
                Charge(1L);
                if (!_symbols.TryGetValue(referred, out List<Bitmap>? exported)) continue;
                if (exported.Count > MaximumSymbols - count) throw new InvalidDataException("JBIG2 symbol references exceed the symbol limit.");
                count += exported.Count;
            }
            var symbols = CreateSymbolList(count);
            foreach (int referred in segment.Referred)
                if (_symbols.TryGetValue(referred, out List<Bitmap>? exported)) { Charge(exported.Count); symbols.AddRange(exported); }
            return symbols;
        }

        private static int CodeLength(int count) {
            int length = 0;
            while ((1L << length) < count) length++;
            return length;
        }

        // Symbol dictionary segment (7.4.2) and decoding procedure (6.5.5), arithmetic coding only.
        private void ProcessSymbolDictionary(Segment segment) {
            int position = segment.DataStart;
            int flags = (ReadByte(segment.Data, ref position, segment.DataEnd) << 8) | ReadByte(segment.Data, ref position, segment.DataEnd);
            bool huffman = (flags & 1) != 0;
            bool refinementAggregate = (flags & 2) != 0;
            int template = (flags >> 10) & 3;
            int refinementTemplate = (flags >> 12) & 1;
            if (huffman) throw new NotSupportedException("Huffman-coded JBIG2 symbol dictionaries are not supported.");
            sbyte[] at = ReadAt(segment.Data, ref position, template == 0 ? 4 : 1, segment.DataEnd);
            sbyte[] refinementAt = refinementAggregate && refinementTemplate == 0 ? ReadAt(segment.Data, ref position, 2, segment.DataEnd) : Array.Empty<sbyte>();
            int exportedCount = (int)ReadUInt32(segment.Data, ref position, segment.DataEnd);
            int newCount = (int)ReadUInt32(segment.Data, ref position, segment.DataEnd);
            List<Bitmap> input = ReferredSymbols(segment);
            if (newCount < 0 || exportedCount < 0 || (long)newCount + input.Count > MaximumSymbols) throw new InvalidDataException("JBIG2 symbol dictionary is too large.");

            var decoder = new ArithmeticDecoder(segment.Data, position, segment.DataEnd, Charge);
            var integers = new IntegerContexts(CreateContexts);
            var genericContexts = CreateContexts(1 << 16);
            var refinementContexts = CreateContexts(1 << 14);
            int codeLength = CodeLength(input.Count + newCount);
            var idContexts = CreateContexts(1 << (codeLength + 1));
            var symbols = CreateSymbolList(input.Count + newCount);
            symbols.AddRange(input);
            int heightClass = 0;
            int decoded = 0;
            while (decoded < newCount) {
                _cancellationToken.ThrowIfCancellationRequested();
                int deltaHeight = decoder.DecodeInteger(integers["IADH"]) ?? throw new InvalidDataException("JBIG2 height class delta is out of band.");
                heightClass += deltaHeight;
                if (heightClass < 0) throw new InvalidDataException("JBIG2 symbol height is negative.");
                int symbolWidth = 0;
                while (true) {
                    int? deltaWidth = decoder.DecodeInteger(integers["IADW"]);
                    if (deltaWidth == null) break;
                    symbolWidth += deltaWidth.Value;
                    if (symbolWidth < 0 || decoded >= newCount) throw new InvalidDataException("JBIG2 symbol width is invalid.");
                    Bitmap symbol;
                    if (!refinementAggregate) {
                        Charge((long)symbolWidth * heightClass);
                        symbol = DecodeGeneric(decoder, genericContexts, symbolWidth, heightClass, template, false, at);
                    } else {
                        int instances = decoder.DecodeInteger(integers["IAAI"]) ?? throw new InvalidDataException("JBIG2 aggregate count is out of band.");
                        if (instances == 1) {
                            int id = decoder.DecodeId(idContexts, codeLength);
                            int dx = decoder.DecodeInteger(integers["IARDX"]) ?? 0;
                            int dy = decoder.DecodeInteger(integers["IARDY"]) ?? 0;
                            if ((uint)id >= (uint)symbols.Count) throw new InvalidDataException("JBIG2 refinement refers to an unknown symbol.");
                            symbol = DecodeRefinement(decoder, refinementContexts, symbolWidth, heightClass, refinementTemplate, symbols[id], dx, dy, false, refinementAt);
                        } else if (instances > 1) {
                            symbol = DecodeTextRegion(decoder, integers, idContexts, refinementContexts, symbolWidth, heightClass, instances,
                                symbols, codeLength, stripLog: 0, referenceCorner: 1, transposed: false, combination: 0, defaultPixel: 0,
                                dsOffset: 0, refine: true, refinementTemplate, refinementAt);
                        } else {
                            throw new InvalidDataException("JBIG2 aggregate count is invalid.");
                        }
                    }
                    symbols.Add(symbol);
                    decoded++;
                }
            }
            // Export flags (6.5.10) select from input then new symbols.
            if (exportedCount > symbols.Count) throw new InvalidDataException("JBIG2 exported symbol count is invalid.");
            var exported = CreateSymbolList(exportedCount);
            bool exporting = false;
            int index = 0;
            while (index < symbols.Count) {
                int run = decoder.DecodeInteger(integers["IAEX"]) ?? throw new InvalidDataException("JBIG2 export run is out of band.");
                if (run < 0 || run > symbols.Count - index) throw new InvalidDataException("JBIG2 export run is invalid.");
                if (exporting) {
                    if (run > exportedCount - exported.Count) throw new InvalidDataException("JBIG2 export runs exceed the declared count.");
                    for (int exportedIndex = index; exportedIndex < index + run; exportedIndex++) {
                        Charge(1L);
                        exported.Add(symbols[exportedIndex]);
                    }
                }
                index += run;
                exporting = !exporting;
            }
            if (exported.Count != exportedCount) throw new InvalidDataException("JBIG2 export runs do not match the declared count.");
            _symbols[segment.Number] = exported;
        }

        // Text region segment (7.4.3), arithmetic coding only.
        private void ProcessTextRegion(Segment segment) {
            int position = segment.DataStart;
            var info = ReadRegionInfo(segment, ref position);
            int flags = (ReadByte(segment.Data, ref position, segment.DataEnd) << 8) | ReadByte(segment.Data, ref position, segment.DataEnd);
            bool huffman = (flags & 1) != 0;
            bool refine = (flags & 2) != 0;
            int stripLog = (flags >> 2) & 3;
            int referenceCorner = (flags >> 4) & 3;
            bool transposed = (flags & 0x40) != 0;
            int combination = (flags >> 7) & 3;
            int defaultPixel = (flags >> 9) & 1;
            int dsOffset = (flags >> 10) & 0x1F;
            if ((dsOffset & 0x10) != 0) dsOffset -= 32;
            int refinementTemplate = (flags >> 15) & 1;
            if (huffman) throw new NotSupportedException("Huffman-coded JBIG2 text regions are not supported.");
            sbyte[] refinementAt = refine && refinementTemplate == 0 ? ReadAt(segment.Data, ref position, 2, segment.DataEnd) : Array.Empty<sbyte>();
            int instances = (int)ReadUInt32(segment.Data, ref position, segment.DataEnd);
            if (instances < 0 || instances > MaximumSymbols * 4) throw new InvalidDataException("JBIG2 text region has too many instances.");
            List<Bitmap> symbols = ReferredSymbols(segment);
            int codeLength = CodeLength(symbols.Count);
            var decoder = new ArithmeticDecoder(segment.Data, position, segment.DataEnd, Charge);
            Charge((long)info.Width * info.Height);
            Bitmap region = DecodeTextRegion(decoder, new IntegerContexts(CreateContexts), CreateContexts(1 << (codeLength + 1)), CreateContexts(1 << 14),
                info.Width, info.Height, instances, symbols, codeLength, stripLog, referenceCorner, transposed, combination, defaultPixel,
                dsOffset, refine, refinementTemplate, refinementAt);
            Place(segment, region, info);
        }

        // Text region decoding procedure (6.4.5). Reference corners: 0 bottom-left, 1 top-left, 2 bottom-right, 3 top-right.
        private Bitmap DecodeTextRegion(ArithmeticDecoder decoder, IntegerContexts integers, byte[] idContexts, byte[] refinementContexts,
            int width, int height, int instances, List<Bitmap> symbols, int codeLength, int stripLog, int referenceCorner, bool transposed,
            int combination, int defaultPixel, int dsOffset, bool refine, int refinementTemplate, sbyte[] refinementAt) {
            var region = CreateBitmap(width, height);
            if (defaultPixel != 0) region.Fill(1);
            int strips = 1 << stripLog;
            int stripT = -(decoder.DecodeInteger(integers["IADT"]) ?? 0) * strips;
            int firstS = 0;
            int count = 0;
            while (count < instances) {
                _cancellationToken.ThrowIfCancellationRequested();
                stripT += (decoder.DecodeInteger(integers["IADT"]) ?? throw new InvalidDataException("JBIG2 strip delta is out of band.")) * strips;
                bool first = true;
                int currentS = 0;
                while (count < instances) {
                    if (first) {
                        firstS += decoder.DecodeInteger(integers["IAFS"]) ?? throw new InvalidDataException("JBIG2 first S is out of band.");
                        currentS = firstS;
                        first = false;
                    } else {
                        int? deltaS = decoder.DecodeInteger(integers["IADS"]);
                        if (deltaS == null) break;
                        currentS += deltaS.Value + dsOffset;
                    }
                    int currentT = strips == 1 ? 0 : decoder.DecodeInteger(integers["IAIT"]) ?? 0;
                    int t = stripT + currentT;
                    int id = decoder.DecodeId(idContexts, codeLength);
                    if ((uint)id >= (uint)symbols.Count) throw new InvalidDataException("JBIG2 text region refers to an unknown symbol.");
                    Bitmap symbol = symbols[id];
                    int refinement = refine ? decoder.DecodeInteger(integers["IARI"]) ?? 0 : 0;
                    if (refinement != 0) {
                        int deltaWidth = decoder.DecodeInteger(integers["IARDW"]) ?? 0;
                        int deltaHeight = decoder.DecodeInteger(integers["IARDH"]) ?? 0;
                        int dx = decoder.DecodeInteger(integers["IARDX"]) ?? 0;
                        int dy = decoder.DecodeInteger(integers["IARDY"]) ?? 0;
                        int refinedWidth = symbol.Width + deltaWidth, refinedHeight = symbol.Height + deltaHeight;
                        if (refinedWidth < 0 || refinedHeight < 0) throw new InvalidDataException("JBIG2 refined symbol size is negative.");
                        symbol = DecodeRefinement(decoder, refinementContexts, refinedWidth, refinedHeight, refinementTemplate, symbol,
                            (deltaWidth >> 1) + dx, (deltaHeight >> 1) + dy, false, refinementAt);
                    }
                    int symbolWidth = symbol.Width, symbolHeight = symbol.Height;
                    if (!transposed && referenceCorner >= 2) currentS += symbolWidth - 1;
                    else if (transposed && (referenceCorner & 1) == 0) currentS += symbolHeight - 1;
                    int s = currentS;
                    int x, y;
                    if (!transposed) {
                        x = referenceCorner >= 2 ? s - symbolWidth + 1 : s;
                        y = (referenceCorner & 1) == 1 ? t : t - symbolHeight + 1;
                    } else {
                        x = referenceCorner >= 2 ? t - symbolWidth + 1 : t;
                        y = (referenceCorner & 1) == 1 ? s : s - symbolHeight + 1;
                    }
                    region.Compose(symbol, x, y, combination);
                    if (!transposed && referenceCorner < 2) currentS += symbolWidth - 1;
                    else if (transposed && (referenceCorner & 1) == 1) currentS += symbolHeight - 1;
                    count++;
                }
            }
            return region;
        }

        private static int ReadByte(byte[] data, ref int position, int end = -1) {
            if (position < 0 || position >= (end < 0 ? data.Length : end)) throw new InvalidDataException("Truncated JBIG2 data.");
            return data[position++];
        }

        private static uint ReadUInt32(byte[] data, ref int position, int end = -1) {
            if (position < 0 || position > (end < 0 ? data.Length : end) - 4) throw new InvalidDataException("Truncated JBIG2 data.");
            uint value = ((uint)data[position] << 24) | ((uint)data[position + 1] << 16) | ((uint)data[position + 2] << 8) | data[position + 3];
            position += 4;
            return value;
        }
    }
}
