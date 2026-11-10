namespace OfficeIMO.Drawing;

internal static partial class OfficeJpeg2000Decoder {
    private const int MarkerSoc = 0xFF4F, MarkerCap = 0xFF50, MarkerSiz = 0xFF51, MarkerCod = 0xFF52, MarkerCoc = 0xFF53,
        MarkerQcd = 0xFF5C, MarkerQcc = 0xFF5D, MarkerRgn = 0xFF5E, MarkerPoc = 0xFF5F, MarkerPpm = 0xFF60, MarkerPpt = 0xFF61,
        MarkerSot = 0xFF90, MarkerSop = 0xFF91, MarkerEph = 0xFF92, MarkerSod = 0xFF93, MarkerEoc = 0xFFD9;
    private const int MaximumComponents = 4096;
    private const int MaximumPrecision = 16;
    private const int MaximumLevels = 32;
    private const int MaximumLayers = 65535;

    private sealed class ComponentCoding {
        internal int Levels;
        internal int CodeBlockWidth; // exponent
        internal int CodeBlockHeight; // exponent
        internal int Style;
        internal bool Reversible;
        internal int[] PrecinctWidths = Array.Empty<int>(); // exponents per resolution
        internal int[] PrecinctHeights = Array.Empty<int>();
    }

    private sealed class GlobalCoding {
        internal bool Sop;
        internal bool Eph;
        internal int Progression;
        internal int Layers;
        internal bool MultipleComponentTransform;
        internal ComponentCoding Component = new();
    }

    private sealed class Quantization {
        internal int Style;
        internal int GuardBits;
        internal int[] Exponents = Array.Empty<int>();
        internal int[] Mantissas = Array.Empty<int>();
    }

    private sealed class ProgressionChange {
        internal int ResolutionStart, ComponentStart, LayerEnd, ResolutionEnd, ComponentEnd, Order;
    }

    /// <summary>Marker values of one header scope: the main header or a tile's first tile-part header.</summary>
    private sealed class HeaderScope {
        internal GlobalCoding? Coding;
        internal readonly Dictionary<int, ComponentCoding> ComponentCodings = new();
        internal Quantization? Quantization;
        internal readonly Dictionary<int, Quantization> ComponentQuantizations = new();
        internal readonly Dictionary<int, int> RegionShifts = new();
        internal readonly List<ProgressionChange> ProgressionChanges = new();
    }

    private sealed class TileData {
        internal readonly List<(int Start, int End)> Bodies = new();
        internal readonly List<(int Index, byte[] Data)> PacketHeaderChunks = new(); // PPT, ordered by Zppt
        internal readonly List<(int Start, int End)> PackedHeaders = new(); // PPM records, in tile-part order
        internal HeaderScope? Scope;
        internal int TileParts;
    }

    private sealed class Codestream {
        internal int Xsiz, Ysiz, XOsiz, YOsiz, XTsiz, YTsiz, XTOsiz, YTOsiz;
        internal int Components;
        internal int[] Precisions = Array.Empty<int>();
        internal bool[] Signed = Array.Empty<bool>();
        internal int[] SubsamplingX = Array.Empty<int>();
        internal int[] SubsamplingY = Array.Empty<int>();
        internal int TilesWide, TilesHigh;
        internal HeaderScope Main = new();
        internal Dictionary<int, TileData> Tiles = new();
        internal byte[]? PackedHeaderData;
        internal int Width => Xsiz - XOsiz;
        internal int Height => Ysiz - YOsiz;
    }

    private static Codestream ReadCodestream(byte[] bytes, int start, int end, Budget budget) {
        var reader = new MarkerReader(bytes, start, end);
        if (reader.ReadUInt16() != MarkerSoc) throw new Jpeg2000Exception("The JPEG 2000 codestream does not start with SOC.");
        var codestream = new Codestream();
        bool sawSize = false;
        var packedHeaders = new List<(int Index, int Start, int End)>();
        // Main header (A.4).
        while (true) {
            budget.Token.ThrowIfCancellationRequested();
            if (reader.Remaining < 2) throw new Jpeg2000Exception("The JPEG 2000 main header is truncated.");
            int marker = reader.ReadUInt16();
            if (marker == MarkerSot) {
                reader.Position -= 2;
                break;
            }
            if (marker < 0xFF30) throw new Jpeg2000Exception("The JPEG 2000 main header contains invalid data.");
            if (marker <= 0xFF3F) continue; // reserved markers without parameters (A.1.3)
            int segmentStart = reader.Position;
            int length = reader.ReadUInt16();
            if (length < 2 || length - 2 > reader.Remaining) throw new Jpeg2000Exception("A JPEG 2000 marker segment is truncated.");
            int segmentEnd = segmentStart + length;
            var segment = new MarkerReader(bytes, segmentStart + 2, segmentEnd);
            if (marker == MarkerSiz) {
                if (sawSize) throw new Jpeg2000Exception("The JPEG 2000 codestream has more than one SIZ marker.");
                ReadSize(segment, codestream, budget);
                sawSize = true;
            } else {
                if (!sawSize) throw new Jpeg2000Exception("The JPEG 2000 SIZ marker must follow SOC.");
                if (marker == MarkerCap) throw new Jpeg2000Exception("High-throughput (Part 15) JPEG 2000 codestreams are not supported.");
                if (marker == MarkerPpm) {
                    int index = segment.ReadByte();
                    packedHeaders.Add((index, segment.Position, segmentEnd));
                } else {
                    ReadScopeMarker(marker, segment, codestream, codestream.Main);
                }
            }
            reader.Position = segmentEnd;
        }
        if (codestream.Main.Coding == null || codestream.Main.Quantization == null)
            throw new Jpeg2000Exception("The JPEG 2000 main header lacks COD or QCD.");
        if (packedHeaders.Count > 0) {
            packedHeaders.Sort(static (left, right) => left.Index.CompareTo(right.Index));
            long total = 0;
            foreach (var chunk in packedHeaders) total += chunk.End - chunk.Start;
            budget.Reserve(total);
            var data = new byte[total];
            int offset = 0;
            foreach (var chunk in packedHeaders) {
                Buffer.BlockCopy(bytes, chunk.Start, data, offset, chunk.End - chunk.Start);
                offset += chunk.End - chunk.Start;
            }
            codestream.PackedHeaderData = data;
        }
        ReadTileParts(bytes, reader, end, codestream, budget);
        return codestream;
    }

    // Image and tile size (A.5.1).
    private static void ReadSize(MarkerReader segment, Codestream codestream, Budget budget) {
        int capabilities = segment.ReadUInt16();
        if ((capabilities & 0x4000) != 0) throw new Jpeg2000Exception("High-throughput (Part 15) JPEG 2000 codestreams are not supported.");
        codestream.Xsiz = ReadPositive(segment);
        codestream.Ysiz = ReadPositive(segment);
        codestream.XOsiz = ReadNonNegative(segment);
        codestream.YOsiz = ReadNonNegative(segment);
        codestream.XTsiz = ReadPositive(segment);
        codestream.YTsiz = ReadPositive(segment);
        codestream.XTOsiz = ReadNonNegative(segment);
        codestream.YTOsiz = ReadNonNegative(segment);
        int components = segment.ReadUInt16();
        if (codestream.XOsiz >= codestream.Xsiz || codestream.YOsiz >= codestream.Ysiz ||
            codestream.XTOsiz > codestream.XOsiz || codestream.YTOsiz > codestream.YOsiz ||
            (long)codestream.XTOsiz + codestream.XTsiz <= codestream.XOsiz ||
            (long)codestream.YTOsiz + codestream.YTsiz <= codestream.YOsiz)
            throw new Jpeg2000Exception("The JPEG 2000 image and tile geometry is invalid.");
        if (components < 1 || components > MaximumComponents) throw new Jpeg2000Exception("The JPEG 2000 component count is unsupported.");
        long tilesWide = ((long)codestream.Xsiz - codestream.XTOsiz + codestream.XTsiz - 1) / codestream.XTsiz;
        long tilesHigh = ((long)codestream.Ysiz - codestream.YTOsiz + codestream.YTsiz - 1) / codestream.YTsiz;
        if (tilesWide * tilesHigh > 65535) throw new Jpeg2000Exception("The JPEG 2000 image has too many tiles.");
        codestream.TilesWide = (int)tilesWide;
        codestream.TilesHigh = (int)tilesHigh;
        budget.Reserve(components * 16L);
        codestream.Components = components;
        codestream.Precisions = new int[components];
        codestream.Signed = new bool[components];
        codestream.SubsamplingX = new int[components];
        codestream.SubsamplingY = new int[components];
        for (int component = 0; component < components; component++) {
            int size = segment.ReadByte();
            codestream.Precisions[component] = (size & 0x7F) + 1;
            codestream.Signed[component] = (size & 0x80) != 0;
            codestream.SubsamplingX[component] = segment.ReadByte();
            codestream.SubsamplingY[component] = segment.ReadByte();
            if (codestream.Precisions[component] > MaximumPrecision) throw new Jpeg2000Exception("JPEG 2000 component precision above 16 bits is not supported.");
            if (codestream.SubsamplingX[component] < 1 || codestream.SubsamplingY[component] < 1) throw new Jpeg2000Exception("JPEG 2000 component subsampling is invalid.");
        }
    }

    private static int ReadPositive(MarkerReader segment) {
        uint value = segment.ReadUInt32();
        if (value == 0 || value > int.MaxValue) throw new Jpeg2000Exception("A JPEG 2000 size is out of range.");
        return (int)value;
    }

    private static int ReadNonNegative(MarkerReader segment) {
        uint value = segment.ReadUInt32();
        if (value > int.MaxValue) throw new Jpeg2000Exception("A JPEG 2000 offset is out of range.");
        return (int)value;
    }

    private static void ReadScopeMarker(int marker, MarkerReader segment, Codestream codestream, HeaderScope scope) {
        switch (marker) {
            case MarkerCod: {
                int style = segment.ReadByte();
                var coding = new GlobalCoding {
                    Sop = (style & 2) != 0,
                    Eph = (style & 4) != 0,
                    Progression = segment.ReadByte(),
                    Layers = segment.ReadUInt16()
                };
                int transform = segment.ReadByte();
                if (coding.Progression > 4 || coding.Layers < 1) throw new Jpeg2000Exception("The JPEG 2000 COD marker is invalid.");
                if (transform > 1) throw new Jpeg2000Exception("JPEG 2000 Part 2 multiple-component transforms are not supported.");
                coding.MultipleComponentTransform = transform == 1;
                coding.Component = ReadComponentCoding(segment, (style & 1) != 0);
                scope.Coding = coding;
                break;
            }
            case MarkerCoc: {
                int component = ReadComponentIndex(segment, codestream);
                int style = segment.ReadByte();
                scope.ComponentCodings[component] = ReadComponentCoding(segment, (style & 1) != 0);
                break;
            }
            case MarkerQcd:
                scope.Quantization = ReadQuantization(segment);
                break;
            case MarkerQcc: {
                int component = ReadComponentIndex(segment, codestream);
                scope.ComponentQuantizations[component] = ReadQuantization(segment);
                break;
            }
            case MarkerRgn: {
                int component = ReadComponentIndex(segment, codestream);
                int style = segment.ReadByte();
                int shift = segment.ReadByte();
                if (style != 0) throw new Jpeg2000Exception("The JPEG 2000 ROI style is unsupported.");
                scope.RegionShifts[component] = shift;
                break;
            }
            case MarkerPoc: {
                bool wide = codestream.Components >= 257;
                int entrySize = wide ? 9 : 7;
                while (segment.Remaining >= entrySize) {
                    var change = new ProgressionChange {
                        ResolutionStart = segment.ReadByte(),
                        ComponentStart = wide ? segment.ReadUInt16() : segment.ReadByte(),
                        LayerEnd = segment.ReadUInt16(),
                        ResolutionEnd = segment.ReadByte(),
                        ComponentEnd = wide ? segment.ReadUInt16() : segment.ReadByte(),
                        Order = segment.ReadByte()
                    };
                    if (change.ComponentEnd == 0) change.ComponentEnd = 256; // CEpoc 0 means 256 (A.6.6).
                    if (change.Order > 4) throw new Jpeg2000Exception("The JPEG 2000 POC progression order is invalid.");
                    if (scope.ProgressionChanges.Count >= 4096) throw new Jpeg2000Exception("The JPEG 2000 codestream has too many progression changes.");
                    scope.ProgressionChanges.Add(change);
                }
                break;
            }
            default:
                // TLM, PLM, PLT, CRG, COM and unknown informational markers do not affect decoding.
                if (marker >= 0xFF70 && marker <= 0xFF7F) throw new Jpeg2000Exception("JPEG 2000 Part 2 extension markers are not supported.");
                break;
        }
    }

    private static int ReadComponentIndex(MarkerReader segment, Codestream codestream) {
        int component = codestream.Components >= 257 ? segment.ReadUInt16() : segment.ReadByte();
        if (component >= codestream.Components) throw new Jpeg2000Exception("A JPEG 2000 marker refers to a missing component.");
        return component;
    }

    // SPcod/SPcoc (A.6.1, A.6.2).
    private static ComponentCoding ReadComponentCoding(MarkerReader segment, bool precincts) {
        var coding = new ComponentCoding {
            Levels = segment.ReadByte(),
            CodeBlockWidth = segment.ReadByte() + 2,
            CodeBlockHeight = segment.ReadByte() + 2,
            Style = segment.ReadByte()
        };
        int transform = segment.ReadByte();
        if (coding.Levels > MaximumLevels || coding.CodeBlockWidth > 10 || coding.CodeBlockHeight > 10 ||
            coding.CodeBlockWidth + coding.CodeBlockHeight > 12 || transform > 1 || (coding.Style & 0xC0) != 0)
            throw new Jpeg2000Exception("The JPEG 2000 coding style is invalid or unsupported.");
        coding.Reversible = transform == 1;
        coding.PrecinctWidths = new int[coding.Levels + 1];
        coding.PrecinctHeights = new int[coding.Levels + 1];
        for (int resolution = 0; resolution <= coding.Levels; resolution++) {
            if (precincts) {
                int value = segment.ReadByte();
                coding.PrecinctWidths[resolution] = value & 0xF;
                coding.PrecinctHeights[resolution] = value >> 4;
                if (resolution > 0 && (coding.PrecinctWidths[resolution] == 0 || coding.PrecinctHeights[resolution] == 0))
                    throw new Jpeg2000Exception("JPEG 2000 precinct sizes below two are only valid at the lowest resolution.");
            } else {
                coding.PrecinctWidths[resolution] = 15;
                coding.PrecinctHeights[resolution] = 15;
            }
        }
        return coding;
    }

    // Sqcd/SPqcd (A.6.4).
    private static Quantization ReadQuantization(MarkerReader segment) {
        int value = segment.ReadByte();
        var quantization = new Quantization { Style = value & 0x1F, GuardBits = value >> 5 };
        if (quantization.Style > 2) throw new Jpeg2000Exception("The JPEG 2000 quantization style is invalid.");
        int count = quantization.Style == 0 ? segment.Remaining : segment.Remaining / 2;
        if (count < 1 || count > 3 * MaximumLevels + 1) throw new Jpeg2000Exception("The JPEG 2000 quantization marker is invalid.");
        quantization.Exponents = new int[count];
        quantization.Mantissas = new int[count];
        for (int band = 0; band < count; band++) {
            if (quantization.Style == 0) {
                quantization.Exponents[band] = segment.ReadByte() >> 3;
            } else {
                int step = segment.ReadUInt16();
                quantization.Exponents[band] = step >> 11;
                quantization.Mantissas[band] = step & 0x7FF;
            }
        }
        return quantization;
    }

    // Tile-parts (A.4.2): SOT, tile-part header, SOD and the packet data that follows.
    private static void ReadTileParts(byte[] bytes, MarkerReader reader, int end, Codestream codestream, Budget budget) {
        int tileCount = codestream.TilesWide * codestream.TilesHigh;
        int tilePartIndex = 0;
        var packedRecords = ReadPackedHeaderRecords(codestream.PackedHeaderData);
        while (reader.Remaining >= 2) {
            budget.Token.ThrowIfCancellationRequested();
            int sotPosition = reader.Position;
            int marker = reader.ReadUInt16();
            if (marker == MarkerEoc) break;
            if (marker != MarkerSot) break; // trailing data after the final tile-part
            if (reader.Remaining < 10 || reader.ReadUInt16() != 10) throw new Jpeg2000Exception("The JPEG 2000 SOT marker is invalid.");
            int tileIndex = reader.ReadUInt16();
            uint partLength = reader.ReadUInt32();
            reader.ReadByte(); // TPsot
            reader.ReadByte(); // TNsot
            if (tileIndex >= tileCount) throw new Jpeg2000Exception("A JPEG 2000 tile-part refers to a missing tile.");
            int partEnd = partLength == 0 || partLength > (uint)(end - sotPosition) ? end : sotPosition + (int)partLength;
            if (partLength != 0 && partLength < 14) throw new Jpeg2000Exception("The JPEG 2000 tile-part length is invalid.");
            if (!codestream.Tiles.TryGetValue(tileIndex, out TileData? tile)) {
                budget.Reserve(256);
                codestream.Tiles[tileIndex] = tile = new TileData();
            }
            bool first = tile.TileParts == 0;
            tile.TileParts++;
            if (tile.TileParts > 65535) throw new Jpeg2000Exception("A JPEG 2000 tile has too many tile-parts.");
            bool sawData = false;
            while (reader.Position + 2 <= partEnd) {
                int headerMarker = reader.ReadUInt16();
                if (headerMarker == MarkerSod) {
                    int bodyEnd = partEnd;
                    // The final tile-part of an unbounded part runs to EOC.
                    if (partLength == 0 && bodyEnd - reader.Position >= 2 && bytes[bodyEnd - 2] == 0xFF && bytes[bodyEnd - 1] == 0xD9) bodyEnd -= 2;
                    budget.Reserve(16);
                    tile.Bodies.Add((reader.Position, bodyEnd));
                    sawData = true;
                    break;
                }
                if (headerMarker >= 0xFF30 && headerMarker <= 0xFF3F) continue;
                int segmentStart = reader.Position;
                if (partEnd - segmentStart < 2) break;
                int length = reader.ReadUInt16();
                if (length < 2 || segmentStart + length > partEnd) throw new Jpeg2000Exception("A JPEG 2000 tile-part marker is truncated.");
                var segment = new MarkerReader(bytes, segmentStart + 2, segmentStart + length);
                if (headerMarker == MarkerPpt) {
                    int index = segment.ReadByte();
                    budget.Reserve(segment.Remaining);
                    var chunk = new byte[segment.Remaining];
                    Buffer.BlockCopy(bytes, segment.Position, chunk, 0, chunk.Length);
                    tile.PacketHeaderChunks.Add((index, chunk));
                } else if (headerMarker is MarkerCod or MarkerCoc or MarkerQcd or MarkerQcc or MarkerRgn or MarkerPoc) {
                    if (!first && headerMarker != MarkerPoc) {
                        // COD, COC, QCD, QCC and RGN belong in the first tile-part only; later copies are ignored.
                    } else {
                        tile.Scope ??= new HeaderScope();
                        ReadScopeMarker(headerMarker, segment, codestream, tile.Scope);
                    }
                } else if (headerMarker < 0xFF30) {
                    throw new Jpeg2000Exception("The JPEG 2000 tile-part header contains invalid data.");
                }
                reader.Position = segmentStart + length;
            }
            if (packedRecords != null) {
                if (tilePartIndex >= packedRecords.Count) throw new Jpeg2000Exception("The JPEG 2000 PPM data has fewer records than tile-parts.");
                tile.PackedHeaders.Add(packedRecords[tilePartIndex]);
            }
            tilePartIndex++;
            if (!sawData && partLength != 0) throw new Jpeg2000Exception("A JPEG 2000 tile-part has no SOD marker.");
            reader.Position = partEnd;
        }
    }

    private static List<(int Start, int End)>? ReadPackedHeaderRecords(byte[]? data) {
        if (data == null) return null;
        var records = new List<(int, int)>();
        int position = 0;
        while (position + 4 <= data.Length) {
            uint length = ReadUInt32(data, position);
            position += 4;
            if (length > (uint)(data.Length - position)) throw new Jpeg2000Exception("A JPEG 2000 PPM record is truncated.");
            records.Add((position, position + (int)length));
            position += (int)length;
        }
        return records;
    }

    private sealed class MarkerReader {
        private readonly byte[] _bytes;
        private readonly int _end;

        internal MarkerReader(byte[] bytes, int start, int end) {
            _bytes = bytes;
            Position = start;
            _end = end;
        }

        internal int Position { get; set; }
        internal int Remaining => _end - Position;

        internal int ReadByte() {
            if (Position >= _end) throw new Jpeg2000Exception("A JPEG 2000 marker segment is truncated.");
            return _bytes[Position++];
        }

        internal int ReadUInt16() {
            if (_end - Position < 2) throw new Jpeg2000Exception("A JPEG 2000 marker segment is truncated.");
            int value = (_bytes[Position] << 8) | _bytes[Position + 1];
            Position += 2;
            return value;
        }

        internal uint ReadUInt32() {
            if (_end - Position < 4) throw new Jpeg2000Exception("A JPEG 2000 marker segment is truncated.");
            uint value = OfficeJpeg2000Decoder.ReadUInt32(_bytes, Position);
            Position += 4;
            return value;
        }
    }
}
