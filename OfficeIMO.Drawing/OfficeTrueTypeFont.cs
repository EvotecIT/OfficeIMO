using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace OfficeIMO.Drawing;

/// <summary>
/// Small managed TrueType/OpenType outline reader used for dependency-free text rasterization.
/// </summary>
/// <remarks>
/// This type reads font files directly and does not call operating-system graphics or font APIs.
/// It supports the simple glyf/cmap/hmtx path needed by OfficeIMO renderers and falls back
/// cleanly when no suitable platform font file is available.
/// </remarks>
public sealed class OfficeTrueTypeFont {
    private static readonly object FontCacheLock = new();
    private static readonly Dictionary<string, OfficeTrueTypeFont?> FontCache = new(StringComparer.OrdinalIgnoreCase);
    private readonly byte[] _data;
    private readonly int _cmap;
    private readonly int _glyf;
    private readonly int _head;
    private readonly int _hhea;
    private readonly int _hmtx;
    private readonly int _gpos;
    private readonly int _kern;
    private readonly int _loca;
    private readonly int _maxp;
    private readonly int _name;
    private readonly int _unitsPerEm;
    private readonly short _ascender;
    private readonly short _descender;
    private readonly ushort _numGlyphs;
    private readonly ushort _numHMetrics;
    private readonly short _indexToLocFormat;
    private readonly int? _collectionIndex;

    private OfficeTrueTypeFont(byte[] data, Dictionary<string, int> tables, int? collectionIndex) {
        _data = data;
        _collectionIndex = collectionIndex;
        _cmap = tables["cmap"];
        _glyf = tables["glyf"];
        _head = tables["head"];
        _hhea = tables["hhea"];
        _hmtx = tables["hmtx"];
        _gpos = tables.TryGetValue("GPOS", out var gpos) ? gpos : -1;
        _kern = tables.TryGetValue("kern", out var kern) ? kern : -1;
        _loca = tables["loca"];
        _maxp = tables["maxp"];
        _name = tables.TryGetValue("name", out var name) ? name : -1;
        _unitsPerEm = ReadUInt16(_data, _head + 18);
        _indexToLocFormat = ReadInt16(_data, _head + 50);
        _ascender = ReadInt16(_data, _hhea + 4);
        _descender = ReadInt16(_data, _hhea + 6);
        _numHMetrics = ReadUInt16(_data, _hhea + 34);
        _numGlyphs = ReadUInt16(_data, _maxp + 4);
    }

    public static OfficeTrueTypeFont? TryLoadDefault() {
        return TryLoadDefault(out _);
    }

    /// <summary>
    /// Attempts to load a common platform font file without using platform font APIs.
    /// </summary>
    public static OfficeTrueTypeFont? TryLoadDefault(out string? resolvedPath) {
        foreach (var path in CandidatePaths()) {
            var font = TryLoadFromPath(path);
            if (font != null && font.HasGlyphs("OfficeIMO 0123456789")) {
                resolvedPath = path;
                return font;
            }
        }

        resolvedPath = null;
        return null;
    }

    public static OfficeTrueTypeFont? TryLoadFromPath(string? path) => TryLoadFromPath(path, null, null);

    public static OfficeTrueTypeFont? TryLoadFromPath(string? path, int? collectionIndex) => TryLoadFromPath(path, collectionIndex, null);

    public static OfficeTrueTypeFont? TryLoadFromPath(string? path, int? collectionIndex, string? faceName) {
        if (string.IsNullOrWhiteSpace(path)) return null;
        try {
            var fullPath = Path.GetFullPath(path);
            var cacheKey = fullPath + "#" + (collectionIndex.HasValue ? collectionIndex.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) : "auto") + "#" + (faceName ?? string.Empty);
            lock (FontCacheLock) {
                if (FontCache.TryGetValue(cacheKey, out var cached)) return cached;
            }

            var font = File.Exists(fullPath) ? TryLoad(File.ReadAllBytes(fullPath), collectionIndex, faceName) : null;
            lock (FontCacheLock) FontCache[cacheKey] = font;
            return font;
        } catch (IOException) {
        } catch (UnauthorizedAccessException) {
        } catch (ArgumentException) {
        } catch (NotSupportedException) {
        } catch (IndexOutOfRangeException) {
        }

        return null;
    }

    public static OfficeTrueTypeFont? TryLoad(byte[] data) => TryLoad(data, null, null);

    public static OfficeTrueTypeFont? TryLoad(byte[] data, int? collectionIndex) => TryLoad(data, collectionIndex, null);

    public static OfficeTrueTypeFont? TryLoad(byte[] data, int? collectionIndex, string? faceName) {
        try {
            if (data.Length < 12) return null;
            var scaler = ReadUInt32(data, 0);
            if (scaler == 0x74746366) {
                var fontCount = ReadUInt32(data, 8);
                if (collectionIndex.HasValue) {
                    if (collectionIndex.Value >= fontCount) return null;
                    var indexedFont = TryLoad(data, CheckedOffset(data, ReadUInt32(data, 12 + collectionIndex.Value * 4)), collectionIndex.Value);
                    return indexedFont != null && indexedFont.MatchesName(faceName) ? indexedFont : null;
                }

                for (var i = 0; i < fontCount; i++) {
                    var directoryOffset = CheckedOffset(data, ReadUInt32(data, 12 + i * 4));
                    var font = TryLoad(data, directoryOffset, (int)i);
                    if (font != null && font.HasGlyphs("OfficeIMO 0123456789") && font.MatchesName(faceName)) return font;
                }

                return null;
            }

            if (collectionIndex.HasValue && collectionIndex.Value > 0) return null;
            var standalone = TryLoad(data, 0, null);
            return standalone != null && standalone.MatchesName(faceName) ? standalone : null;
        } catch (ArgumentOutOfRangeException) {
            return null;
        } catch (IndexOutOfRangeException) {
            return null;
        }
    }

    private static OfficeTrueTypeFont? TryLoad(byte[] data, int directoryOffset, int? collectionIndex) {
        if (directoryOffset < 0 || directoryOffset + 12 > data.Length) return null;
        var scaler = ReadUInt32(data, directoryOffset);
        if (scaler != 0x00010000 && scaler != 0x74727565) return null;
        var count = ReadUInt16(data, directoryOffset + 4);
        var tables = new Dictionary<string, int>(StringComparer.Ordinal);
        for (var i = 0; i < count; i++) {
            var record = directoryOffset + 12 + i * 16;
            if (record + 16 > data.Length) return null;
            var tag = ((char)data[record]).ToString() + (char)data[record + 1] + (char)data[record + 2] + (char)data[record + 3];
            var offset = CheckedOffset(data, ReadUInt32(data, record + 8));
            tables[tag] = offset;
        }

        foreach (var required in new[] { "cmap", "glyf", "head", "hhea", "hmtx", "loca", "maxp" }) if (!tables.ContainsKey(required)) return null;
        return new OfficeTrueTypeFont(data, tables, collectionIndex);
    }

    public double Measure(string text, double fontSize) {
        var scale = ScaleFor(fontSize);
        var width = 0.0;
        ushort? previous = null;
        foreach (var ch in text) {
            var glyph = MapGlyph(ch);
            if (previous.HasValue) width += Kerning(previous.Value, glyph) * scale;
            width += AdvanceWidth(glyph) * scale;
            previous = glyph;
        }
        return width;
    }

    public double LineHeight(double fontSize) {
        return Math.Max(1, _ascender - _descender) * ScaleFor(fontSize);
    }

    /// <summary>
    /// Reads flattened fill contours for the supplied text at a top-left baseline box.
    /// </summary>
    public List<List<OfficePoint>> GetTextContours(string text, double x, double y, double fontSize) {
        var contours = new List<List<OfficePoint>>();
        if (string.IsNullOrEmpty(text)) {
            return contours;
        }

        var scale = ScaleFor(fontSize);
        var cursor = x;
        var baseline = y + _ascender * scale;
        ushort? previous = null;
        foreach (var ch in text) {
            var glyph = MapGlyph(ch);
            if (previous.HasValue) cursor += Kerning(previous.Value, glyph) * scale;
            contours.AddRange(ReadGlyphContours(glyph, new FontTransform(scale, 0, 0, -scale, cursor, baseline), 0));
            cursor += AdvanceWidth(glyph) * scale;
            previous = glyph;
        }

        return contours;
    }

    /// <summary>Best-effort display name read from the font name table.</summary>
    public string? DisplayName => FirstName(4) ?? FirstName(1) ?? FirstName(6) ?? FirstName(2);

    /// <summary>Collection index when the font was loaded from a TrueType collection.</summary>
    public int? CollectionIndex => _collectionIndex;

    private bool HasGlyphs(string value) {
        foreach (var ch in value) {
            if (!char.IsWhiteSpace(ch) && MapGlyph(ch) == 0) return false;
        }

        return true;
    }

    private bool MatchesName(string? faceName) {
        if (string.IsNullOrWhiteSpace(faceName)) return true;
        var requested = faceName!;
        foreach (var name in ReadNames()) {
            if (name.Equals(requested, StringComparison.OrdinalIgnoreCase)) return true;
            if (name.IndexOf(requested, StringComparison.OrdinalIgnoreCase) >= 0) return true;
        }

        return false;
    }

    private IEnumerable<string> ReadNames() {
        if (_name < 0 || _name + 6 > _data.Length) yield break;
        var count = ReadUInt16(_data, _name + 2);
        var stringOffset = _name + ReadUInt16(_data, _name + 4);
        for (var i = 0; i < count; i++) {
            var record = _name + 6 + i * 12;
            if (record + 12 > _data.Length) yield break;
            var nameId = ReadUInt16(_data, record + 6);
            if (nameId != 1 && nameId != 2 && nameId != 4 && nameId != 6) continue;
            var platform = ReadUInt16(_data, record);
            var length = ReadUInt16(_data, record + 8);
            var offset = stringOffset + ReadUInt16(_data, record + 10);
            if (offset < 0 || length == 0 || offset + length > _data.Length) continue;
            var value = DecodeName(platform, offset, length).Trim();
            if (value.Length > 0) yield return value;
        }
    }

    private string DecodeName(ushort platform, int offset, int length) {
        if (platform == 0 || platform == 3) return Encoding.BigEndianUnicode.GetString(_data, offset, length);
        return Encoding.ASCII.GetString(_data, offset, length);
    }

    private string? FirstName(ushort requestedNameId) {
        if (_name < 0 || _name + 6 > _data.Length) return null;
        var count = ReadUInt16(_data, _name + 2);
        var stringOffset = _name + ReadUInt16(_data, _name + 4);
        for (var i = 0; i < count; i++) {
            var record = _name + 6 + i * 12;
            if (record + 12 > _data.Length) return null;
            if (ReadUInt16(_data, record + 6) != requestedNameId) continue;
            var platform = ReadUInt16(_data, record);
            var length = ReadUInt16(_data, record + 8);
            var offset = stringOffset + ReadUInt16(_data, record + 10);
            if (offset < 0 || length == 0 || offset + length > _data.Length) continue;
            var value = DecodeName(platform, offset, length).Trim();
            if (value.Length > 0) return value;
        }

        return null;
    }

    private double ScaleFor(double fontSize) {
        return fontSize / Math.Max(1, _unitsPerEm);
    }

    private ushort MapGlyph(char ch) {
        var cmapOffset = _cmap;
        var subtableCount = ReadUInt16(_data, cmapOffset + 2);
        var best = 0;
        var bestScore = 0;
        for (var i = 0; i < subtableCount; i++) {
            var record = cmapOffset + 4 + i * 8;
            var platform = ReadUInt16(_data, record);
            var encoding = ReadUInt16(_data, record + 2);
            var offset = CheckedOffset(_data, ReadUInt32(_data, record + 4));
            var absolute = cmapOffset + offset;
            if (absolute < 0 || absolute + 2 > _data.Length) continue;
            var format = ReadUInt16(_data, absolute);
            var score = (platform == 3 && encoding == 10 ? 4 : platform == 3 && encoding == 1 ? 3 : platform == 0 ? 2 : 1);
            if ((format == 4 || format == 12) && score > bestScore) {
                best = absolute;
                bestScore = score;
            }
        }

        if (best == 0) return 0;
        var selectedFormat = ReadUInt16(_data, best);
        return selectedFormat == 12 ? MapFormat12(best, ch) : MapFormat4(best, ch);
    }

    private ushort MapFormat4(int table, char ch) {
        var code = ch;
        var segCount = ReadUInt16(_data, table + 6) / 2;
        var endCodes = table + 14;
        var startCodes = endCodes + segCount * 2 + 2;
        var idDeltas = startCodes + segCount * 2;
        var idRangeOffsets = idDeltas + segCount * 2;

        for (var i = 0; i < segCount; i++) {
            var end = ReadUInt16(_data, endCodes + i * 2);
            if (code > end) continue;
            var start = ReadUInt16(_data, startCodes + i * 2);
            if (code < start) return 0;
            var delta = ReadInt16(_data, idDeltas + i * 2);
            var rangeOffset = ReadUInt16(_data, idRangeOffsets + i * 2);
            if (rangeOffset == 0) return (ushort)((code + delta) & 0xffff);
            var glyphOffset = idRangeOffsets + i * 2 + rangeOffset + (code - start) * 2;
            if (glyphOffset < 0 || glyphOffset + 2 > _data.Length) return 0;
            var glyph = ReadUInt16(_data, glyphOffset);
            return glyph == 0 ? (ushort)0 : (ushort)((glyph + delta) & 0xffff);
        }

        return 0;
    }

    private ushort MapFormat12(int table, char ch) {
        var code = ch;
        var groups = ReadUInt32(_data, table + 12);
        var groupOffset = table + 16;
        for (var i = 0; i < groups; i++) {
            var start = ReadUInt32(_data, groupOffset + i * 12);
            var end = ReadUInt32(_data, groupOffset + i * 12 + 4);
            if (code < start || code > end) continue;
            var glyph = ReadUInt32(_data, groupOffset + i * 12 + 8) + code - start;
            return glyph > ushort.MaxValue ? (ushort)0 : (ushort)glyph;
        }

        return 0;
    }

    private int AdvanceWidth(ushort glyph) {
        if (glyph < _numHMetrics) return ReadUInt16(_data, _hmtx + glyph * 4);
        return ReadUInt16(_data, _hmtx + (_numHMetrics - 1) * 4);
    }

    private int Kerning(ushort left, ushort right) {
        return KernPairAdjustment(left, right) + GposPairAdjustment(left, right);
    }

    private int KernPairAdjustment(ushort left, ushort right) {
        if (_kern < 0 || _kern + 4 > _data.Length || ReadUInt16(_data, _kern) != 0) return 0;
        var count = ReadUInt16(_data, _kern + 2);
        var p = _kern + 4;
        var adjustment = 0;
        for (var table = 0; table < count; table++) {
            if (p + 6 > _data.Length) break;
            var length = ReadUInt16(_data, p + 2);
            var coverage = ReadUInt16(_data, p + 4);
            var next = p + length;
            if (length < 14 || next <= p || next > _data.Length) break;
            if ((coverage >> 8) == 0) adjustment += KerningFormat0(p, left, right);
            p = next;
        }

        return adjustment;
    }

    private int KerningFormat0(int table, ushort left, ushort right) {
        var pairs = ReadUInt16(_data, table + 6);
        var pairOffset = table + 14;
        var key = ((uint)left << 16) | right;
        var low = 0;
        var high = pairs - 1;
        while (low <= high) {
            var mid = low + (high - low) / 2;
            var record = pairOffset + mid * 6;
            if (record + 6 > _data.Length) return 0;
            var candidate = (ReadUInt32(_data, record) & 0xffffffffu);
            if (candidate == key) return ReadInt16(_data, record + 4);
            if (candidate < key) low = mid + 1;
            else high = mid - 1;
        }

        return 0;
    }

    private int GposPairAdjustment(ushort left, ushort right) {
        if (_gpos < 0 || !InBounds(_gpos, 10) || ReadUInt16(_data, _gpos) != 1) return 0;
        var featureList = _gpos + ReadUInt16(_data, _gpos + 6);
        var lookupList = _gpos + ReadUInt16(_data, _gpos + 8);
        if (!InBounds(featureList, 2) || !InBounds(lookupList, 2)) return 0;

        var adjustment = 0;
        var seen = new HashSet<ushort>();
        foreach (var lookupIndex in GposFeatureLookupIndexes(featureList, "kern")) {
            if (seen.Add(lookupIndex)) adjustment += GposPairAdjustmentFromLookup(lookupList, lookupIndex, left, right);
        }

        return adjustment;
    }

    private IEnumerable<ushort> GposFeatureLookupIndexes(int featureList, string featureTag) {
        var featureCount = ReadUInt16(_data, featureList);
        for (var i = 0; i < featureCount; i++) {
            var record = featureList + 2 + i * 6;
            if (!InBounds(record, 6)) yield break;
            if (!TagEquals(record, featureTag)) continue;
            var feature = featureList + ReadUInt16(_data, record + 4);
            if (!InBounds(feature, 4)) yield break;
            var lookupCount = ReadUInt16(_data, feature + 2);
            for (var lookup = 0; lookup < lookupCount; lookup++) {
                var indexOffset = feature + 4 + lookup * 2;
                if (!InBounds(indexOffset, 2)) yield break;
                yield return ReadUInt16(_data, indexOffset);
            }
        }
    }

    private int GposPairAdjustmentFromLookup(int lookupList, ushort lookupIndex, ushort left, ushort right) {
        var lookupCount = ReadUInt16(_data, lookupList);
        if (lookupIndex >= lookupCount) return 0;
        var lookupOffset = lookupList + 2 + lookupIndex * 2;
        if (!InBounds(lookupOffset, 2)) return 0;
        var lookup = lookupList + ReadUInt16(_data, lookupOffset);
        if (!InBounds(lookup, 6) || ReadUInt16(_data, lookup) != 2) return 0;

        var adjustment = 0;
        var subtableCount = ReadUInt16(_data, lookup + 4);
        for (var i = 0; i < subtableCount; i++) {
            var subtableOffset = lookup + 6 + i * 2;
            if (!InBounds(subtableOffset, 2)) break;
            adjustment += GposPairAdjustmentFromSubtable(lookup + ReadUInt16(_data, subtableOffset), left, right);
        }

        return adjustment;
    }

    private int GposPairAdjustmentFromSubtable(int subtable, ushort left, ushort right) {
        if (!InBounds(subtable, 10) || ReadUInt16(_data, subtable) != 1) return 0;
        var coverage = subtable + ReadUInt16(_data, subtable + 2);
        var valueFormat1 = ReadUInt16(_data, subtable + 4);
        var valueFormat2 = ReadUInt16(_data, subtable + 6);
        var pairSetCount = ReadUInt16(_data, subtable + 8);
        var coverageIndex = CoverageIndex(coverage, left);
        if (coverageIndex < 0 || coverageIndex >= pairSetCount) return 0;

        var pairSetOffset = subtable + 10 + coverageIndex * 2;
        if (!InBounds(pairSetOffset, 2)) return 0;
        var pairSet = subtable + ReadUInt16(_data, pairSetOffset);
        if (!InBounds(pairSet, 2)) return 0;

        var value1Size = ValueRecordSize(valueFormat1);
        var value2Size = ValueRecordSize(valueFormat2);
        var recordSize = 2 + value1Size + value2Size;
        var low = 0;
        var high = ReadUInt16(_data, pairSet) - 1;
        while (low <= high) {
            var mid = low + (high - low) / 2;
            var record = pairSet + 2 + mid * recordSize;
            if (!InBounds(record, recordSize)) return 0;
            var candidate = ReadUInt16(_data, record);
            if (candidate == right) return ReadValueRecordXAdvance(record + 2, valueFormat1);
            if (candidate < right) low = mid + 1;
            else high = mid - 1;
        }

        return 0;
    }

    private int CoverageIndex(int coverage, ushort glyph) {
        if (!InBounds(coverage, 4)) return -1;
        var format = ReadUInt16(_data, coverage);
        if (format == 1) {
            var count = ReadUInt16(_data, coverage + 2);
            var low = 0;
            var high = count - 1;
            while (low <= high) {
                var mid = low + (high - low) / 2;
                var offset = coverage + 4 + mid * 2;
                if (!InBounds(offset, 2)) return -1;
                var candidate = ReadUInt16(_data, offset);
                if (candidate == glyph) return mid;
                if (candidate < glyph) low = mid + 1;
                else high = mid - 1;
            }

            return -1;
        }

        if (format != 2) return -1;
        var rangeCount = ReadUInt16(_data, coverage + 2);
        for (var i = 0; i < rangeCount; i++) {
            var range = coverage + 4 + i * 6;
            if (!InBounds(range, 6)) return -1;
            var start = ReadUInt16(_data, range);
            var end = ReadUInt16(_data, range + 2);
            if (glyph < start || glyph > end) continue;
            return ReadUInt16(_data, range + 4) + glyph - start;
        }

        return -1;
    }

    private int ReadValueRecordXAdvance(int offset, ushort valueFormat) {
        if ((valueFormat & 0x0001) != 0) offset += 2;
        if ((valueFormat & 0x0002) != 0) offset += 2;
        if ((valueFormat & 0x0004) == 0) return 0;
        return InBounds(offset, 2) ? ReadInt16(_data, offset) : 0;
    }

    private static int ValueRecordSize(ushort valueFormat) {
        var size = 0;
        for (var bit = 1; bit <= 0x0080; bit <<= 1) if ((valueFormat & bit) != 0) size += 2;
        return size;
    }

    private bool TagEquals(int offset, string tag) {
        return InBounds(offset, 4) && _data[offset] == tag[0] && _data[offset + 1] == tag[1] && _data[offset + 2] == tag[2] && _data[offset + 3] == tag[3];
    }

    private List<List<OfficePoint>> ReadGlyphContours(ushort glyph, FontTransform transform, int depth) {
        var contours = new List<List<OfficePoint>>();
        if (glyph == 0 || glyph >= _numGlyphs || depth > 8) return contours;
        var glyphStart = GlyphOffset(glyph);
        var glyphEnd = GlyphOffset((ushort)(glyph + 1));
        if (glyphStart == glyphEnd) return contours;
        var offset = _glyf + glyphStart;
        if (offset + 10 > _data.Length) return contours;
        var contourCount = ReadInt16(_data, offset);
        if (contourCount < 0) {
            ReadCompositeGlyphContours(offset, transform, depth, contours);
            return contours;
        }

        if (contourCount <= 0) return contours;

        var endPts = new ushort[contourCount];
        for (var i = 0; i < contourCount; i++) endPts[i] = ReadUInt16(_data, offset + 10 + i * 2);
        var pointCount = endPts[contourCount - 1] + 1;
        var instructionLengthOffset = offset + 10 + contourCount * 2;
        var instructionLength = ReadUInt16(_data, instructionLengthOffset);
        var p = instructionLengthOffset + 2 + instructionLength;
        var flags = new byte[pointCount];
        for (var i = 0; i < pointCount; i++) {
            var flag = _data[p++];
            flags[i] = flag;
            if ((flag & 8) == 0) continue;
            var repeat = _data[p++];
            for (var r = 0; r < repeat && i + 1 < pointCount; r++) flags[++i] = flag;
        }

        var xs = new short[pointCount];
        DecodeCoordinates(_data, flags, xs, ref p, true);
        var ys = new short[pointCount];
        DecodeCoordinates(_data, flags, ys, ref p, false);

        var start = 0;
        for (var c = 0; c < contourCount; c++) {
            var end = endPts[c];
            var points = new List<GlyphPoint>();
            for (var i = start; i <= end; i++) {
                var point = transform.Apply(xs[i], ys[i]);
                points.Add(new GlyphPoint(point.X, point.Y, (flags[i] & 1) != 0));
            }

            AddFlattenedContour(points, contours);
            start = end + 1;
        }

        return contours;
    }

    private void ReadCompositeGlyphContours(int glyphOffset, FontTransform transform, int depth, List<List<OfficePoint>> contours) {
        const ushort argWords = 1;
        const ushort argsAreXy = 2;
        const ushort haveScale = 8;
        const ushort moreComponents = 32;
        const ushort haveXyScale = 64;
        const ushort haveTwoByTwo = 128;

        var p = glyphOffset + 10;
        ushort flags;
        do {
            if (p + 4 > _data.Length) return;
            flags = ReadUInt16(_data, p);
            var componentGlyph = ReadUInt16(_data, p + 2);
            p += 4;
            double arg1;
            double arg2;
            if ((flags & argWords) != 0) {
                if (p + 4 > _data.Length) return;
                arg1 = ReadInt16(_data, p);
                arg2 = ReadInt16(_data, p + 2);
                p += 4;
            } else {
                if (p + 2 > _data.Length) return;
                arg1 = (sbyte)_data[p];
                arg2 = (sbyte)_data[p + 1];
                p += 2;
            }

            var dx = (flags & argsAreXy) != 0 ? arg1 : 0;
            var dy = (flags & argsAreXy) != 0 ? arg2 : 0;
            var a = 1.0;
            var b = 0.0;
            var c = 0.0;
            var d = 1.0;
            if ((flags & haveScale) != 0) {
                if (p + 2 > _data.Length) return;
                a = d = ReadF2Dot14(_data, p);
                p += 2;
            } else if ((flags & haveXyScale) != 0) {
                if (p + 4 > _data.Length) return;
                a = ReadF2Dot14(_data, p);
                d = ReadF2Dot14(_data, p + 2);
                p += 4;
            } else if ((flags & haveTwoByTwo) != 0) {
                if (p + 8 > _data.Length) return;
                a = ReadF2Dot14(_data, p);
                b = ReadF2Dot14(_data, p + 2);
                c = ReadF2Dot14(_data, p + 4);
                d = ReadF2Dot14(_data, p + 6);
                p += 8;
            }

            contours.AddRange(ReadGlyphContours(componentGlyph, transform.Compose(a, b, c, d, dx, dy), depth + 1));
        } while ((flags & moreComponents) != 0);
    }

    private int GlyphOffset(ushort glyph) {
        if (_indexToLocFormat == 0) return ReadUInt16(_data, _loca + glyph * 2) * 2;
        return CheckedOffset(_data, ReadUInt32(_data, _loca + glyph * 4));
    }

    private static void DecodeCoordinates(byte[] data, byte[] flags, short[] values, ref int p, bool xAxis) {
        var shortFlag = xAxis ? 2 : 4;
        var sameOrPositiveFlag = xAxis ? 16 : 32;
        var current = 0;
        for (var i = 0; i < flags.Length; i++) {
            var flag = flags[i];
            int delta;
            if ((flag & shortFlag) != 0) {
                delta = data[p++];
                if ((flag & sameOrPositiveFlag) == 0) delta = -delta;
            } else if ((flag & sameOrPositiveFlag) != 0) {
                delta = 0;
            } else {
                delta = ReadInt16(data, p);
                p += 2;
            }

            current += delta;
            values[i] = (short)current;
        }
    }

    private static void AddFlattenedContour(List<GlyphPoint> source, List<List<OfficePoint>> contours) {
        if (source.Count == 0) return;
        var result = new List<OfficePoint>();
        var last = source[source.Count - 1];
        var first = source[0];
        var current = first.OnCurve ? first : last.OnCurve ? last : Mid(last, first);
        result.Add(current.Point);
        var index = first.OnCurve ? 1 : 0;

        while (index < source.Count) {
            var point = source[index % source.Count];
            if (point.OnCurve) {
                result.Add(point.Point);
                current = point;
                index++;
                continue;
            }

            var next = source[(index + 1) % source.Count];
            var end = next.OnCurve ? next : Mid(point, next);
            FlattenQuadratic(current, point, end, result);
            current = end;
            index += next.OnCurve ? 2 : 1;
        }

        if (result.Count >= 3) contours.Add(result);
    }

    private static void FlattenQuadratic(GlyphPoint start, GlyphPoint control, GlyphPoint end, List<OfficePoint> output) {
        var chord = Math.Sqrt((end.X - start.X) * (end.X - start.X) + (end.Y - start.Y) * (end.Y - start.Y));
        var bend = Math.Sqrt((start.X - 2 * control.X + end.X) * (start.X - 2 * control.X + end.X) + (start.Y - 2 * control.Y + end.Y) * (start.Y - 2 * control.Y + end.Y));
        var steps = Math.Max(6, Math.Min(18, (int)Math.Ceiling((chord + bend * 2.0) / 120.0)));
        for (var i = 1; i <= steps; i++) {
            var t = i / (double)steps;
            var mt = 1 - t;
            output.Add(new OfficePoint(mt * mt * start.X + 2 * mt * t * control.X + t * t * end.X, mt * mt * start.Y + 2 * mt * t * control.Y + t * t * end.Y));
        }
    }

    private static GlyphPoint Mid(GlyphPoint left, GlyphPoint right) => new((left.X + right.X) / 2.0, (left.Y + right.Y) / 2.0, true);

    private static IEnumerable<string> CandidatePaths() {
        yield return "/System/Library/Fonts/SFNS.ttf";
        yield return "/System/Library/Fonts/SFCompact.ttf";
        yield return "/System/Library/Fonts/HelveticaNeue.ttc";
        yield return "/System/Library/Fonts/Geneva.ttf";
        yield return "/Library/Fonts/Arial.ttf";
        yield return "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf";
        yield return "/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf";
        var windows = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
        if (!string.IsNullOrEmpty(windows)) {
            yield return Path.Combine(windows, "Fonts", "arial.ttf");
            yield return Path.Combine(windows, "Fonts", "segoeui.ttf");
        }
    }

    private static ushort ReadUInt16(byte[] data, int offset) => (ushort)((data[offset] << 8) | data[offset + 1]);
    private static short ReadInt16(byte[] data, int offset) => (short)ReadUInt16(data, offset);
    private static double ReadF2Dot14(byte[] data, int offset) => ReadInt16(data, offset) / 16384.0;
    private static uint ReadUInt32(byte[] data, int offset) => ((uint)data[offset] << 24) | ((uint)data[offset + 1] << 16) | ((uint)data[offset + 2] << 8) | data[offset + 3];
    private bool InBounds(int offset, int length) => offset >= 0 && length >= 0 && offset <= _data.Length - length;
    private static int CheckedOffset(byte[] data, uint offset) {
        if (offset > int.MaxValue || offset >= data.Length) throw new ArgumentOutOfRangeException(nameof(offset));
        return (int)offset;
    }

    private static string? FullPathOrNull(string? path) {
        if (string.IsNullOrWhiteSpace(path)) return null;
        try {
            return Path.GetFullPath(path);
        } catch (ArgumentException) {
        } catch (NotSupportedException) {
        }

        return path;
    }

    private readonly struct FontTransform {
        public FontTransform(double xx, double xy, double yx, double yy, double dx, double dy) {
            Xx = xx;
            Xy = xy;
            Yx = yx;
            Yy = yy;
            Dx = dx;
            Dy = dy;
        }

        private double Xx { get; }
        private double Xy { get; }
        private double Yx { get; }
        private double Yy { get; }
        private double Dx { get; }
        private double Dy { get; }

        public OfficePoint Apply(double x, double y) => new(Dx + Xx * x + Xy * y, Dy + Yx * x + Yy * y);

        public FontTransform Compose(double xx, double xy, double yx, double yy, double dx, double dy) {
            return new FontTransform(
                Xx * xx + Xy * yx,
                Xx * xy + Xy * yy,
                Yx * xx + Yy * yx,
                Yx * xy + Yy * yy,
                Dx + Xx * dx + Xy * dy,
                Dy + Yx * dx + Yy * dy);
        }
    }

    private readonly struct GlyphPoint {
        public GlyphPoint(double x, double y, bool onCurve) {
            X = x;
            Y = y;
            OnCurve = onCurve;
            Point = new OfficePoint(x, y);
        }

        public double X { get; }
        public double Y { get; }
        public bool OnCurve { get; }
        public OfficePoint Point { get; }
    }
}
