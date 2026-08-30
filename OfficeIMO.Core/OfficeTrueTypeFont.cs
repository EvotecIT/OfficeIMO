using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>
/// Small managed TrueType/OpenType outline reader used for dependency-free text rasterization.
/// </summary>
/// <remarks>
/// This type reads font files directly and does not call operating-system graphics or font APIs.
/// It supports the simple glyf/cmap/hmtx path needed by OfficeIMO renderers and falls back
/// cleanly when no suitable platform font file is available.
/// </remarks>
public sealed partial class OfficeTrueTypeFont : IOfficeBoundedFontProgram, IOfficeFontBaselineMetrics, IOfficeVariableFontProgram {
    private const uint MaxTrueTypeCollectionFonts = 256;
    private const int MaxFontTableRecords = 512;
    private const int MaxFontCacheEntries = 1024;
    private static readonly object FontCacheLock = new();
    private static readonly Dictionary<string, OfficeTrueTypeFont?> FontCache = new(StringComparer.OrdinalIgnoreCase);
    private static readonly Dictionary<string, FontFamilyResolution> FontFamilyCache = new(StringComparer.OrdinalIgnoreCase);
    private readonly byte[] _data;
    private readonly int _cmap;
    private readonly int _cmapLength;
    private readonly int _glyf;
    private readonly int _head;
    private readonly int _hhea;
    private readonly int _hmtx;
    private readonly OfficeOpenTypeKerning _kerning;
    private readonly int _loca;
    private readonly int _maxp;
    private readonly int _name;
    private readonly HashSet<int> _validFormat4Subtables;
    private readonly HashSet<int> _validFormat12Subtables;
    private readonly OfficeTrueTypeVariations? _variations;
    private readonly OfficeFontVariationModel _variationModel;
    private readonly int _unitsPerEm;
    private readonly int _ascender;
    private readonly int _descender;
    private readonly int _lineGap;
    private readonly ushort _numGlyphs;
    private readonly ushort _numHMetrics;
    private readonly short _indexToLocFormat;
    private readonly int? _collectionIndex;
    private readonly string _fingerprint;

    private OfficeTrueTypeFont(
        byte[] data,
        Dictionary<string, int> tables,
        IReadOnlyDictionary<string, int> tableLengths,
        int? collectionIndex,
        OfficeFontVariationModel? variationModel = null) {
        _data = data;
        _fingerprint = ComputeFingerprint(data, collectionIndex)
            + (variationModel != null && variationModel.IsVariable ? ":axes=" + variationModel.Identity : string.Empty);
        _collectionIndex = collectionIndex;
        _variationModel = variationModel ?? OfficeFontVariationModel.None;
        _cmap = tables["cmap"];
        _cmapLength = tableLengths["cmap"];
        _glyf = tables["glyf"];
        _head = tables["head"];
        _hhea = tables["hhea"];
        _hmtx = tables["hmtx"];
        int gpos = tables.TryGetValue("GPOS", out int gposOffset) ? gposOffset : -1;
        int kern = tables.TryGetValue("kern", out int kernOffset) ? kernOffset : -1;
        OfficeOpenTypeReader? reader = null;
        if (_variationModel.IsVariable) {
            reader = OfficeOpenTypeReader.TryCreate(data)
                ?? throw new InvalidDataException("The variable TrueType font table directory is invalid.");
        }
        _kerning = reader != null
            ? OfficeOpenTypeKerning.FromReader(reader, _variationModel)
            : new OfficeOpenTypeKerning(_data, kern, gpos, includeExtendedGpos: true);
        _loca = tables["loca"];
        _maxp = tables["maxp"];
        _name = tables.TryGetValue("name", out var name) ? name : -1;
        _validFormat4Subtables = OfficeOpenTypeCmap.CollectValidFormat4Subtables(
            _data,
            _cmap,
            _cmapLength,
            OfficeOpenTypeCmap.MaximumSubtables);
        _validFormat12Subtables = OfficeOpenTypeCmap.CollectValidFormat12Subtables(
            _data,
            _cmap,
            _cmapLength,
            OfficeOpenTypeCmap.MaximumSubtables,
            OfficeOpenTypeCmap.MaximumFormat12Groups);
        _unitsPerEm = ReadUInt16(_data, _head + 18);
        _indexToLocFormat = ReadInt16(_data, _head + 50);
        OfficeOpenTypeMvarMetrics? mvar = reader != null
            ? OfficeOpenTypeMvarMetrics.TryParse(reader, _variationModel)
            : null;
        _ascender = checked(ReadInt16(_data, _hhea + 4) + (mvar?.HorizontalAscenderDelta ?? 0));
        _descender = checked(ReadInt16(_data, _hhea + 6) + (mvar?.HorizontalDescenderDelta ?? 0));
        _lineGap = checked(ReadInt16(_data, _hhea + 8) + (mvar?.HorizontalLineGapDelta ?? 0));
        _numHMetrics = ReadUInt16(_data, _hhea + 34);
        _numGlyphs = ReadUInt16(_data, _maxp + 4);
        _variations = variationModel != null && variationModel.IsVariable
            ? OfficeTrueTypeVariations.Parse(data, tables, variationModel, _numGlyphs)
            : null;
    }

    /// <inheritdoc />
    public string Fingerprint => _fingerprint;
    IReadOnlyDictionary<string, float> IOfficeVariableFontProgram.VariationCoordinatesForShaping =>
        _variationModel?.DesignCoordinates ?? OfficeFontVariationModel.None.DesignCoordinates;

    private static string ComputeFingerprint(byte[] data, int? collectionIndex) {
        using HashAlgorithm hash = SHA256.Create();
        byte[] digest = hash.ComputeHash(data);
        return "sha256:" + ToLowerHex(digest) + ":face="
            + (collectionIndex.HasValue
                ? collectionIndex.Value.ToString(System.Globalization.CultureInfo.InvariantCulture)
                : "auto");
    }

    private static string ToLowerHex(byte[] bytes) {
        const string alphabet = "0123456789abcdef";
        var result = new char[bytes.Length * 2];
        for (int index = 0; index < bytes.Length; index++) {
            result[index * 2] = alphabet[bytes[index] >> 4];
            result[index * 2 + 1] = alphabet[bytes[index] & 0x0F];
        }
        return new string(result);
    }

    public static OfficeTrueTypeFont? TryLoadDefault() {
        return TryLoadDefault(out _);
    }

    /// <summary>
    /// Attempts to load a common platform font file without using platform font APIs.
    /// </summary>
    public static OfficeTrueTypeFont? TryLoadDefault(out string? resolvedPath) {
        foreach (var path in CandidatePaths()) {
            var font = TryLoad(path);
            if (font != null && font.HasGlyphs("OfficeIMO 0123456789")) {
                resolvedPath = path;
                return font;
            }
        }

        resolvedPath = null;
        return null;
    }

    /// <summary>
    /// Attempts to load the first resolvable font from a CSS/Office-style font-family fallback list.
    /// </summary>
    public static OfficeTrueTypeFont? TryLoadFontFamily(string? fontFamily) {
        return TryLoadFontFamily(fontFamily, out _);
    }

    /// <summary>
    /// Attempts to load the first resolvable font from a CSS/Office-style font-family fallback list.
    /// </summary>
    public static OfficeTrueTypeFont? TryLoadFontFamily(string? fontFamily, out string? resolvedPath) {
        string cacheKey = NormalizeFontFamilyCacheKey(fontFamily);
        lock (FontCacheLock) {
            if (FontFamilyCache.TryGetValue(cacheKey, out FontFamilyResolution cached)) {
                resolvedPath = cached.Path;
                return cached.Font;
            }
        }

        FontFamilyResolution resolved = ResolveFontFamily(fontFamily);
        lock (FontCacheLock) {
            if (FontFamilyCache.Count >= MaxFontCacheEntries) FontFamilyCache.Clear();
            FontFamilyCache[cacheKey] = resolved;
        }

        resolvedPath = resolved.Path;
        return resolved.Font;
    }

    public static OfficeTrueTypeFont? TryLoad(string? path) => TryLoad(path, null, null);

    public static OfficeTrueTypeFont? TryLoad(string? path, int? collectionIndex) => TryLoad(path, collectionIndex, null);

    public static OfficeTrueTypeFont? TryLoad(string? path, int? collectionIndex, string? faceName) {
        if (string.IsNullOrWhiteSpace(path)) return null;
        try {
            var fullPath = Path.GetFullPath(path);
            var cacheKey = fullPath + "#" + (collectionIndex.HasValue ? collectionIndex.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) : "auto") + "#" + (faceName ?? string.Empty);
            lock (FontCacheLock) {
                if (FontCache.TryGetValue(cacheKey, out var cached)) return cached;
            }

            var font = File.Exists(fullPath) ? TryLoad(File.ReadAllBytes(fullPath), collectionIndex, faceName) : null;
            lock (FontCacheLock) {
                if (FontCache.Count >= MaxFontCacheEntries) FontCache.Clear();
                FontCache[cacheKey] = font;
            }
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
                if (fontCount == 0 || fontCount > MaxTrueTypeCollectionFonts) return null;
                if (data.Length < 12 + fontCount * 4) return null;
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

    internal static OfficeTrueTypeFont? TryLoad(
        byte[] data,
        OfficeFontVariationModel variationModel,
        out string? error) {
        if (variationModel == null) throw new ArgumentNullException(nameof(variationModel));
        error = null;
        try {
            return TryLoad(data, 0, null, variationModel);
        } catch (Exception exception) when (exception is InvalidDataException
                                            || exception is NotSupportedException
                                            || exception is OverflowException
                                            || exception is ArgumentOutOfRangeException
                                            || exception is IndexOutOfRangeException) {
            error = exception.Message;
            return null;
        }
    }

    private static OfficeTrueTypeFont? TryLoad(byte[] data, int directoryOffset, int? collectionIndex) {
        return TryLoad(data, directoryOffset, collectionIndex, null);
    }

    private static OfficeTrueTypeFont? TryLoad(
        byte[] data,
        int directoryOffset,
        int? collectionIndex,
        OfficeFontVariationModel? variationModel) {
        if (directoryOffset < 0 || directoryOffset + 12 > data.Length) return null;
        var scaler = ReadUInt32(data, directoryOffset);
        if (scaler != 0x00010000 && scaler != 0x74727565) return null;
        var count = ReadUInt16(data, directoryOffset + 4);
        if (count == 0 || count > MaxFontTableRecords) return null;
        if (data.Length < directoryOffset + 12 + count * 16) return null;
        var tables = new Dictionary<string, int>(StringComparer.Ordinal);
        var tableLengths = new Dictionary<string, int>(StringComparer.Ordinal);
        for (var i = 0; i < count; i++) {
            var record = directoryOffset + 12 + i * 16;
            if (record + 16 > data.Length) return null;
            var tag = ((char)data[record]).ToString() + (char)data[record + 1] + (char)data[record + 2] + (char)data[record + 3];
            uint offsetValue = ReadUInt32(data, record + 8);
            uint lengthValue = ReadUInt32(data, record + 12);
            if (offsetValue > int.MaxValue || lengthValue > int.MaxValue || tables.ContainsKey(tag)) return null;
            var offset = CheckedOffset(data, offsetValue);
            int length = checked((int)lengthValue);
            if (offset > data.Length - length) return null;
            tables[tag] = offset;
            tableLengths[tag] = length;
        }

        foreach (var required in new[] { "cmap", "glyf", "head", "hhea", "hmtx", "loca", "maxp" }) if (!tables.ContainsKey(required)) return null;
        return new OfficeTrueTypeFont(data, tables, tableLengths, collectionIndex, variationModel);
    }

    public double Measure(string text, double fontSize) {
        var scale = ScaleFor(fontSize);
        var width = 0.0;
        OfficeTrueTypeVariations.WorkBudget? variationWorkBudget = _variations?.CreateWorkBudget();
        var glyphs = new List<int>();
        var scalars = new List<int>();
        for (int index = 0; index < text.Length;) {
            int glyph = ReadMappedGlyph(text, ref index, out int scalar);
            if (glyph < 0) continue;
            glyphs.Add(glyph);
            scalars.Add(scalar);
        }
        OfficeOpenTypeGlyphPositioning[] positioning = _kerning.PositionRun(glyphs, scalars);
        for (int index = 0; index < glyphs.Count; index++) {
            width += checked(AdvanceWidth((ushort)glyphs[index], variationWorkBudget, CancellationToken.None) +
                             positioning[index].XAdvance) * scale;
        }
        return width;
    }

    internal IReadOnlyList<double> MeasureTextElements(IReadOnlyList<string> elements, double fontSize) {
        var widths = new double[elements.Count];
        double scale = ScaleFor(fontSize);
        OfficeTrueTypeVariations.WorkBudget? variationWorkBudget = _variations?.CreateWorkBudget();
        var glyphs = new List<int>();
        var scalars = new List<int>();
        var elementIndexes = new List<int>();
        for (int elementIndex = 0; elementIndex < elements.Count; elementIndex++) {
            string text = elements[elementIndex];
            for (int textIndex = 0; textIndex < text.Length;) {
                int glyph = ReadMappedGlyph(text, ref textIndex, out int scalar);
                if (glyph < 0) continue;
                glyphs.Add(glyph);
                scalars.Add(scalar);
                elementIndexes.Add(elementIndex);
            }
        }
        OfficeOpenTypeGlyphPositioning[] positioning = _kerning.PositionRun(glyphs, scalars);
        for (int index = 0; index < glyphs.Count; index++) {
            widths[elementIndexes[index]] += checked(
                AdvanceWidth((ushort)glyphs[index], variationWorkBudget, CancellationToken.None) +
                positioning[index].XAdvance) * scale;
        }
        return widths;
    }

    public double LineHeight(double fontSize) {
        return Math.Max(1, _ascender - _descender) * ScaleFor(fontSize);
    }

    /// <inheritdoc />
    public double BaselineOffset(double fontSize) => _ascender * ScaleFor(fontSize);

    internal double LineSpacingRatio =>
        Math.Max(1, _ascender - _descender + _lineGap) / (double)_unitsPerEm;

    /// <summary>
    /// Reads flattened fill contours for the supplied text at a top-left baseline box.
    /// </summary>
    public List<List<OfficePoint>> GetTextContours(string text, double x, double y, double fontSize) {
        return GetTextContoursBounded(text, x, y, fontSize, int.MaxValue, CancellationToken.None);
    }

    /// <inheritdoc />
    public List<List<OfficePoint>> GetTextContoursBounded(
        string text,
        double x,
        double y,
        double fontSize,
        int maximumPointCount,
        CancellationToken cancellationToken) {
        if (maximumPointCount <= 0) throw new ArgumentOutOfRangeException(nameof(maximumPointCount));
        var contours = new List<List<OfficePoint>>();
        if (string.IsNullOrEmpty(text)) {
            return contours;
        }

        var scale = ScaleFor(fontSize);
        var cursor = x;
        var baseline = y + _ascender * scale;
        int pointCount = 0;
        OfficeTrueTypeVariations.WorkBudget? variationWorkBudget = _variations?.CreateWorkBudget();
        var glyphs = new List<(ushort Glyph, int Scalar)>();
        for (int index = 0; index < text.Length;) {
            cancellationToken.ThrowIfCancellationRequested();
            int glyph = ReadMappedGlyph(text, ref index, out int scalar);
            if (glyph < 0) continue;
            glyphs.Add((checked((ushort)glyph), scalar));
        }

        var glyphIds = new List<int>(glyphs.Count);
        var scalars = new List<int>(glyphs.Count);
        foreach ((ushort glyph, int scalar) in glyphs) {
            glyphIds.Add(glyph);
            scalars.Add(scalar);
        }
        OfficeOpenTypeGlyphPositioning[] positioning = _kerning.PositionRun(glyphIds, scalars);

        for (int index = 0; index < glyphs.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            ushort glyph = glyphs[index].Glyph;
            double glyphX = cursor + positioning[index].XPlacement * scale;
            List<List<OfficePoint>> glyphContours = ReadGlyphContours(
                glyph,
                new FontTransform(scale, 0, 0, -scale, glyphX, baseline),
                0,
                variationWorkBudget,
                maximumPointCount,
                ref pointCount,
                cancellationToken,
                attachmentPoints: null);
            contours.AddRange(glyphContours);
            int positionedAdvance = checked(
                AdvanceWidth(glyph, variationWorkBudget, cancellationToken) +
                positioning[index].XAdvance);
            cursor += positionedAdvance * scale;
        }

        return contours;
    }

    /// <summary>Best-effort display name read from the font name table.</summary>
    public string? DisplayName => FirstName(4) ?? FirstName(1) ?? FirstName(6) ?? FirstName(2);

    /// <summary>Collection index when the font was loaded from a TrueType collection.</summary>
    public int? CollectionIndex => _collectionIndex;

    internal bool HasGlyphs(string value) {
        return OfficeOpenTypeCmap.HasGlyphs(
            value,
            scalar => MapGlyph(scalar),
            MapVariationGlyph);
    }

    private int ReadMappedGlyph(string text, ref int index, out int scalar) =>
        OfficeOpenTypeCmap.ReadMappedGlyph(text, ref index, scalarValue => MapGlyph(scalarValue), MapVariationGlyph, out scalar);

    private int MapVariationGlyph(int scalar, int selector) => OfficeOpenTypeCmap.MapVariationSequence(
        _data,
        _cmap,
        _cmapLength,
        _numGlyphs,
        scalar,
        selector,
        mappedScalar => MapGlyph(mappedScalar));

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

    private static int ReadScalar(string value, ref int index) {
        char first = value[index++];
        if (char.IsHighSurrogate(first) && index < value.Length && char.IsLowSurrogate(value[index])) {
            return char.ConvertToUtf32(first, value[index++]);
        }

        return first;
    }

    private static bool IsWhitespaceScalar(int scalar) => scalar <= char.MaxValue && char.IsWhiteSpace((char)scalar);

    private ushort MapGlyph(int scalar) {
        if (scalar < 0 || scalar > 0x10FFFF) return 0;
        var cmapOffset = _cmap;
        int cmapEnd = checked(cmapOffset + _cmapLength);
        if (_cmapLength < 4) return 0;
        var subtableCount = ReadUInt16(_data, cmapOffset + 2);
        if (subtableCount == 0 || subtableCount > OfficeOpenTypeCmap.MaximumSubtables) return 0;
        if (cmapOffset + 4 > cmapEnd - subtableCount * 8) return 0;
        var best = 0;
        var bestScore = 0;
        for (var i = 0; i < subtableCount; i++) {
            var record = cmapOffset + 4 + i * 8;
            var platform = ReadUInt16(_data, record);
            var encoding = ReadUInt16(_data, record + 2);
            uint offsetValue = ReadUInt32(_data, record + 4);
            if (offsetValue > (uint)(_cmapLength - 2)) continue;
            int offset = checked((int)offsetValue);
            var absolute = cmapOffset + offset;
            if (absolute < cmapOffset || absolute > cmapEnd - 2) continue;
            var format = ReadUInt16(_data, absolute);
            if (!OfficeOpenTypeCmap.IsUnicodeEncoding(platform, encoding)) continue;
            if (format == 4 && !_validFormat4Subtables.Contains(absolute)) continue;
            if (format == 12 && !_validFormat12Subtables.Contains(absolute)) continue;
            var score = OfficeOpenTypeCmap.ScoreSubtable(
                format,
                platform,
                encoding,
                preferFormat12: scalar > 0xFFFF);
            if ((format == 4 || format == 12) && score > bestScore) {
                best = absolute;
                bestScore = score;
            }
        }

        if (best == 0) return 0;
        ushort glyph = MapCmapSubtable(best, cmapEnd, scalar);
        return glyph != 0 ? glyph : MapFallbackCmapSubtable(best, cmapEnd, scalar, subtableCount);
    }

    private ushort MapCmapSubtable(int table, int cmapEnd, int scalar) =>
        ReadUInt16(_data, table) == 12
            ? MapFormat12(table, cmapEnd, scalar)
            : MapFormat4(table, cmapEnd, scalar);

    private ushort MapFallbackCmapSubtable(int selectedTable, int cmapEnd, int scalar, int subtableCount) {
        int bestScore = 0;
        ushort bestGlyph = 0;
        for (int index = 0; index < subtableCount; index++) {
            int record = _cmap + 4 + index * 8;
            int platform = ReadUInt16(_data, record);
            int encoding = ReadUInt16(_data, record + 2);
            uint offsetValue = ReadUInt32(_data, record + 4);
            if (offsetValue > (uint)(_cmapLength - 2)) continue;
            int table = _cmap + checked((int)offsetValue);
            if (table == selectedTable || table < _cmap || table > cmapEnd - 2 ||
                !OfficeOpenTypeCmap.IsUnicodeEncoding(platform, encoding)) continue;
            int format = ReadUInt16(_data, table);
            if (format == 4 && (scalar > 0xFFFF || !_validFormat4Subtables.Contains(table))) continue;
            if (format == 12 && !_validFormat12Subtables.Contains(table)) continue;
            if (format != 4 && format != 12) continue;
            int score = OfficeOpenTypeCmap.ScoreSubtable(
                format,
                platform,
                encoding,
                preferFormat12: scalar > 0xFFFF);
            if (score <= bestScore) continue;
            ushort glyph = MapCmapSubtable(table, cmapEnd, scalar);
            if (glyph == 0) continue;
            bestGlyph = glyph;
            bestScore = score;
        }
        return bestGlyph;
    }

    private ushort MapFormat4(int table, int cmapEnd, int scalar) {
        if (!_validFormat4Subtables.Contains(table)) return 0;
        if (scalar > char.MaxValue) return 0;
        if (table < _cmap || table > cmapEnd - 16) return 0;
        var length = ReadUInt16(_data, table + 2);
        if (length < 16 || table > cmapEnd - length) return 0;
        var code = scalar;
        var segCount = ReadUInt16(_data, table + 6) / 2;
        if (segCount == 0 || segCount > OfficeOpenTypeCmap.MaximumSubtables * 16) return 0;
        var endCodes = table + 14;
        var startCodes = endCodes + segCount * 2 + 2;
        var idDeltas = startCodes + segCount * 2;
        var idRangeOffsets = idDeltas + segCount * 2;
        if (idRangeOffsets < table || idRangeOffsets + segCount * 2 > table + length) return 0;

        int low = 0;
        int high = segCount - 1;
        while (low <= high) {
            int i = low + (high - low) / 2;
            int end = ReadUInt16(_data, endCodes + i * 2);
            int start = ReadUInt16(_data, startCodes + i * 2);
            if (code < start) {
                high = i - 1;
                continue;
            }
            if (code > end) {
                low = i + 1;
                continue;
            }
            int delta = ReadInt16(_data, idDeltas + i * 2);
            int rangeOffset = ReadUInt16(_data, idRangeOffsets + i * 2);
            if (rangeOffset == 0) return ValidateMappedGlyph((ushort)((code + delta) & 0xffff));
            int glyphOffset = idRangeOffsets + i * 2 + rangeOffset + (code - start) * 2;
            if (glyphOffset < table || glyphOffset > table + length - 2) return 0;
            ushort glyph = ReadUInt16(_data, glyphOffset);
            return glyph == 0 ? (ushort)0 : ValidateMappedGlyph((ushort)((glyph + delta) & 0xffff));
        }

        return 0;
    }

    private ushort MapFormat12(int table, int cmapEnd, int scalar) {
        if (!_validFormat12Subtables.Contains(table)) return 0;
        if (table < _cmap || table > cmapEnd - 16) return 0;
        var length = ReadUInt32(_data, table + 4);
        if (length < 16 || length > int.MaxValue || table > cmapEnd - (int)length) return 0;
        uint code = (uint)scalar;
        var groups = ReadUInt32(_data, table + 12);
        if (groups > OfficeOpenTypeCmap.MaximumFormat12Groups || groups > (length - 16U) / 12U) return 0;
        var groupOffset = table + 16;
        uint low = 0;
        uint high = groups;
        while (low < high) {
            uint middle = low + (high - low) / 2;
            int current = checked(groupOffset + (int)middle * 12);
            uint start = ReadUInt32(_data, current);
            uint end = ReadUInt32(_data, current + 4);
            if (code < start) {
                high = middle;
                continue;
            }
            if (code > end) {
                low = middle + 1;
                continue;
            }
            ulong glyph = (ulong)ReadUInt32(_data, current + 8) + code - start;
            return glyph < _numGlyphs ? checked((ushort)glyph) : (ushort)0;
        }

        return 0;
    }

    private ushort ValidateMappedGlyph(ushort glyph) => glyph > 0 && glyph < _numGlyphs ? glyph : (ushort)0;

    private int BaseAdvanceWidth(ushort glyph) => glyph < _numHMetrics
            ? ReadUInt16(_data, _hmtx + glyph * 4)
            : ReadUInt16(_data, _hmtx + (_numHMetrics - 1) * 4);

    private int AdvanceWidth(ushort glyph) => AdvanceWidth(glyph, null, CancellationToken.None);

    private int Kerning(ushort left, ushort right, int leftScalar, int rightScalar) =>
        _kerning.Adjustment(left, right, leftScalar, rightScalar);

    private int AdvanceWidth(
        ushort glyph,
        OfficeTrueTypeVariations.WorkBudget? workBudget,
        CancellationToken cancellationToken) {
        int advance = BaseAdvanceWidth(glyph);
        return checked(advance + VariationAdvanceWidthDelta(glyph, advance, workBudget, cancellationToken));
    }

    private List<List<OfficePoint>> ReadGlyphContours(
        ushort glyph,
        FontTransform transform,
        int depth,
        OfficeTrueTypeVariations.WorkBudget? variationWorkBudget,
        int maximumPointCount,
        ref int expandedPointCount,
        CancellationToken cancellationToken,
        List<OfficePoint>? attachmentPoints) {
        cancellationToken.ThrowIfCancellationRequested();
        var contours = new List<List<OfficePoint>>();
        if (glyph == 0 || glyph >= _numGlyphs || depth > 8) return contours;
        var glyphStart = GlyphOffset(glyph);
        var glyphEnd = GlyphOffset((ushort)(glyph + 1));
        if (glyphStart == glyphEnd) return contours;
        var offset = _glyf + glyphStart;
        if (offset + 10 > _data.Length) return contours;
        var contourCount = ReadInt16(_data, offset);
        if (contourCount < 0) {
            ReadCompositeGlyphContours(
                glyph,
                offset,
                transform,
                depth,
                contours,
                variationWorkBudget,
                maximumPointCount,
                ref expandedPointCount,
                cancellationToken,
                attachmentPoints);
            return contours;
        }

        if (contourCount <= 0) return contours;

        var endPts = new ushort[contourCount];
        for (var i = 0; i < contourCount; i++) endPts[i] = ReadUInt16(_data, offset + 10 + i * 2);
        var rawPointCount = endPts[contourCount - 1] + 1;
        if (rawPointCount > maximumPointCount - expandedPointCount) {
            throw new InvalidOperationException("Font outline expansion exceeded the configured point budget.");
        }
        var instructionLengthOffset = offset + 10 + contourCount * 2;
        var instructionLength = ReadUInt16(_data, instructionLengthOffset);
        var p = instructionLengthOffset + 2 + instructionLength;
        var flags = new byte[rawPointCount];
        for (var i = 0; i < rawPointCount; i++) {
            var flag = _data[p++];
            flags[i] = flag;
            if ((flag & 8) == 0) continue;
            var repeat = _data[p++];
            for (var r = 0; r < repeat && i + 1 < rawPointCount; r++) flags[++i] = flag;
        }

        var xs = new short[rawPointCount];
        DecodeCoordinates(_data, flags, xs, ref p, true);
        var ys = new short[rawPointCount];
        DecodeCoordinates(_data, flags, ys, ref p, false);
        double[] variedXs;
        double[] variedYs;
        if (_variations != null) {
            variedXs = new double[rawPointCount];
            variedYs = new double[rawPointCount];
            for (int index = 0; index < rawPointCount; index++) {
                variedXs[index] = xs[index];
                variedYs[index] = ys[index];
            }
            _variations.ApplySimpleGlyph(
                glyph,
                variedXs,
                variedYs,
                endPts,
                variationWorkBudget ?? _variations.CreateWorkBudget(),
                cancellationToken);
        } else {
            variedXs = new double[rawPointCount];
            variedYs = new double[rawPointCount];
            for (int index = 0; index < rawPointCount; index++) {
                variedXs[index] = xs[index];
                variedYs[index] = ys[index];
            }
        }

        if (attachmentPoints != null) {
            for (var index = 0; index < rawPointCount; index++) {
                attachmentPoints.Add(transform.Apply(variedXs[index], variedYs[index]));
            }
        }

        var start = 0;
        for (var c = 0; c < contourCount; c++) {
            var end = endPts[c];
            var points = new List<GlyphPoint>();
            for (var i = start; i <= end; i++) {
                var point = transform.Apply(variedXs[i], variedYs[i]);
                points.Add(new GlyphPoint(point.X, point.Y, (flags[i] & 1) != 0));
            }

            AddFlattenedContour(points, contours, ref expandedPointCount, maximumPointCount);
            start = end + 1;
        }

        return contours;
    }

    private void ReadCompositeGlyphContours(
        ushort glyph,
        int glyphOffset,
        FontTransform transform,
        int depth,
        List<List<OfficePoint>> contours,
        OfficeTrueTypeVariations.WorkBudget? variationWorkBudget,
        int maximumPointCount,
        ref int expandedPointCount,
        CancellationToken cancellationToken,
        List<OfficePoint>? attachmentPoints) {
        const ushort argWords = 1;
        const ushort argsAreXy = 2;
        const ushort haveScale = 8;
        const ushort moreComponents = 32;
        const ushort haveXyScale = 64;
        const ushort haveTwoByTwo = 128;

        var p = glyphOffset + 10;
        ushort flags;
        var components = new List<GlyphComponent>();
        do {
            cancellationToken.ThrowIfCancellationRequested();
            if (p + 4 > _data.Length) return;
            flags = ReadUInt16(_data, p);
            if (OfficeOpenTypeCompositeGlyph.HasConflictingTransformFlags(flags) ||
                OfficeOpenTypeCompositeGlyph.HasConflictingOffsetFlags(flags)) return;
            var componentGlyph = ReadUInt16(_data, p + 2);
            p += 4;
            int arg1;
            int arg2;
            if ((flags & argWords) != 0) {
                if (p + 4 > _data.Length) return;
                if ((flags & argsAreXy) != 0) {
                    arg1 = ReadInt16(_data, p);
                    arg2 = ReadInt16(_data, p + 2);
                } else {
                    arg1 = ReadUInt16(_data, p);
                    arg2 = ReadUInt16(_data, p + 2);
                }
                p += 4;
            } else {
                if (p + 2 > _data.Length) return;
                if ((flags & argsAreXy) != 0) {
                    arg1 = (sbyte)_data[p];
                    arg2 = (sbyte)_data[p + 1];
                } else {
                    arg1 = _data[p];
                    arg2 = _data[p + 1];
                }
                p += 2;
            }

            var dx = (flags & argsAreXy) != 0 ? arg1 : 0;
            var dy = (flags & argsAreXy) != 0 ? arg2 : 0;
            int parentPointIndex = (flags & argsAreXy) == 0 ? arg1 : -1;
            int componentPointIndex = (flags & argsAreXy) == 0 ? arg2 : -1;
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

            components.Add(new GlyphComponent(
                componentGlyph,
                a,
                b,
                c,
                d,
                dx,
                dy,
                flags,
                (flags & argsAreXy) != 0,
                parentPointIndex,
                componentPointIndex));
        } while ((flags & moreComponents) != 0);

        if (_variations != null && components.Count > 0) {
            var xs = new double[components.Count];
            var ys = new double[components.Count];
            for (int index = 0; index < components.Count; index++) {
                xs[index] = components[index].Dx;
                ys[index] = components[index].Dy;
            }
            _variations.ApplyCompositeGlyph(
                glyph,
                xs,
                ys,
                variationWorkBudget ?? _variations.CreateWorkBudget(),
                cancellationToken);
            for (int index = 0; index < components.Count; index++) {
                components[index] = components[index].WithOffset(xs[index], ys[index]);
            }
        }

        var compositeContours = new List<List<OfficePoint>>();
        var compositePoints = new List<OfficePoint>();
        foreach (GlyphComponent component in components) {
            var componentPoints = new List<OfficePoint>();
            List<List<OfficePoint>> componentContours = ReadGlyphContours(
                component.Glyph,
                transform.Compose(component.A, component.B, component.C, component.D, 0, 0),
                depth + 1,
                variationWorkBudget,
                maximumPointCount,
                ref expandedPointCount,
                cancellationToken,
                componentPoints);

            OfficePoint translation;
            if (component.UsesXyOffsets) {
                OfficePoint offset = OfficeOpenTypeCompositeGlyph.ResolveXyOffset(
                    component.Flags,
                    component.A,
                    component.B,
                    component.C,
                    component.D,
                    component.Dx,
                    component.Dy);
                translation = transform.ApplyVector(offset.X, offset.Y);
            } else {
                if (!OfficeOpenTypeCompositeGlyph.TryResolvePointAttachment(
                    compositePoints,
                    component.ParentPointIndex,
                    componentPoints,
                    component.ComponentPointIndex,
                    out translation)) {
                    return;
                }

                OfficePoint variationTranslation = transform.ApplyVector(component.Dx, component.Dy);
                translation = OfficeOpenTypeCompositeGlyph.ApplyComponentVariation(translation, variationTranslation);
            }

            TranslateContours(componentContours, translation);
            TranslatePoints(componentPoints, translation);
            compositeContours.AddRange(componentContours);
            compositePoints.AddRange(componentPoints);
        }

        contours.AddRange(compositeContours);
        attachmentPoints?.AddRange(compositePoints);
    }

    private static void TranslateContours(List<List<OfficePoint>> contours, OfficePoint translation) {
        foreach (List<OfficePoint> contour in contours) TranslatePoints(contour, translation);
    }

    private static void TranslatePoints(List<OfficePoint> points, OfficePoint translation) {
        for (int index = 0; index < points.Count; index++) {
            OfficePoint point = points[index];
            points[index] = new OfficePoint(point.X + translation.X, point.Y + translation.Y);
        }
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

    private static void AddFlattenedContour(
        List<GlyphPoint> source,
        List<List<OfficePoint>> contours,
        ref int expandedPointCount,
        int maximumPointCount) {
        if (source.Count == 0) return;
        var result = new List<OfficePoint>();
        var last = source[source.Count - 1];
        var first = source[0];
        var current = first.OnCurve ? first : last.OnCurve ? last : Mid(last, first);
        AddBoundedOutlinePoint(result, current.Point, ref expandedPointCount, maximumPointCount);
        var index = first.OnCurve ? 1 : 0;

        while (index < source.Count) {
            var point = source[index % source.Count];
            if (point.OnCurve) {
                AddBoundedOutlinePoint(result, point.Point, ref expandedPointCount, maximumPointCount);
                current = point;
                index++;
                continue;
            }

            var next = source[(index + 1) % source.Count];
            var end = next.OnCurve ? next : Mid(point, next);
            FlattenQuadratic(current, point, end, result, ref expandedPointCount, maximumPointCount);
            current = end;
            index += next.OnCurve ? 2 : 1;
        }

        if (result.Count >= 3) contours.Add(result);
    }

    private static void FlattenQuadratic(
        GlyphPoint start,
        GlyphPoint control,
        GlyphPoint end,
        List<OfficePoint> output,
        ref int expandedPointCount,
        int maximumPointCount) {
        var chord = Math.Sqrt((end.X - start.X) * (end.X - start.X) + (end.Y - start.Y) * (end.Y - start.Y));
        var bend = Math.Sqrt((start.X - 2 * control.X + end.X) * (start.X - 2 * control.X + end.X) + (start.Y - 2 * control.Y + end.Y) * (start.Y - 2 * control.Y + end.Y));
        var steps = Math.Max(6, Math.Min(18, (int)Math.Ceiling((chord + bend * 2.0) / 120.0)));
        for (var i = 1; i <= steps; i++) {
            var t = i / (double)steps;
            var mt = 1 - t;
            AddBoundedOutlinePoint(
                output,
                new OfficePoint(mt * mt * start.X + 2 * mt * t * control.X + t * t * end.X, mt * mt * start.Y + 2 * mt * t * control.Y + t * t * end.Y),
                ref expandedPointCount,
                maximumPointCount);
        }
    }

    private static void AddBoundedOutlinePoint(
        ICollection<OfficePoint> target,
        OfficePoint point,
        ref int expandedPointCount,
        int maximumPointCount) {
        if (expandedPointCount >= maximumPointCount) {
            throw new InvalidOperationException("Font outline expansion exceeded the configured point budget.");
        }
        expandedPointCount++;
        target.Add(point);
    }

    private static GlyphPoint Mid(GlyphPoint left, GlyphPoint right) => new((left.X + right.X) / 2.0, (left.Y + right.Y) / 2.0, true);

    private static FontFamilyResolution ResolveFontFamily(string? fontFamily) {
        if (string.IsNullOrEmpty(fontFamily)) {
            OfficeTrueTypeFont? defaultFont = TryLoadDefault(out string? defaultPath);
            return new FontFamilyResolution(defaultFont, defaultPath);
        }

        foreach (string family in ExpandFontFamilyFallbacks(fontFamily!)) {
            foreach (string path in CandidateFamilyPaths(family)) {
                OfficeTrueTypeFont? font = TryLoad(path, null, family);
                if (font == null) {
                    font = TryLoad(path);
                }

                if (font != null && font.HasGlyphs("OfficeIMO 0123456789")) {
                    return new FontFamilyResolution(font, path);
                }
            }
        }

        return new FontFamilyResolution(null, null);
    }

    private static string NormalizeFontFamilyCacheKey(string? fontFamily) {
        if (string.IsNullOrEmpty(fontFamily)) {
            return "__default";
        }

        var builder = new StringBuilder();
        foreach (string family in ExpandFontFamilyFallbacks(fontFamily!)) {
            if (builder.Length > 0) {
                builder.Append('|');
            }

            builder.Append(family);
        }

        return builder.Length == 0 ? "__empty" : builder.ToString();
    }

    private static IEnumerable<string> ExpandFontFamilyFallbacks(string fontFamily) {
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        int emitted = 0;
        foreach (string family in OfficeFontFamilyParser.Parse(fontFamily)) {
            foreach (string expanded in ExpandGenericFontFamily(family)) {
                if (seen.Add(expanded)) {
                    yield return expanded;
                    emitted++;
                    if (emitted >= OfficeFontFamilyParser.DefaultMaximumCandidates) yield break;
                }
            }
        }
    }

    private static IEnumerable<string> ExpandGenericFontFamily(string family) {
        string key = NormalizeFontFamilyKey(family);
        if (key == "sansserif" || key == "sans") {
            yield return "Aptos";
            yield return "Calibri";
            yield return "Arial";
            yield return "Segoe UI";
            yield return "Liberation Sans";
            yield return "DejaVu Sans";
            yield break;
        }

        if (key == "serif") {
            yield return "Times New Roman";
            yield return "Georgia";
            yield return "Liberation Serif";
            yield return "DejaVu Serif";
            yield break;
        }

        if (key == "monospace" || key == "mono") {
            yield return "Consolas";
            yield return "Courier New";
            yield return "Liberation Mono";
            yield return "DejaVu Sans Mono";
            yield break;
        }

        yield return family;
    }

    private static IEnumerable<string> CandidateFamilyPaths(string family) {
        string key = NormalizeFontFamilyKey(family);
        foreach (string path in CandidateKnownFamilyPaths(key)) {
            yield return path;
        }

        foreach (string path in CandidateFontDirectoryPaths(key)) {
            yield return path;
        }
    }

    private static IEnumerable<string> CandidateKnownFamilyPaths(string key) {
        string windows = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
        if (!string.IsNullOrEmpty(windows)) {
            string fonts = Path.Combine(windows, "Fonts");
            if (key == "aptos") {
                yield return Path.Combine(fonts, "aptos.ttf");
                yield return Path.Combine(fonts, "aptosdisplay.ttf");
            } else if (key == "aptosdisplay") {
                yield return Path.Combine(fonts, "aptosdisplay.ttf");
                yield return Path.Combine(fonts, "aptos.ttf");
            } else if (key == "aptosnarrow") {
                yield return Path.Combine(fonts, "aptosnarrow.ttf");
                yield return Path.Combine(fonts, "aptos.ttf");
            } else if (key == "calibri") {
                yield return Path.Combine(fonts, "calibri.ttf");
            } else if (key == "arial") {
                yield return Path.Combine(fonts, "arial.ttf");
            } else if (key == "timesnewroman") {
                yield return Path.Combine(fonts, "times.ttf");
            } else if (key == "couriernew") {
                yield return Path.Combine(fonts, "cour.ttf");
            } else if (key == "segoeui") {
                yield return Path.Combine(fonts, "segoeui.ttf");
            } else if (key == "consolas") {
                yield return Path.Combine(fonts, "consola.ttf");
            } else if (key == "tahoma") {
                yield return Path.Combine(fonts, "tahoma.ttf");
            } else if (key == "verdana") {
                yield return Path.Combine(fonts, "verdana.ttf");
            } else if (key == "georgia") {
                yield return Path.Combine(fonts, "georgia.ttf");
            } else if (key == "trebuchetms") {
                yield return Path.Combine(fonts, "trebuc.ttf");
            }
        }

        if (key == "arial") {
            yield return "/Library/Fonts/Arial.ttf";
            yield return "/System/Library/Fonts/Supplemental/Arial.ttf";
        } else if (key == "helvetica" || key == "helveticaneue") {
            yield return "/System/Library/Fonts/Helvetica.ttc";
            yield return "/System/Library/Fonts/HelveticaNeue.ttc";
        } else if (key == "timesnewroman") {
            yield return "/Library/Fonts/Times New Roman.ttf";
            yield return "/System/Library/Fonts/Supplemental/Times New Roman.ttf";
        } else if (key == "couriernew") {
            yield return "/Library/Fonts/Courier New.ttf";
            yield return "/System/Library/Fonts/Supplemental/Courier New.ttf";
        } else if (key == "dejavusans") {
            yield return "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf";
            yield return "/usr/share/fonts/dejavu/DejaVuSans.ttf";
        } else if (key == "dejavuserif") {
            yield return "/usr/share/fonts/truetype/dejavu/DejaVuSerif.ttf";
            yield return "/usr/share/fonts/dejavu/DejaVuSerif.ttf";
        } else if (key == "dejavusansmono") {
            yield return "/usr/share/fonts/truetype/dejavu/DejaVuSansMono.ttf";
            yield return "/usr/share/fonts/dejavu/DejaVuSansMono.ttf";
        } else if (key == "liberationsans") {
            yield return "/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf";
            yield return "/usr/share/fonts/liberation/LiberationSans-Regular.ttf";
        } else if (key == "liberationserif") {
            yield return "/usr/share/fonts/truetype/liberation2/LiberationSerif-Regular.ttf";
            yield return "/usr/share/fonts/liberation/LiberationSerif-Regular.ttf";
        } else if (key == "liberationmono") {
            yield return "/usr/share/fonts/truetype/liberation2/LiberationMono-Regular.ttf";
            yield return "/usr/share/fonts/liberation/LiberationMono-Regular.ttf";
        }
    }

    private static IEnumerable<string> CandidateFontDirectoryPaths(string key) {
        foreach (string directory in CandidateFontDirectories()) {
            if (!Directory.Exists(directory)) {
                continue;
            }

            foreach (string path in SafeEnumerateFontFiles(directory)) {
                string fileKey = NormalizeFontFamilyKey(Path.GetFileNameWithoutExtension(path));
                if (fileKey == key || fileKey.StartsWith(key, StringComparison.OrdinalIgnoreCase) || fileKey.Contains(key)) {
                    yield return path;
                }
            }
        }
    }

    private static IEnumerable<string> CandidateFontDirectories() {
        string windows = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
        if (!string.IsNullOrEmpty(windows)) {
            yield return Path.Combine(windows, "Fonts");
        }

        yield return "/System/Library/Fonts";
        yield return "/System/Library/Fonts/Supplemental";
        yield return "/Library/Fonts";
        yield return "/usr/share/fonts/truetype/dejavu";
        yield return "/usr/share/fonts/truetype/liberation2";
        yield return "/usr/share/fonts/dejavu";
        yield return "/usr/share/fonts/liberation";
    }

    private static IEnumerable<string> SafeEnumerateFontFiles(string directory) {
        string[] files;
        try {
            files = Directory.GetFiles(directory, "*.*", SearchOption.TopDirectoryOnly);
        } catch (IOException) {
            yield break;
        } catch (UnauthorizedAccessException) {
            yield break;
        } catch (ArgumentException) {
            yield break;
        } catch (NotSupportedException) {
            yield break;
        }

        foreach (string file in files) {
            string extension = Path.GetExtension(file);
            if (string.Equals(extension, ".ttf", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(extension, ".otf", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(extension, ".ttc", StringComparison.OrdinalIgnoreCase)) {
                yield return file;
            }
        }
    }

    private static string NormalizeFontFamilyKey(string family) {
        var builder = new StringBuilder(family.Length);
        for (int i = 0; i < family.Length; i++) {
            char value = family[i];
            if (char.IsLetterOrDigit(value)) {
                builder.Append(char.ToLowerInvariant(value));
            }
        }

        return builder.ToString();
    }

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

    private static bool TryCheckedOffset(byte[] data, uint offset, out int checkedOffset) {
        if (offset > int.MaxValue || offset >= data.Length) {
            checkedOffset = 0;
            return false;
        }

        checkedOffset = (int)offset;
        return true;
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

    private readonly struct FontFamilyResolution {
        public FontFamilyResolution(OfficeTrueTypeFont? font, string? path) {
            Font = font;
            Path = path;
        }

        public OfficeTrueTypeFont? Font { get; }

        public string? Path { get; }
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

        public OfficePoint ApplyVector(double x, double y) => new(Xx * x + Xy * y, Yx * x + Yy * y);

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

    private readonly struct GlyphComponent {
        internal GlyphComponent(
            ushort glyph,
            double a,
            double b,
            double c,
            double d,
            double dx,
            double dy,
            ushort flags,
            bool usesXyOffsets,
            int parentPointIndex,
            int componentPointIndex) {
            Glyph = glyph;
            A = a;
            B = b;
            C = c;
            D = d;
            Dx = dx;
            Dy = dy;
            Flags = flags;
            UsesXyOffsets = usesXyOffsets;
            ParentPointIndex = parentPointIndex;
            ComponentPointIndex = componentPointIndex;
        }

        internal ushort Glyph { get; }
        internal double A { get; }
        internal double B { get; }
        internal double C { get; }
        internal double D { get; }
        internal double Dx { get; }
        internal double Dy { get; }
        internal ushort Flags { get; }
        internal bool UsesXyOffsets { get; }
        internal int ParentPointIndex { get; }
        internal int ComponentPointIndex { get; }
        internal GlyphComponent WithOffset(double dx, double dy) =>
            new(Glyph, A, B, C, D, dx, dy, Flags, UsesXyOffsets, ParentPointIndex, ComponentPointIndex);
    }
}
