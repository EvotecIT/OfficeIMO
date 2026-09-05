using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;

namespace OfficeIMO.Drawing;

internal interface IOfficeColorFontProgram {
    bool HasColorGlyph(int glyphId);
    bool TryGetColorLayers(int glyphId, string? palette, OfficeColor foreground, out IReadOnlyList<OfficeColorGlyphLayer> layers);
}

internal readonly struct OfficeColorGlyphLayer {
    internal OfficeColorGlyphLayer(int glyphId, OfficeColor color) {
        GlyphId = glyphId;
        Color = color;
    }

    internal int GlyphId { get; }
    internal OfficeColor Color { get; }
}

/// <summary>Bounded COLR v0 and CPAL v0/v1 reader shared by raster and vector renderers.</summary>
internal sealed class OfficeOpenTypeColorGlyphs {
    private const int MaximumBaseGlyphs = 65535;
    private const int MaximumLayers = 262144;
    private const int MaximumPalettes = 4096;
    private const int MaximumPaletteEntries = 4096;

    private readonly IReadOnlyDictionary<int, LayerRecord[]> _layers;
    private readonly OfficeColor[][] _palettes;
    private readonly uint[] _paletteTypes;

    private OfficeOpenTypeColorGlyphs(
        IReadOnlyDictionary<int, LayerRecord[]> layers,
        OfficeColor[][] palettes,
        uint[] paletteTypes) {
        _layers = layers;
        _palettes = palettes;
        _paletteTypes = paletteTypes;
    }

    internal static OfficeOpenTypeColorGlyphs? TryParse(OfficeOpenTypeReader reader) {
        if (reader == null) throw new ArgumentNullException(nameof(reader));
        try {
            if (!reader.TryGetTable("COLR", out int colr, out int colrLength) || colrLength < 14 ||
                !reader.TryGetTable("CPAL", out int cpal, out int cpalLength) || cpalLength < 12) {
                return null;
            }

            int colrEnd = checked(colr + colrLength);
            int cpalEnd = checked(cpal + cpalLength);
            ushort colrVersion = reader.ReadUInt16(colr);
            if (colrVersion != 0) return null;

            int baseGlyphCount = reader.ReadUInt16(colr + 2);
            int baseRecords = Relative32(reader, colr, colr + 4, colr, colrEnd, 6);
            int layerRecords = Relative32(reader, colr, colr + 8, colr, colrEnd, 4);
            int layerCount = reader.ReadUInt16(colr + 12);
            if (baseGlyphCount > MaximumBaseGlyphs || layerCount > MaximumLayers) {
                throw new InvalidDataException("The COLR table exceeds the managed color-glyph limits.");
            }
            EnsureRange(baseRecords, checked(baseGlyphCount * 6), colr, colrEnd);
            EnsureRange(layerRecords, checked(layerCount * 4), colr, colrEnd);

            var layerMap = new Dictionary<int, LayerRecord[]>(baseGlyphCount);
            int previousBaseGlyph = -1;
            for (int index = 0; index < baseGlyphCount; index++) {
                int record = baseRecords + index * 6;
                int glyphId = reader.ReadUInt16(record);
                int firstLayer = reader.ReadUInt16(record + 2);
                int glyphLayerCount = reader.ReadUInt16(record + 4);
                if (glyphId <= previousBaseGlyph || glyphId >= reader.GlyphCount ||
                    glyphLayerCount <= 0 || firstLayer > layerCount - glyphLayerCount) {
                    throw new InvalidDataException("The COLR base-glyph records are invalid.");
                }
                previousBaseGlyph = glyphId;
                var glyphLayers = new LayerRecord[glyphLayerCount];
                for (int layerIndex = 0; layerIndex < glyphLayerCount; layerIndex++) {
                    int layer = layerRecords + (firstLayer + layerIndex) * 4;
                    int layerGlyph = reader.ReadUInt16(layer);
                    int paletteEntry = reader.ReadUInt16(layer + 2);
                    if (layerGlyph <= 0 || layerGlyph >= reader.GlyphCount) {
                        throw new InvalidDataException("A COLR layer references an invalid glyph.");
                    }
                    glyphLayers[layerIndex] = new LayerRecord(layerGlyph, paletteEntry);
                }
                layerMap.Add(glyphId, glyphLayers);
            }

            ushort cpalVersion = reader.ReadUInt16(cpal);
            if (cpalVersion > 1) return null;
            int entriesPerPalette = reader.ReadUInt16(cpal + 2);
            int paletteCount = reader.ReadUInt16(cpal + 4);
            int colorRecordCount = reader.ReadUInt16(cpal + 6);
            int colorRecords = Relative32(reader, cpal, cpal + 8, cpal, cpalEnd, 4);
            if (entriesPerPalette <= 0 || entriesPerPalette > MaximumPaletteEntries ||
                paletteCount <= 0 || paletteCount > MaximumPalettes ||
                colorRecordCount <= 0 || colorRecordCount > MaximumLayers) {
                throw new InvalidDataException("The CPAL table exceeds the managed palette limits.");
            }
            int paletteIndexes = cpal + 12;
            EnsureRange(paletteIndexes, checked(paletteCount * 2), cpal, cpalEnd);
            EnsureRange(colorRecords, checked(colorRecordCount * 4), cpal, cpalEnd);

            var palettes = new OfficeColor[paletteCount][];
            for (int paletteIndex = 0; paletteIndex < paletteCount; paletteIndex++) {
                int firstColor = reader.ReadUInt16(paletteIndexes + paletteIndex * 2);
                if (firstColor > colorRecordCount - entriesPerPalette) {
                    throw new InvalidDataException("A CPAL palette range is invalid.");
                }
                var colors = new OfficeColor[entriesPerPalette];
                for (int entry = 0; entry < entriesPerPalette; entry++) {
                    int color = colorRecords + (firstColor + entry) * 4;
                    colors[entry] = OfficeColor.FromRgba(
                        reader.Data[color + 2],
                        reader.Data[color + 1],
                        reader.Data[color],
                        reader.Data[color + 3]);
                }
                palettes[paletteIndex] = colors;
            }

            var paletteTypes = new uint[paletteCount];
            if (cpalVersion == 1) {
                int versionOneHeader = checked(paletteIndexes + paletteCount * 2);
                EnsureRange(versionOneHeader, 12, cpal, cpalEnd);
                uint typesRelative = reader.ReadUInt32(versionOneHeader);
                if (typesRelative != 0) {
                    if (typesRelative > int.MaxValue) throw new InvalidDataException("The CPAL palette-type offset is invalid.");
                    int types = checked(cpal + (int)typesRelative);
                    EnsureRange(types, checked(paletteCount * 4), cpal, cpalEnd);
                    for (int index = 0; index < paletteCount; index++) paletteTypes[index] = reader.ReadUInt32(types + index * 4);
                }
            }

            return new OfficeOpenTypeColorGlyphs(
                new ReadOnlyDictionary<int, LayerRecord[]>(layerMap),
                palettes,
                paletteTypes);
        } catch (Exception exception) when (exception is InvalidDataException
                                            || exception is OverflowException
                                            || exception is ArgumentOutOfRangeException
                                            || exception is IndexOutOfRangeException) {
            return null;
        }
    }

    internal bool HasColorGlyph(int glyphId) => _layers.ContainsKey(glyphId);

    internal bool TryGetLayers(
        int glyphId,
        string? palette,
        OfficeColor foreground,
        out IReadOnlyList<OfficeColorGlyphLayer> layers) {
        if (!_layers.TryGetValue(glyphId, out LayerRecord[]? records)) {
            layers = Array.Empty<OfficeColorGlyphLayer>();
            return false;
        }

        int paletteIndex = ResolvePaletteIndex(palette);
        OfficeColor[] colors = _palettes[paletteIndex];
        var resolved = new OfficeColorGlyphLayer[records.Length];
        for (int index = 0; index < records.Length; index++) {
            LayerRecord record = records[index];
            OfficeColor color = record.PaletteEntry == ushort.MaxValue
                ? foreground
                : record.PaletteEntry < colors.Length
                    ? colors[record.PaletteEntry]
                    : foreground;
            resolved[index] = new OfficeColorGlyphLayer(record.GlyphId, color);
        }
        layers = Array.AsReadOnly(resolved);
        return true;
    }

    private int ResolvePaletteIndex(string? value) {
        string palette = string.IsNullOrWhiteSpace(value) ? "normal" : value!.Trim().ToLowerInvariant();
        const uint UsableWithLightBackground = 0x00000001;
        const uint UsableWithDarkBackground = 0x00000002;
        uint requestedType = palette == "light"
            ? UsableWithLightBackground
            : palette == "dark"
                ? UsableWithDarkBackground
                : 0;
        if (requestedType != 0) {
            for (int index = 0; index < _paletteTypes.Length; index++) {
                if ((_paletteTypes[index] & requestedType) != 0) return index;
            }
        }
        return 0;
    }

    private static int Relative32(
        OfficeOpenTypeReader reader,
        int origin,
        int offsetLocation,
        int start,
        int end,
        int minimumLength) {
        uint relative = reader.ReadUInt32(offsetLocation);
        if (relative > int.MaxValue) throw new InvalidDataException("An OpenType color-table offset is invalid.");
        int offset = checked(origin + (int)relative);
        EnsureRange(offset, minimumLength, start, end);
        return offset;
    }

    private static void EnsureRange(int offset, int length, int start, int end) {
        if (offset < start || length < 0 || offset > end - length) {
            throw new InvalidDataException("An OpenType color table is truncated.");
        }
    }

    private readonly struct LayerRecord {
        internal LayerRecord(int glyphId, int paletteEntry) {
            GlyphId = glyphId;
            PaletteEntry = paletteEntry;
        }

        internal int GlyphId { get; }
        internal int PaletteEntry { get; }
    }
}
