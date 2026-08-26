using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeTrueTypeFont {
    private readonly object _variationMetricSync = new();
    private Dictionary<ushort, int>? _gvarAdvanceDeltas;

    private int VariationAdvanceWidthDelta(
        ushort glyph,
        int baseAdvanceWidth,
        OfficeTrueTypeVariations.WorkBudget? workBudget,
        CancellationToken cancellationToken) {
        if (_variations == null) return 0;
        if (_variations.HasHvar) return _variations.AdvanceWidthDelta(glyph);
        OfficeTrueTypeVariations.WorkBudget effectiveWorkBudget = workBudget ?? _variations.CreateWorkBudget();
        lock (_variationMetricSync) {
            return ResolveSelectedAdvanceWidth(
                glyph,
                effectiveWorkBudget,
                cancellationToken,
                new HashSet<ushort>(),
                0) - baseAdvanceWidth;
        }
    }

    private int ResolveSelectedAdvanceWidth(
        ushort glyph,
        OfficeTrueTypeVariations.WorkBudget workBudget,
        CancellationToken cancellationToken,
        ISet<ushort> activeGlyphs,
        int depth) {
        int baseAdvanceWidth = BaseAdvanceWidth(glyph);
        if (_gvarAdvanceDeltas != null && _gvarAdvanceDeltas.TryGetValue(glyph, out int cachedDelta)) {
            return checked(baseAdvanceWidth + cachedDelta);
        }
        if (depth >= 16 || !activeGlyphs.Add(glyph)) {
            throw new InvalidDataException("A variable composite glyph has a cyclic USE_MY_METRICS reference.");
        }
        try {
            cancellationToken.ThrowIfCancellationRequested();
            ReadGlyphVariationPoints(
                glyph,
                out double[] xs,
                out double[] ys,
                out ushort[] contourEnds,
                out ushort? metricsGlyph);
            int selectedAdvanceWidth;
            if (metricsGlyph.HasValue) {
                selectedAdvanceWidth = ResolveSelectedAdvanceWidth(
                    metricsGlyph.Value,
                    workBudget,
                    cancellationToken,
                    activeGlyphs,
                    depth + 1);
            } else {
                int glyphOffset = _glyf + GlyphOffset(glyph);
                int xMin = InBounds(glyphOffset, 10) ? ReadInt16(_data, glyphOffset + 2) : 0;
                selectedAdvanceWidth = checked(baseAdvanceWidth + _variations!.GvarAdvanceWidthDelta(
                    glyph,
                    xs,
                    ys,
                    contourEnds,
                    baseAdvanceWidth,
                    LeftSideBearing(glyph),
                    xMin,
                    workBudget,
                    cancellationToken));
            }
            int delta = checked(selectedAdvanceWidth - baseAdvanceWidth);
            (_gvarAdvanceDeltas ??= new Dictionary<ushort, int>())[glyph] = delta;
            return selectedAdvanceWidth;
        } finally {
            activeGlyphs.Remove(glyph);
        }
    }

    private int LeftSideBearing(ushort glyph) => glyph < _numHMetrics
        ? ReadInt16(_data, _hmtx + glyph * 4 + 2)
        : ReadInt16(_data, _hmtx + _numHMetrics * 4 + (glyph - _numHMetrics) * 2);

    private void ReadGlyphVariationPoints(
        ushort glyph,
        out double[] xs,
        out double[] ys,
        out ushort[] contourEnds,
        out ushort? metricsGlyph) {
        xs = Array.Empty<double>();
        ys = Array.Empty<double>();
        contourEnds = Array.Empty<ushort>();
        metricsGlyph = null;
        if (glyph >= _numGlyphs) return;
        int relativeStart = GlyphOffset(glyph);
        int relativeEnd = GlyphOffset((ushort)(glyph + 1));
        if (relativeStart == relativeEnd) return;
        int offset = checked(_glyf + relativeStart);
        int end = checked(_glyf + relativeEnd);
        if (!InBounds(offset, 10) || end < offset + 10 || end > _data.Length) {
            throw new InvalidDataException("A variable glyph record is truncated.");
        }
        int contourCount = ReadInt16(_data, offset);
        if (contourCount < 0) {
            ReadCompositeVariationPoints(offset, end, out xs, out ys, out metricsGlyph);
            return;
        }
        if (contourCount == 0) return;
        if (contourCount > 32767 || offset + 10 > end - checked(contourCount * 2)) {
            throw new InvalidDataException("A variable simple-glyph contour directory is invalid.");
        }
        contourEnds = new ushort[contourCount];
        int previousEnd = -1;
        for (int contour = 0; contour < contourCount; contour++) {
            int point = ReadUInt16(_data, offset + 10 + contour * 2);
            if (point <= previousEnd) {
                throw new InvalidDataException("A variable simple-glyph contour endpoint is invalid.");
            }
            contourEnds[contour] = (ushort)point;
            previousEnd = point;
        }
        int pointCount = checked(previousEnd + 1);
        if (pointCount > 1_000_000) throw new InvalidDataException("A variable glyph has too many points.");
        int instructionLengthOffset = checked(offset + 10 + contourCount * 2);
        if (instructionLengthOffset > end - 2) {
            throw new InvalidDataException("A variable glyph instruction length is truncated.");
        }
        int instructionLength = ReadUInt16(_data, instructionLengthOffset);
        int cursor = checked(instructionLengthOffset + 2 + instructionLength);
        if (cursor > end) throw new InvalidDataException("A variable glyph instruction program is truncated.");
        var flags = new byte[pointCount];
        for (int point = 0; point < pointCount; point++) {
            if (cursor >= end) throw new InvalidDataException("A variable glyph flag array is truncated.");
            byte flag = _data[cursor++];
            flags[point] = flag;
            if ((flag & 8) == 0) continue;
            if (cursor >= end) throw new InvalidDataException("A variable glyph flag repeat is truncated.");
            int repeat = _data[cursor++];
            if (repeat > pointCount - point - 1) throw new InvalidDataException("A variable glyph flag repeat is invalid.");
            for (int index = 0; index < repeat; index++) flags[++point] = flag;
        }
        xs = DecodeVariationCoordinates(flags, ref cursor, end, xAxis: true);
        ys = DecodeVariationCoordinates(flags, ref cursor, end, xAxis: false);
    }

    private void ReadCompositeVariationPoints(
        int glyphOffset,
        int glyphEnd,
        out double[] xs,
        out double[] ys,
        out ushort? metricsGlyph) {
        const ushort argWords = 1;
        const ushort argsAreXy = 2;
        const ushort haveScale = 8;
        const ushort moreComponents = 32;
        const ushort haveXyScale = 64;
        const ushort haveTwoByTwo = 128;
        const ushort useMyMetrics = 512;
        int cursor = glyphOffset + 10;
        var xValues = new List<double>();
        var yValues = new List<double>();
        metricsGlyph = null;
        ushort flags;
        do {
            if (xValues.Count >= 4096 || cursor > glyphEnd - 4) {
                throw new InvalidDataException("A variable composite glyph directory is invalid.");
            }
            flags = ReadUInt16(_data, cursor);
            if (OfficeOpenTypeCompositeGlyph.HasConflictingTransformFlags(flags)) {
                throw new InvalidDataException("A variable composite glyph declares conflicting transform flags.");
            }
            ushort componentGlyph = ReadUInt16(_data, cursor + 2);
            if (componentGlyph >= _numGlyphs) {
                throw new InvalidDataException("A variable composite glyph references an invalid component glyph.");
            }
            if ((flags & useMyMetrics) != 0) {
                if (metricsGlyph.HasValue) {
                    throw new InvalidDataException("A variable composite glyph has multiple USE_MY_METRICS components.");
                }
                metricsGlyph = componentGlyph;
            }
            cursor += 4;
            int argumentSize = (flags & argWords) != 0 ? 4 : 2;
            if (cursor > glyphEnd - argumentSize) {
                throw new InvalidDataException("A variable composite glyph argument is truncated.");
            }
            double first = (flags & argWords) != 0
                ? ReadInt16(_data, cursor)
                : unchecked((sbyte)_data[cursor]);
            double second = (flags & argWords) != 0
                ? ReadInt16(_data, cursor + 2)
                : unchecked((sbyte)_data[cursor + 1]);
            cursor += argumentSize;
            xValues.Add((flags & argsAreXy) != 0 ? first : 0D);
            yValues.Add((flags & argsAreXy) != 0 ? second : 0D);
            int transformSize = (flags & haveScale) != 0 ? 2
                : (flags & haveXyScale) != 0 ? 4
                : (flags & haveTwoByTwo) != 0 ? 8
                : 0;
            if (cursor > glyphEnd - transformSize) {
                throw new InvalidDataException("A variable composite glyph transform is truncated.");
            }
            cursor += transformSize;
        } while ((flags & moreComponents) != 0);
        xs = xValues.ToArray();
        ys = yValues.ToArray();
    }

    private double[] DecodeVariationCoordinates(
        byte[] flags,
        ref int cursor,
        int end,
        bool xAxis) {
        int shortFlag = xAxis ? 2 : 4;
        int sameOrPositiveFlag = xAxis ? 16 : 32;
        int current = 0;
        var values = new double[flags.Length];
        for (int index = 0; index < flags.Length; index++) {
            byte flag = flags[index];
            int delta;
            if ((flag & shortFlag) != 0) {
                if (cursor >= end) throw new InvalidDataException("A variable glyph coordinate array is truncated.");
                delta = _data[cursor++];
                if ((flag & sameOrPositiveFlag) == 0) delta = -delta;
            } else if ((flag & sameOrPositiveFlag) != 0) {
                delta = 0;
            } else {
                if (cursor > end - 2) throw new InvalidDataException("A variable glyph coordinate array is truncated.");
                delta = ReadInt16(_data, cursor);
                cursor += 2;
            }
            current = checked(current + delta);
            values[index] = current;
        }
        return values;
    }
}
