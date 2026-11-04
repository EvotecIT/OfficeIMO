using System;
using System.Collections.Generic;
using System.IO;

namespace OfficeIMO.Drawing;

/// <summary>
/// Flattens a COLR v1 paint graph into the painter's-order glyph layers the outline writers paint.
/// A gradient becomes the mean of its stops and a blend mode becomes source-over; both mark the
/// glyph as approximated so the caller can report the difference.
/// </summary>
internal sealed partial class OfficeOpenTypeColorGlyphs {
    private const int MaximumPaintDepth = 64;
    private const int MaximumFlattenedLayersPerGlyph = 4096;
    private const int CompositeModeSourceOver = 3;

    private sealed class PaintGraph {
        private readonly OfficeOpenTypeReader _reader;
        private readonly int _colr, _colrEnd, _layerList, _layerCount;
        private readonly Dictionary<int, int> _basePaints;
        private readonly List<LayerRecord> _layers = new();
        private bool _approximated;

        private PaintGraph(OfficeOpenTypeReader reader, int colr, int colrEnd, int layerList, int layerCount, Dictionary<int, int> basePaints) {
            _reader = reader;
            _colr = colr;
            _colrEnd = colrEnd;
            _layerList = layerList;
            _layerCount = layerCount;
            _basePaints = basePaints;
        }

        internal static void Flatten(
            OfficeOpenTypeReader reader,
            int colr,
            int colrEnd,
            Dictionary<int, LayerRecord[]> layerMap,
            HashSet<int> approximatedGlyphs) {
            EnsureRange(colr, 34, colr, colrEnd);
            uint baseListOffset = reader.ReadUInt32(colr + 14);
            uint layerListOffset = reader.ReadUInt32(colr + 18);
            if (baseListOffset == 0) return;
            int baseList = Relative32(reader, colr, colr + 14, colr, colrEnd, 4);
            int layerList = layerListOffset == 0 ? 0 : Relative32(reader, colr, colr + 18, colr, colrEnd, 4);
            int layerCount = layerList == 0 ? 0 : CheckedCount(reader.ReadUInt32(layerList));
            if (layerList != 0) EnsureRange(layerList + 4, checked(layerCount * 4), colr, colrEnd);
            int baseCount = CheckedCount(reader.ReadUInt32(baseList));
            EnsureRange(baseList + 4, checked(baseCount * 6), colr, colrEnd);

            var basePaints = new Dictionary<int, int>(baseCount);
            for (int index = 0; index < baseCount; index++) {
                int record = baseList + 4 + index * 6;
                basePaints[reader.ReadUInt16(record)] = Relative32(reader, baseList, record + 2, colr, colrEnd, 1);
            }

            foreach (KeyValuePair<int, int> basePaint in basePaints) {
                var graph = new PaintGraph(reader, colr, colrEnd, layerList, layerCount, basePaints);
                graph.Walk(basePaint.Value, LayerTransform.Identity, 1D, -1, LayerTransform.Identity, 0);
                if (graph._layers.Count == 0) continue;
                layerMap[basePaint.Key] = graph._layers.ToArray();
                if (graph._approximated) approximatedGlyphs.Add(basePaint.Key);
            }
        }

        private static int CheckedCount(uint count) =>
            count > MaximumExpandedLayers ? throw new InvalidDataException("The COLR v1 lists exceed the managed color-glyph limits.") : (int)count;

        private void Walk(int paint, LayerTransform transform, double alpha, int clipGlyph, LayerTransform clipTransform, int depth) {
            if (depth > MaximumPaintDepth) throw new InvalidDataException("The COLR v1 paint graph is too deep or cyclic.");
            EnsureRange(paint, 1, _colr, _colrEnd);
            int next = depth + 1;
            int format = _reader.Data[paint];
            switch (format) {
                case 1: {
                    EnsureRange(paint, 6, _colr, _colrEnd);
                    int count = _reader.Data[paint + 1];
                    uint first = _reader.ReadUInt32(paint + 2);
                    if (first > (uint)_layerCount || count > _layerCount - (int)first) throw new InvalidDataException("A COLR v1 layer range is invalid.");
                    for (int index = 0; index < count; index++) {
                        int entry = _layerList + 4 + ((int)first + index) * 4;
                        Walk(Relative32(_reader, _layerList, entry, _colr, _colrEnd, 1), transform, alpha, clipGlyph, clipTransform, next);
                    }
                    return;
                }
                case 2:
                case 3:
                    EnsureRange(paint, 5, _colr, _colrEnd);
                    Emit(clipGlyph, clipTransform, new[] { new ColorStop(_reader.ReadUInt16(paint + 1), _reader.ReadF2Dot14(paint + 3)) }, alpha);
                    return;
                case 4:
                case 5:
                case 6:
                case 7:
                case 8:
                case 9:
                    _approximated = true;
                    Emit(clipGlyph, clipTransform, ReadColorStops(Child(paint), variable: format % 2 == 1), alpha);
                    return;
                case 10:
                    EnsureRange(paint, 6, _colr, _colrEnd);
                    Walk(Child(paint), transform, alpha, _reader.ReadUInt16(paint + 4), transform, next);
                    return;
                case 11:
                    EnsureRange(paint, 3, _colr, _colrEnd);
                    if (_basePaints.TryGetValue(_reader.ReadUInt16(paint + 1), out int colrGlyph)) {
                        Walk(colrGlyph, transform, alpha, clipGlyph, clipTransform, next);
                    }
                    return;
                case 32: {
                    EnsureRange(paint, 8, _colr, _colrEnd);
                    if (_reader.Data[paint + 4] != CompositeModeSourceOver) _approximated = true;
                    Walk(ReadOffset24(paint, paint + 5), transform, alpha, clipGlyph, clipTransform, next);
                    Walk(ReadOffset24(paint, paint + 1), transform, alpha, clipGlyph, clipTransform, next);
                    return;
                }
                default:
                    if (format < 12 || format > 31) {
                        _approximated = true;
                        return;
                    }
                    Walk(Child(paint), transform.Then(ReadTransform(paint, format)), alpha, clipGlyph, clipTransform, next);
                    return;
            }
        }

        // Each variable format shares the prefix of its static format, so the default instance reads the same.
        private LayerTransform ReadTransform(int paint, int format) {
            int staticFormat = format % 2 == 1 ? format - 1 : format;
            int field = paint + 4;
            double Fixed(int offset) => _reader.ReadInt32(offset) / 65536D;
            double Word(int offset) {
                EnsureRange(offset, 2, _colr, _colrEnd);
                return _reader.ReadInt16(offset);
            }
            double Angle(int offset) {
                EnsureRange(offset, 2, _colr, _colrEnd);
                return _reader.ReadF2Dot14(offset) * Math.PI;
            }
            double Scale(int offset) {
                EnsureRange(offset, 2, _colr, _colrEnd);
                return _reader.ReadF2Dot14(offset);
            }
            switch (staticFormat) {
                case 12: {
                    int affine = ReadOffset24(paint, paint + 4);
                    EnsureRange(affine, 24, _colr, _colrEnd);
                    return new LayerTransform(Fixed(affine), Fixed(affine + 8), Fixed(affine + 4), Fixed(affine + 12), Fixed(affine + 16), Fixed(affine + 20));
                }
                case 14:
                    return LayerTransform.Translation(Word(field), Word(field + 2));
                case 16:
                    return new LayerTransform(Scale(field), 0D, 0D, Scale(field + 2), 0D, 0D);
                case 18:
                    return new LayerTransform(Scale(field), 0D, 0D, Scale(field + 2), 0D, 0D).AroundCenter(Word(field + 4), Word(field + 6));
                case 20:
                    return new LayerTransform(Scale(field), 0D, 0D, Scale(field), 0D, 0D);
                case 22:
                    return new LayerTransform(Scale(field), 0D, 0D, Scale(field), 0D, 0D).AroundCenter(Word(field + 2), Word(field + 4));
                case 24:
                    return LayerTransform.Rotation(Angle(field));
                case 26:
                    return LayerTransform.Rotation(Angle(field)).AroundCenter(Word(field + 2), Word(field + 4));
                case 28:
                    return new LayerTransform(1D, Math.Tan(Angle(field)), Math.Tan(Angle(field + 2)), 1D, 0D, 0D);
                default:
                    return new LayerTransform(1D, Math.Tan(Angle(field)), Math.Tan(Angle(field + 2)), 1D, 0D, 0D)
                        .AroundCenter(Word(field + 4), Word(field + 6));
            }
        }

        private ColorStop[] ReadColorStops(int colorLine, bool variable) {
            EnsureRange(colorLine, 3, _colr, _colrEnd);
            int count = _reader.ReadUInt16(colorLine + 1);
            int stride = variable ? 10 : 6;
            EnsureRange(colorLine + 3, checked(count * stride), _colr, _colrEnd);
            var stops = new ColorStop[count];
            for (int index = 0; index < count; index++) {
                int stop = colorLine + 3 + index * stride;
                stops[index] = new ColorStop(_reader.ReadUInt16(stop + 2), _reader.ReadF2Dot14(stop + 4));
            }
            return stops;
        }

        private void Emit(int clipGlyph, LayerTransform clipTransform, ColorStop[] stops, double alpha) {
            if (clipGlyph < 0 || stops.Length == 0) {
                // A paint with no glyph to clip it floods the whole clip box; there is no contour to fill.
                _approximated = true;
                return;
            }
            if (clipGlyph >= _reader.GlyphCount) throw new InvalidDataException("A COLR v1 paint references an invalid glyph.");
            if (_layers.Count >= MaximumFlattenedLayersPerGlyph) throw new InvalidDataException("A COLR v1 glyph exceeds the managed layer limit.");
            _layers.Add(new LayerRecord(clipGlyph, stops, alpha, clipTransform));
        }

        private int Child(int paint) => ReadOffset24(paint, paint + 1);

        private int ReadOffset24(int origin, int location) {
            EnsureRange(location, 3, _colr, _colrEnd);
            byte[] data = _reader.Data;
            int offset = origin + (data[location] << 16 | data[location + 1] << 8 | data[location + 2]);
            EnsureRange(offset, 1, _colr, _colrEnd);
            return offset;
        }
    }

    private readonly struct ColorStop {
        internal ColorStop(int paletteEntry, double alpha) {
            PaletteEntry = paletteEntry;
            Alpha = alpha;
        }

        internal int PaletteEntry { get; }
        internal double Alpha { get; }
    }
}

/// <summary>Affine transform in font units, using the OpenType x' = xx·x + xy·y + dx convention.</summary>
internal readonly struct LayerTransform {
    internal static readonly LayerTransform Identity = new(1D, 0D, 0D, 1D, 0D, 0D);

    internal LayerTransform(double xx, double xy, double yx, double yy, double dx, double dy) {
        Xx = xx;
        Xy = xy;
        Yx = yx;
        Yy = yy;
        Dx = dx;
        Dy = dy;
    }

    internal double Xx { get; }
    internal double Xy { get; }
    internal double Yx { get; }
    internal double Yy { get; }
    internal double Dx { get; }
    internal double Dy { get; }

    internal static LayerTransform Translation(double dx, double dy) => new(1D, 0D, 0D, 1D, dx, dy);

    internal static LayerTransform Rotation(double angle) {
        double cos = Math.Cos(angle), sin = Math.Sin(angle);
        return new LayerTransform(cos, -sin, sin, cos, 0D, 0D);
    }

    internal LayerTransform AroundCenter(double x, double y) => Translation(x, y).Then(this).Then(Translation(-x, -y));

    /// <summary>Returns this transform applied after <paramref name="inner"/>, as a nested paint requires.</summary>
    internal LayerTransform Then(LayerTransform inner) => new(
        Xx * inner.Xx + Xy * inner.Yx,
        Xx * inner.Xy + Xy * inner.Yy,
        Yx * inner.Xx + Yy * inner.Yx,
        Yx * inner.Xy + Yy * inner.Yy,
        Dx + Xx * inner.Dx + Xy * inner.Dy,
        Dy + Yx * inner.Dx + Yy * inner.Dy);
}
