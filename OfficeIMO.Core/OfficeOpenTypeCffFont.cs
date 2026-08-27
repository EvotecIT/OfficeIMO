using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>First-party managed CFF1/CFF2 measurement and outline program.</summary>
internal sealed class OfficeOpenTypeCffFont : IOfficeCffBoundedFontProgram, IOfficeFontBaselineMetrics, IOfficeVariableFontProgram {
    private readonly byte[] _data;
    private readonly OfficeOpenTypeReader _reader;
    private readonly OfficeCffFontData _cff;
    private readonly OfficeFontVariationModel _variations;
    private readonly OfficeOpenTypeHvarMetrics? _hvar;
    private readonly OfficeOpenTypeKerning _kerning;
    private readonly int _ascender;
    private readonly int _descender;
    private readonly int _lineGap;
    private readonly string _fingerprint;

    private OfficeOpenTypeCffFont(
        byte[] data,
        OfficeOpenTypeReader reader,
        OfficeCffFontData cff,
        OfficeFontVariationModel variations) {
        _data = (byte[])data.Clone();
        _reader = reader;
        _cff = cff;
        _variations = variations;
        _kerning = OfficeOpenTypeKerning.FromReader(reader);
        _hvar = variations.IsVariable
            ? OfficeOpenTypeHvarMetrics.TryParse(reader, variations)
            : null;
        OfficeOpenTypeMvarMetrics? mvar = variations.IsVariable
            ? OfficeOpenTypeMvarMetrics.TryParse(reader, variations)
            : null;
        _ascender = checked(reader.Ascender + (mvar?.HorizontalAscenderDelta ?? 0));
        _descender = checked(reader.Descender + (mvar?.HorizontalDescenderDelta ?? 0));
        _lineGap = checked(reader.LineGap + (mvar?.HorizontalLineGapDelta ?? 0));
        _fingerprint = ComputeFingerprint(data, variations.Identity);
    }

    internal static OfficeOpenTypeCffFont? TryLoad(
        byte[] data,
        IReadOnlyDictionary<string, float>? variationValues,
        out string? error) {
        error = null;
        try {
            OfficeOpenTypeReader? reader = OfficeOpenTypeReader.TryCreate(data);
            if (reader == null || !reader.TryGetTable("CFF ", out _, out _) && !reader.TryGetTable("CFF2", out _, out _)) return null;
            OfficeFontVariationModel variations = OfficeFontVariationModel.Create(reader, variationValues);
            OfficeCffFontData cff = OfficeCffFontData.Parse(reader, variations);
            return new OfficeOpenTypeCffFont(data, reader, cff, variations);
        } catch (Exception exception) when (exception is InvalidDataException
                                            || exception is NotSupportedException
                                            || exception is ArgumentException
                                            || exception is OverflowException
                                            || exception is IndexOutOfRangeException) {
            error = exception.Message;
            return null;
        }
    }

    public string Fingerprint => _fingerprint;
    IReadOnlyDictionary<string, float> IOfficeVariableFontProgram.VariationCoordinatesForShaping =>
        _variations.DesignCoordinates;
    public string? DisplayName => _reader.ReadDisplayName();
    public int? CollectionIndex => null;
    public int UnitsPerEm => _reader.UnitsPerEm;
    public bool IsOpenTypeCff => true;
    public bool ProvidesComplexTextLayout => false;
    public double LineSpacingRatio => Math.Max(1, _ascender - _descender + _lineGap) / (double)_reader.UnitsPerEm;
    internal bool IsVariable => _variations.IsVariable || _cff.IsCff2;

    public byte[] GetFontDataForShaping() => (byte[])_data.Clone();

    public bool HasGlyphs(string text) {
        if (text == null) throw new ArgumentNullException(nameof(text));
        for (int index = 0; index < text.Length;) {
            int scalar = ReadScalar(text, ref index);
            if (OfficeTextElements.IsIgnorableFontCoverageScalar(scalar)) continue;
            if (_reader.MapGlyph(scalar) == 0) return false;
        }
        return true;
    }

    public double Measure(string text, double fontSize) {
        ValidateTextAndSize(text, fontSize);
        double scale = Scale(fontSize);
        long width = 0;
        int? previousGlyph = null;
        for (int index = 0; index < text.Length;) {
            int scalar = ReadScalar(text, ref index);
            if (OfficeTextElements.IsIgnorableFontCoverageScalar(scalar)) continue;
            int glyphId = _reader.MapGlyph(scalar);
            if (previousGlyph.HasValue) width = checked(width + _kerning.Adjustment(previousGlyph.Value, glyphId));
            width = checked(width + AdvanceWidth(glyphId));
            previousGlyph = glyphId;
        }
        return width * scale;
    }

    public IReadOnlyList<double> MeasureTextElements(IReadOnlyList<string> elements, double fontSize) {
        if (elements == null) throw new ArgumentNullException(nameof(elements));
        ValidateSize(fontSize);
        var result = new double[elements.Count];
        double scale = Scale(fontSize);
        int? previousGlyph = null;
        for (int elementIndex = 0; elementIndex < elements.Count; elementIndex++) {
            string text = elements[elementIndex];
            long width = 0;
            for (int textIndex = 0; textIndex < text.Length;) {
                int scalar = ReadScalar(text, ref textIndex);
                if (OfficeTextElements.IsIgnorableFontCoverageScalar(scalar)) continue;
                int glyphId = _reader.MapGlyph(scalar);
                if (previousGlyph.HasValue) width = checked(width + _kerning.Adjustment(previousGlyph.Value, glyphId));
                width = checked(width + AdvanceWidth(glyphId));
                previousGlyph = glyphId;
            }
            result[elementIndex] = width * scale;
        }
        return result;
    }

    public double LineHeight(double fontSize) {
        ValidateSize(fontSize);
        return Math.Max(1, _ascender - _descender) * Scale(fontSize);
    }

    public double BaselineOffset(double fontSize) {
        ValidateSize(fontSize);
        return _ascender * Scale(fontSize);
    }

    public List<List<OfficePoint>> GetTextContours(string text, double x, double y, double fontSize) =>
        GetTextContoursBounded(text, x, y, fontSize, int.MaxValue, CancellationToken.None);

    public List<List<OfficePoint>> GetTextContoursBounded(
        string text,
        double x,
        double y,
        double fontSize,
        int maximumPointCount,
        CancellationToken cancellationToken) => GetTextContoursBounded(
            text,
            x,
            y,
            fontSize,
            maximumPointCount,
            cancellationToken,
            new OfficeCffOperationBudget());

    public List<List<OfficePoint>> GetTextContoursBounded(
        string text,
        double x,
        double y,
        double fontSize,
        int maximumPointCount,
        CancellationToken cancellationToken,
        OfficeCffOperationBudget operationBudget) {
        ValidateTextAndSize(text, fontSize);
        ValidatePosition(x, y);
        if (maximumPointCount <= 0) throw new ArgumentOutOfRangeException(nameof(maximumPointCount));
        if (operationBudget == null) throw new ArgumentNullException(nameof(operationBudget));
        var contours = new List<List<OfficePoint>>();
        double scale = Scale(fontSize);
        double cursor = x;
        double baseline = EnsureFiniteGeometry(y + _ascender * scale);
        int pointCount = 0;
        int? previousGlyph = null;
        for (int index = 0; index < text.Length;) {
            cancellationToken.ThrowIfCancellationRequested();
            int scalar = ReadScalar(text, ref index);
            if (OfficeTextElements.IsIgnorableFontCoverageScalar(scalar)) continue;
            int glyphId = _reader.MapGlyph(scalar);
            if (previousGlyph.HasValue) {
                cursor = EnsureFiniteGeometry(cursor + _kerning.Adjustment(previousGlyph.Value, glyphId) * scale);
            }
            RenderGlyph(glyphId, cursor, baseline, scale, contours, ref pointCount, maximumPointCount, cancellationToken, operationBudget);
            cursor = EnsureFiniteGeometry(cursor + AdvanceWidth(glyphId) * scale);
            previousGlyph = glyphId;
        }
        return contours;
    }

    public bool TryGetGlyphMetrics(int scalar, out int glyphId, out int advanceWidth) {
        glyphId = _reader.MapGlyph(scalar);
        advanceWidth = glyphId == 0 ? 0 : AdvanceWidth(glyphId);
        return glyphId != 0;
    }

    public double MeasureShapedText(string text, OfficeTextShapingResult result, double fontSize) {
        PositionedGlyph[] glyphs = ValidateShapedGlyphs(text, result);
        ValidateSize(fontSize);
        long width = 0;
        for (int index = 0; index < glyphs.Length; index++) width = checked(width + glyphs[index].AdvanceWidth);
        return Math.Abs(width * Scale(fontSize));
    }

    public List<List<OfficePoint>> GetShapedTextContours(
        string text,
        OfficeTextShapingResult result,
        double x,
        double y,
        double fontSize) => GetShapedTextContoursBounded(
            text,
            result,
            x,
            y,
            fontSize,
            int.MaxValue,
            CancellationToken.None);

    public List<List<OfficePoint>> GetShapedTextContoursBounded(
        string text,
        OfficeTextShapingResult result,
        double x,
        double y,
        double fontSize,
        int maximumPointCount,
        CancellationToken cancellationToken) => GetShapedTextContoursBounded(
            text,
            result,
            x,
            y,
            fontSize,
            maximumPointCount,
            cancellationToken,
            new OfficeCffOperationBudget());

    public List<List<OfficePoint>> GetShapedTextContoursBounded(
        string text,
        OfficeTextShapingResult result,
        double x,
        double y,
        double fontSize,
        int maximumPointCount,
        CancellationToken cancellationToken,
        OfficeCffOperationBudget operationBudget) {
        PositionedGlyph[] glyphs = ValidateShapedGlyphs(text, result);
        ValidateSize(fontSize);
        ValidatePosition(x, y);
        if (maximumPointCount <= 0) throw new ArgumentOutOfRangeException(nameof(maximumPointCount));
        if (operationBudget == null) throw new ArgumentNullException(nameof(operationBudget));
        var contours = new List<List<OfficePoint>>();
        double scale = Scale(fontSize);
        long totalAdvance = 0;
        for (int index = 0; index < glyphs.Length; index++) totalAdvance = checked(totalAdvance + glyphs[index].AdvanceWidth);
        bool negativeDirection = totalAdvance < 0;
        double cursor = negativeDirection ? EnsureFiniteGeometry(x - (totalAdvance * scale)) : x;
        double baseline = EnsureFiniteGeometry(y + _ascender * scale);
        int pointCount = 0;
        for (int index = 0; index < glyphs.Length; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            PositionedGlyph glyph = glyphs[index];
            if (negativeDirection) cursor = EnsureFiniteGeometry(cursor + glyph.AdvanceWidth * scale);
            RenderGlyph(
                glyph.GlyphId,
                EnsureFiniteGeometry(cursor + glyph.OffsetX * scale),
                EnsureFiniteGeometry(baseline - glyph.OffsetY * scale),
                scale,
                contours,
                ref pointCount,
                maximumPointCount,
                cancellationToken,
                operationBudget);
            if (!negativeDirection) cursor = EnsureFiniteGeometry(cursor + glyph.AdvanceWidth * scale);
        }
        return contours;
    }

    private void RenderGlyph(
        int glyphId,
        double x,
        double baseline,
        double scale,
        List<List<OfficePoint>> contours,
        ref int pointCount,
        int maximumPointCount,
        CancellationToken cancellationToken,
        OfficeCffOperationBudget operationBudget) {
        if (glyphId < 0 || glyphId >= _cff.GlyphCount) throw new InvalidDataException("A CFF glyph identifier is outside CharStrings.");
        var sink = new CffPathSink(x, baseline, scale, contours, pointCount, maximumPointCount, cancellationToken);
        var interpreter = new OfficeType2CharStringInterpreter(_cff, glyphId, sink, cancellationToken, operationBudget);
        interpreter.Render(_cff.GetCharString(glyphId));
        pointCount = sink.PointCount;
    }

    private PositionedGlyph[] ValidateShapedGlyphs(string text, OfficeTextShapingResult result) {
        if (text == null) throw new ArgumentNullException(nameof(text));
        if (result == null) throw new ArgumentNullException(nameof(result));
        if (text.Length > 0 && result.Glyphs.Count == 0) throw new ArgumentException("Drawing text shaping provider returned no glyphs for non-empty text.", nameof(result));
        var glyphs = new PositionedGlyph[result.Glyphs.Count];
        for (int index = 0; index < result.Glyphs.Count; index++) {
            OfficeShapedGlyph glyph = result.Glyphs[index];
            if (glyph.GlyphId <= 0 || glyph.GlyphId >= _reader.GlyphCount) {
                throw new ArgumentException("Drawing text shaping provider returned a glyph outside the selected font range.", nameof(result));
            }
            if (glyph.TextIndex < 0 || glyph.TextIndex > text.Length || glyph.UnicodeText.Length > text.Length - glyph.TextIndex
                || !string.Equals(text.Substring(glyph.TextIndex, glyph.UnicodeText.Length), glyph.UnicodeText, StringComparison.Ordinal)) {
                throw new ArgumentException("Drawing text shaping provider returned a Unicode mapping outside the source text.", nameof(result));
            }
            glyphs[index] = new PositionedGlyph(
                glyph.GlyphId,
                glyph.AdvanceWidth ?? AdvanceWidth(glyph.GlyphId),
                glyph.OffsetX,
                glyph.OffsetY);
        }
        return glyphs;
    }

    private double Scale(double fontSize) => fontSize / _reader.UnitsPerEm;

    private int AdvanceWidth(int glyphId) => checked(
        _reader.AdvanceWidth(glyphId) + (_hvar?.AdvanceWidthDelta(glyphId) ?? 0));

    private static void ValidateTextAndSize(string text, double fontSize) {
        if (text == null) throw new ArgumentNullException(nameof(text));
        ValidateSize(fontSize);
    }

    private static void ValidateSize(double fontSize) {
        if (fontSize <= 0D || double.IsNaN(fontSize) || double.IsInfinity(fontSize)) throw new ArgumentOutOfRangeException(nameof(fontSize));
    }

    private static void ValidatePosition(double x, double y) {
        if (double.IsNaN(x) || double.IsInfinity(x)) throw new ArgumentOutOfRangeException(nameof(x));
        if (double.IsNaN(y) || double.IsInfinity(y)) throw new ArgumentOutOfRangeException(nameof(y));
    }

    private static double EnsureFiniteGeometry(double value) {
        if (double.IsNaN(value) || double.IsInfinity(value)) {
            throw new ArgumentOutOfRangeException("fontSize", "CFF outline positioning produced non-finite geometry.");
        }
        return value;
    }

    private static int ReadScalar(string text, ref int index) {
        char first = text[index++];
        return char.IsHighSurrogate(first) && index < text.Length && char.IsLowSurrogate(text[index])
            ? char.ConvertToUtf32(first, text[index++])
            : first;
    }

    private static string ComputeFingerprint(byte[] data, string variationIdentity) {
        using HashAlgorithm hash = SHA256.Create();
        byte[] identity = Encoding.UTF8.GetBytes("OfficeIMO.CFF.FontProgram.v1\n" + variationIdentity);
        hash.TransformBlock(identity, 0, identity.Length, identity, 0);
        hash.TransformFinalBlock(data, 0, data.Length);
        return "sha256:" + ToLowerHex(hash.Hash!);
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

    private readonly struct PositionedGlyph {
        internal PositionedGlyph(int glyphId, int advanceWidth, int offsetX, int offsetY) {
            GlyphId = glyphId;
            AdvanceWidth = advanceWidth;
            OffsetX = offsetX;
            OffsetY = offsetY;
        }

        internal int GlyphId { get; }
        internal int AdvanceWidth { get; }
        internal int OffsetX { get; }
        internal int OffsetY { get; }
    }

    private sealed class CffPathSink : IOfficeCffPathSink {
        private const double CurveTolerance = 0.05D;
        private const int MaximumCurveDepth = 12;
        private readonly double _originX;
        private readonly double _baseline;
        private readonly double _scale;
        private readonly List<List<OfficePoint>> _contours;
        private readonly int _maximumPointCount;
        private readonly CancellationToken _cancellationToken;
        private List<OfficePoint>? _current;
        private OfficePoint _currentPoint;

        internal CffPathSink(
            double originX,
            double baseline,
            double scale,
            List<List<OfficePoint>> contours,
            int existingPointCount,
            int maximumPointCount,
            CancellationToken cancellationToken) {
            _originX = originX;
            _baseline = baseline;
            _scale = scale;
            _contours = contours;
            PointCount = existingPointCount;
            _maximumPointCount = maximumPointCount;
            _cancellationToken = cancellationToken;
        }

        internal int PointCount { get; private set; }

        public void MoveTo(double x, double y) {
            CloseContour();
            _current = new List<OfficePoint>();
            Add(Transform(x, y));
        }

        public void LineTo(double x, double y) => Add(Transform(x, y));

        public void CurveTo(double control1X, double control1Y, double control2X, double control2Y, double x, double y) {
            EnsureCurrent();
            FlattenCubic(
                _currentPoint,
                Transform(control1X, control1Y),
                Transform(control2X, control2Y),
                Transform(x, y),
                0);
        }

        public void CloseContour() {
            if (_current == null) return;
            if (_current.Count >= 3) {
                OfficePoint first = _current[0];
                if (_current[_current.Count - 1] != first) Add(first);
                _contours.Add(_current);
            }
            _current = null;
        }

        private OfficePoint Transform(double x, double y) {
            var point = new OfficePoint(_originX + x * _scale, _baseline - y * _scale);
            EnsureFinitePoint(point);
            return point;
        }

        private void FlattenCubic(OfficePoint start, OfficePoint control1, OfficePoint control2, OfficePoint end, int depth) {
            double flatness = Math.Max(DistanceToLine(control1, start, end), DistanceToLine(control2, start, end));
            if (depth >= MaximumCurveDepth || flatness <= CurveTolerance) {
                Add(end);
                return;
            }
            OfficePoint startControl = Mid(start, control1);
            OfficePoint controls = Mid(control1, control2);
            OfficePoint controlEnd = Mid(control2, end);
            OfficePoint leftControl = Mid(startControl, controls);
            OfficePoint rightControl = Mid(controls, controlEnd);
            OfficePoint midpoint = Mid(leftControl, rightControl);
            FlattenCubic(start, startControl, leftControl, midpoint, depth + 1);
            FlattenCubic(midpoint, rightControl, controlEnd, end, depth + 1);
        }

        private void Add(OfficePoint point) {
            _cancellationToken.ThrowIfCancellationRequested();
            EnsureFinitePoint(point);
            if (PointCount >= _maximumPointCount) throw new InvalidOperationException("Font outline expansion exceeded the configured point budget.");
            EnsureCurrent();
            _current!.Add(point);
            _currentPoint = point;
            PointCount++;
        }

        private void EnsureCurrent() {
            if (_current == null) _current = new List<OfficePoint>();
        }

        private static OfficePoint Mid(OfficePoint left, OfficePoint right) {
            var midpoint = new OfficePoint((left.X / 2D) + (right.X / 2D), (left.Y / 2D) + (right.Y / 2D));
            EnsureFinitePoint(midpoint);
            return midpoint;
        }

        private static void EnsureFinitePoint(OfficePoint point) {
            if (double.IsNaN(point.X) || double.IsInfinity(point.X) ||
                double.IsNaN(point.Y) || double.IsInfinity(point.Y)) {
                throw new InvalidDataException("A transformed CFF path coordinate is not finite.");
            }
        }

        private static double DistanceToLine(OfficePoint point, OfficePoint start, OfficePoint end) {
            double deltaX = end.X - start.X;
            double deltaY = end.Y - start.Y;
            double length = Math.Sqrt((deltaX * deltaX) + (deltaY * deltaY));
            if (length <= double.Epsilon) {
                double dx = point.X - start.X;
                double dy = point.Y - start.Y;
                return Math.Sqrt((dx * dx) + (dy * dy));
            }
            return Math.Abs((deltaY * point.X) - (deltaX * point.Y) + (end.X * start.Y) - (end.Y * start.X)) / length;
        }
    }
}
