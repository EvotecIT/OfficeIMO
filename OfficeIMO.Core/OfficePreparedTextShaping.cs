using System;
using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>Retains a complete horizontal monochrome result with its exact font program.</summary>
internal sealed class OfficePreparedTextShaping {
    private readonly int[] _firstGlyph, _endGlyph, _clusterEnd;
    private readonly OfficeTrueTypeFont.ShapedTextRun? _trueTypeRun;
    private readonly CancellationToken _cancellationToken;
    private readonly bool? _negativeAdvance;

    internal OfficePreparedTextShaping(IOfficeFontProgram program, string text, OfficeTextShapingResult result,
        CancellationToken cancellationToken = default)
        : this(program, text, result, null, cancellationToken) { }

    private OfficePreparedTextShaping(IOfficeFontProgram program, string text, OfficeTextShapingResult result,
        OfficeTrueTypeFont.ShapedTextRun? nativeRun, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        _cancellationToken = cancellationToken;
        Program = program ?? throw new ArgumentNullException(nameof(program));
        Text = text ?? throw new ArgumentNullException(nameof(text));
        Result = result ?? throw new ArgumentNullException(nameof(result));
        if (text.Length == 0 || result.Glyphs.Count == 0 || result.Direction == OfficeTextDirection.TopToBottom ||
            text.IndexOfAny(new[] { '\r', '\n' }) >= 0 || program is not IOfficeBoundedFontProgram)
            throw new ArgumentException("Prepared shaping requires nonempty horizontal text and a bounded font program.", nameof(result));
        _firstGlyph = new int[text.Length + 1];
        _endGlyph = new int[text.Length + 1];
        _clusterEnd = new int[text.Length + 1];
        for (int index = 0; index < _firstGlyph.Length; index++) {
            if ((index & 255) == 0) cancellationToken.ThrowIfCancellationRequested();
            _firstGlyph[index] = -1;
        }
        var covered = new bool[text.Length];
        for (int index = 0; index < result.Glyphs.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            OfficeShapedGlyph glyph = result.Glyphs[index];
            int cluster = result.GetLogicalClusterStart(index);
            if (cluster < 0 || cluster >= text.Length || cluster > glyph.TextIndex ||
                glyph.TextIndex > text.Length || glyph.UnicodeText.Length > text.Length - glyph.TextIndex ||
                string.CompareOrdinal(text, glyph.TextIndex, glyph.UnicodeText, 0, glyph.UnicodeText.Length) != 0 ||
                glyph.AdvanceHeight.GetValueOrDefault() != 0)
                throw new ArgumentException("Prepared shaping has an invalid horizontal source mapping.", nameof(result));
            if (program is IOfficeColorFontProgram color && color.HasColorGlyph(glyph.GlyphId))
                throw new ArgumentException("Prepared monochrome shaping cannot flatten color glyphs.", nameof(result));
            for (int offset = glyph.TextIndex; offset < glyph.TextIndex + glyph.UnicodeText.Length; offset++) {
                if ((offset & 255) == 0) cancellationToken.ThrowIfCancellationRequested();
                if (covered[offset]) throw new ArgumentException("Prepared shaping repeats logical source coverage.", nameof(result));
                covered[offset] = true;
            }
            if (_firstGlyph[cluster] < 0) _firstGlyph[cluster] = index;
            else if (_endGlyph[cluster] != index)
                throw new ArgumentException("Prepared shaping requires contiguous glyphs within each cluster.", nameof(result));
            _endGlyph[cluster] = index + 1;
        }
        int previous = -1;
        for (int index = 0; index < text.Length; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!covered[index]) throw new ArgumentException("Prepared shaping omits logical source coverage.", nameof(result));
            if (_firstGlyph[index] < 0) continue;
            if (char.IsLowSurrogate(text[index]) && index > 0 && char.IsHighSurrogate(text[index - 1]))
                throw new ArgumentException("Prepared shaping splits a source scalar.", nameof(result));
            if (previous >= 0) _clusterEnd[previous] = index;
            previous = index;
        }
        if (_firstGlyph[0] < 0) throw new ArgumentException("Prepared shaping has no initial source cluster.", nameof(result));
        _clusterEnd[previous] = text.Length;
        for (int index = 0; index < result.Glyphs.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            OfficeShapedGlyph glyph = result.Glyphs[index];
            if (glyph.TextIndex + glyph.UnicodeText.Length > _clusterEnd[result.GetLogicalClusterStart(index)])
                throw new ArgumentException("Prepared shaping maps text across distinct clusters.", nameof(result));
        }
        _trueTypeRun = nativeRun ?? (program as OfficeTrueTypeFont)?.CreateShapedTextRun(text, result, cancellationToken: cancellationToken);
        if (_trueTypeRun != null) _negativeAdvance = _trueTypeRun.HasNegativeAdvance;
        else if (program is OfficeOpenTypeCffFont cff) _negativeAdvance = cff.GetShapedTextAdvance(text, result) < 0L;
        else program.MeasureShapedText(text, result, 1D);
    }

    internal IOfficeFontProgram Program { get; }
    internal string Text { get; }
    internal OfficeTextShapingResult Result { get; }

    internal OfficeTrueTypeFont.ShapedTextRun? WithMetricFont(OfficeTrueTypeFont font) =>
        _trueTypeRun?.WithMetricFont(font);

    internal double Measure(double size) {
        ValidateSize(size);
        return _trueTypeRun?.Measure(size) ?? Program.MeasureShapedText(Text, Result, size);
    }

    internal List<List<OfficePoint>> GetContours(double x, double y, double size, int maximumPointCount,
        CancellationToken cancellationToken, OfficeCffOperationBudget? cffBudget = null) {
        ValidateSize(size);
        cancellationToken.ThrowIfCancellationRequested();
        if (maximumPointCount <= 0) throw new ArgumentOutOfRangeException(nameof(maximumPointCount));
        if (_trueTypeRun != null)
            return _trueTypeRun.GetContours(x, y, size, maximumPointCount, cancellationToken);
        return Program is IOfficeCffBoundedFontProgram cff
            ? cff.GetShapedTextContoursBounded(Text, Result, x, y, size, maximumPointCount, cancellationToken,
                cffBudget ?? new OfficeCffOperationBudget())
            : ((IOfficeBoundedFontProgram)Program).GetShapedTextContoursBounded(Text, Result, x, y, size,
                maximumPointCount, cancellationToken);
    }

    /// <summary>Slices whole logical clusters without reshaping or dropping continuation glyphs.</summary>
    internal bool TrySlice(int logicalStart, int length, out OfficePreparedTextShaping? slice) {
        slice = null;
        _cancellationToken.ThrowIfCancellationRequested();
        if (logicalStart < 0 || logicalStart >= Text.Length || length <= 0 || length > Text.Length - logicalStart) return false;
        int logicalEnd = logicalStart + length;
        if (_firstGlyph[logicalStart] < 0 || logicalEnd < Text.Length && _firstGlyph[logicalEnd] < 0) return false;
        if (logicalStart == 0 && logicalEnd == Text.Length) { slice = this; return true; }
        // Unknown bounded programs need a whole result; their pen normalization is not qualified here.
        if (!_negativeAdvance.HasValue) return false;
        int first = Result.Glyphs.Count, end = 0, count = 0;
        for (int cluster = logicalStart; cluster < logicalEnd; cluster = _clusterEnd[cluster]) {
            _cancellationToken.ThrowIfCancellationRequested();
            first = Math.Min(first, _firstGlyph[cluster]);
            end = Math.Max(end, _endGlyph[cluster]);
            count += _endGlyph[cluster] - _firstGlyph[cluster];
        }
        if (end - first != count) return false;
        var glyphs = new OfficeShapedGlyph[count];
        var adjustments = new int[count];
        var clusters = new int[count];
        OfficeTrueTypeFont.PositionedGlyph[]? nativeGlyphs = _trueTypeRun == null
            ? null : new OfficeTrueTypeFont.PositionedGlyph[count];
        for (int index = 0; index < count; index++) {
            _cancellationToken.ThrowIfCancellationRequested();
            OfficeShapedGlyph glyph = Result.Glyphs[first + index];
            glyphs[index] = glyph.WithTextIndex(glyph.TextIndex - logicalStart);
            adjustments[index] = Result.GetAdvanceAdjustment(first + index);
            clusters[index] = Result.GetLogicalClusterStart(first + index) - logicalStart;
            if (nativeGlyphs != null) nativeGlyphs[index] = _trueTypeRun!._glyphs[first + index];
        }
        var result = new OfficeTextShapingResult(glyphs, adjustments, Result.Direction, clusters);
        var native = nativeGlyphs == null ? null : new OfficeTrueTypeFont.ShapedTextRun(
            (OfficeTrueTypeFont)Program, nativeGlyphs, Result.Direction);
        // Absolute measurement cannot expose a stationary cluster's signed tracking contribution.
        if (native?.HasZeroTrackedAdvance == true) return false;
        var candidate = new OfficePreparedTextShaping(Program, Text.Substring(logicalStart, length), result, native, _cancellationToken);
        // A sign change switches before/after-glyph pen movement, which cannot be repaired by translation.
        if (candidate._negativeAdvance != _negativeAdvance) return false;
        slice = candidate;
        return true;
    }

    private static void ValidateSize(double size) {
        if (size <= 0D || double.IsNaN(size) || double.IsInfinity(size)) throw new ArgumentOutOfRangeException(nameof(size));
    }
}
