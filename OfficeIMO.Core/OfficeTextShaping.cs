using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace OfficeIMO.Drawing;

/// <summary>
/// Provides externally shaped glyph runs to OfficeIMO renderers without requiring a shaping-engine dependency.
/// </summary>
/// <remarks>
/// A host can adapt HarfBuzz, DirectWrite, Core Text, or another shaping engine once and reuse it
/// across PDF and other OfficeIMO renderers. Returning <see langword="null"/> declines a run and
/// lets the consuming renderer use its dependency-free fallback.
/// </remarks>
public interface IOfficeTextShapingProvider {
    /// <summary>Shapes a Unicode run into glyphs in visual write order.</summary>
    /// <param name="request">Text, font bytes, metrics, and language context for the run.</param>
    /// <returns>A shaped glyph run, or <see langword="null"/> to use the renderer fallback.</returns>
    OfficeTextShapingResult? ShapeText(OfficeTextShapingRequest request);
}

/// <summary>Describes a Unicode text run and font passed to a shared shaping provider.</summary>
public sealed class OfficeTextShapingRequest {
    private readonly byte[] _fontData;
    private readonly object? _fontProgramCacheKey;
    private readonly IReadOnlyDictionary<string, float> _variationCoordinates;
    private readonly OfficeTextFeatureSettings _featureSettings;

    /// <summary>Creates an immutable text-shaping request.</summary>
    /// <param name="text">Original UTF-16 text to shape.</param>
    /// <param name="fontName">Display or resource name of the selected font.</param>
    /// <param name="fontData">Complete font-program snapshot.</param>
    /// <param name="isOpenTypeCff">True for OpenType/CFF outlines; false for TrueType outlines.</param>
    /// <param name="unitsPerEm">Font design units per em.</param>
    /// <param name="direction">Resolved base direction for the run.</param>
    /// <param name="language">Optional BCP 47 language hint.</param>
    public OfficeTextShapingRequest(
        string text,
        string fontName,
        byte[] fontData,
        bool isOpenTypeCff,
        int unitsPerEm,
        OfficeTextDirection direction = OfficeTextDirection.Auto,
        string? language = null)
        : this(
            text,
            fontName,
            fontData,
            isOpenTypeCff,
            unitsPerEm,
            direction,
            language,
            default,
            fontCollectionIndex: null,
            variationCoordinates: null,
            cloneFontData: true,
            featureSettings: null) {
    }

    /// <summary>Creates an immutable text-shaping request with explicit OpenType feature values.</summary>
    public OfficeTextShapingRequest(
        string text,
        string fontName,
        byte[] fontData,
        bool isOpenTypeCff,
        int unitsPerEm,
        OfficeTextDirection direction,
        string? language,
        OfficeTextFeatureSettings featureSettings)
        : this(
            text,
            fontName,
            fontData,
            isOpenTypeCff,
            unitsPerEm,
            direction,
            language,
            default,
            fontCollectionIndex: null,
            variationCoordinates: null,
            cloneFontData: true,
            featureSettings: featureSettings) {
    }

    /// <summary>Creates an immutable text-shaping request with cooperative cancellation.</summary>
    public OfficeTextShapingRequest(
        string text,
        string fontName,
        byte[] fontData,
        bool isOpenTypeCff,
        int unitsPerEm,
        OfficeTextDirection direction,
        string? language,
        System.Threading.CancellationToken cancellationToken)
        : this(
            text,
            fontName,
            fontData,
            isOpenTypeCff,
            unitsPerEm,
            direction,
            language,
            cancellationToken,
            fontCollectionIndex: null,
            variationCoordinates: null,
            cloneFontData: true,
            featureSettings: null) {
    }

    /// <summary>Creates an immutable text-shaping request for a selected face in a font collection.</summary>
    public OfficeTextShapingRequest(
        string text,
        string fontName,
        byte[] fontData,
        bool isOpenTypeCff,
        int unitsPerEm,
        OfficeTextDirection direction,
        string? language,
        System.Threading.CancellationToken cancellationToken,
        int? fontCollectionIndex)
        : this(
            text,
            fontName,
            fontData,
            isOpenTypeCff,
            unitsPerEm,
            direction,
            language,
            cancellationToken,
            fontCollectionIndex,
            variationCoordinates: null,
            cloneFontData: true,
            featureSettings: null) {
    }

    /// <summary>Creates an immutable text-shaping request for a selected variable-font instance.</summary>
    public OfficeTextShapingRequest(
        string text,
        string fontName,
        byte[] fontData,
        bool isOpenTypeCff,
        int unitsPerEm,
        OfficeTextDirection direction,
        string? language,
        System.Threading.CancellationToken cancellationToken,
        int? fontCollectionIndex,
        IReadOnlyDictionary<string, float>? variationCoordinates)
        : this(
            text,
            fontName,
            fontData,
            isOpenTypeCff,
            unitsPerEm,
            direction,
            language,
            cancellationToken,
            fontCollectionIndex,
            variationCoordinates,
            cloneFontData: true,
            featureSettings: null) {
    }

    /// <summary>Creates an immutable text-shaping request for a selected variable-font instance with explicit OpenType feature values.</summary>
    public OfficeTextShapingRequest(
        string text,
        string fontName,
        byte[] fontData,
        bool isOpenTypeCff,
        int unitsPerEm,
        OfficeTextDirection direction,
        string? language,
        System.Threading.CancellationToken cancellationToken,
        int? fontCollectionIndex,
        IReadOnlyDictionary<string, float>? variationCoordinates,
        OfficeTextFeatureSettings featureSettings)
        : this(
            text,
            fontName,
            fontData,
            isOpenTypeCff,
            unitsPerEm,
            direction,
            language,
            cancellationToken,
            fontCollectionIndex,
            variationCoordinates,
            cloneFontData: true,
            featureSettings: featureSettings) {
    }

    internal OfficeTextShapingRequest(
        string text,
        string fontName,
        byte[] fontData,
        bool isOpenTypeCff,
        int unitsPerEm,
        OfficeTextDirection direction,
        string? language,
        System.Threading.CancellationToken cancellationToken,
        int? fontCollectionIndex,
        IReadOnlyDictionary<string, float>? variationCoordinates,
        bool cloneFontData,
        object? fontProgramCacheKey = null,
        OfficeTextFeatureSettings? featureSettings = null) {
        Text = text ?? throw new ArgumentNullException(nameof(text));
        if (fontData == null) {
            throw new ArgumentNullException(nameof(fontData));
        }
        if (unitsPerEm <= 0) {
            throw new ArgumentOutOfRangeException(nameof(unitsPerEm), "Text shaping units per em must be positive.");
        }
        if (direction != OfficeTextDirection.Auto &&
            direction != OfficeTextDirection.LeftToRight &&
            direction != OfficeTextDirection.RightToLeft) {
            throw new ArgumentOutOfRangeException(nameof(direction), "Text shaping direction must be Auto, LeftToRight, or RightToLeft.");
        }
        if (fontCollectionIndex < 0) {
            throw new ArgumentOutOfRangeException(nameof(fontCollectionIndex), "Font collection indexes cannot be negative.");
        }

        FontName = fontName ?? string.Empty;
        _fontData = cloneFontData ? (byte[])fontData.Clone() : fontData;
        _fontProgramCacheKey = fontProgramCacheKey;
        IsOpenTypeCff = isOpenTypeCff;
        UnitsPerEm = unitsPerEm;
        FontCollectionIndex = fontCollectionIndex;
        Direction = direction;
        Language = string.IsNullOrWhiteSpace(language) ? null : language;
        CancellationToken = cancellationToken;
        _variationCoordinates = SnapshotVariationCoordinates(variationCoordinates);
        _featureSettings = featureSettings ?? OfficeTextFeatureSettings.Default;
    }

    /// <summary>Original UTF-16 text to shape.</summary>
    public string Text { get; }

    /// <summary>Display or resource name of the selected font.</summary>
    public string FontName { get; }

    /// <summary>Defensive snapshot of the complete font program.</summary>
    public byte[] FontData => (byte[])_fontData.Clone();

    internal byte[] FontDataForShaping => _fontData;

    internal object? FontProgramCacheKeyForShaping => _fontProgramCacheKey;

    /// <summary>True for OpenType/CFF outlines; false for TrueType outlines.</summary>
    public bool IsOpenTypeCff { get; }

    /// <summary>Font design units per em used by glyph advances and offsets.</summary>
    public int UnitsPerEm { get; }

    /// <summary>Selected zero-based face when <see cref="FontData"/> contains a TrueType collection.</summary>
    public int? FontCollectionIndex { get; }

    /// <summary>Resolved base direction of the run.</summary>
    public OfficeTextDirection Direction { get; }

    /// <summary>Optional BCP 47 language hint.</summary>
    public string? Language { get; }

    /// <summary>Cancellation observed by cooperative host shapers.</summary>
    public System.Threading.CancellationToken CancellationToken { get; }

    /// <summary>Clamped design-space values for the selected variable-font instance.</summary>
    public IReadOnlyDictionary<string, float> VariationCoordinates => _variationCoordinates;

    internal IReadOnlyDictionary<string, float> VariationCoordinatesForShaping => _variationCoordinates;

    /// <summary>Explicit OpenType feature values for this run.</summary>
    public OfficeTextFeatureSettings FeatureSettings => _featureSettings;

    private static IReadOnlyDictionary<string, float> SnapshotVariationCoordinates(
        IReadOnlyDictionary<string, float>? coordinates) {
        var snapshot = new Dictionary<string, float>(StringComparer.Ordinal);
        if (coordinates != null) {
            if (coordinates.Count > 64) {
                throw new ArgumentException("Text shaping requests support at most 64 variable-font axes.", nameof(coordinates));
            }
            foreach (KeyValuePair<string, float> coordinate in coordinates) {
                if (coordinate.Key == null || coordinate.Key.Length != 4) {
                    throw new ArgumentException("Variable-font axis tags must contain exactly four characters.", nameof(coordinates));
                }
                for (int index = 0; index < coordinate.Key.Length; index++) {
                    if (coordinate.Key[index] < 0x20 || coordinate.Key[index] > 0x7E) {
                        throw new ArgumentException("Variable-font axis tags must contain printable ASCII characters.", nameof(coordinates));
                    }
                }
                if (float.IsNaN(coordinate.Value) || float.IsInfinity(coordinate.Value)) {
                    throw new ArgumentOutOfRangeException(nameof(coordinates), "Variable-font axis values must be finite.");
                }
                snapshot.Add(coordinate.Key, coordinate.Value);
            }
        }
        return new ReadOnlyDictionary<string, float>(snapshot);
    }
}

/// <summary>A shaped glyph run in visual write order.</summary>
public sealed class OfficeTextShapingResult {
    private readonly int[]? _advanceAdjustments;

    /// <summary>Creates an immutable result from shaped glyph mappings.</summary>
    public OfficeTextShapingResult(IEnumerable<OfficeShapedGlyph> glyphs)
        : this(glyphs, advanceAdjustments: null) {
    }

    internal OfficeTextShapingResult(
        IEnumerable<OfficeShapedGlyph> glyphs,
        IReadOnlyList<int>? advanceAdjustments) {
        if (glyphs == null) {
            throw new ArgumentNullException(nameof(glyphs));
        }

        var snapshot = new List<OfficeShapedGlyph>();
        foreach (OfficeShapedGlyph glyph in glyphs) {
            snapshot.Add(glyph);
        }
        Glyphs = Array.AsReadOnly(snapshot.ToArray());
        if (advanceAdjustments != null) {
            if (advanceAdjustments.Count != snapshot.Count) {
                throw new ArgumentException("Shaped glyph advance adjustments must match the glyph count.", nameof(advanceAdjustments));
            }
            _advanceAdjustments = new int[advanceAdjustments.Count];
            for (int index = 0; index < advanceAdjustments.Count; index++) {
                _advanceAdjustments[index] = advanceAdjustments[index];
            }
        }
    }

    /// <summary>Glyph identifiers, source mappings, advances, and offsets.</summary>
    public IReadOnlyList<OfficeShapedGlyph> Glyphs { get; }

    internal int GetAdvanceAdjustment(int glyphIndex) => _advanceAdjustments?[glyphIndex] ?? 0;
}

/// <summary>Maps one shaped font glyph back to its logical Unicode source.</summary>
public readonly struct OfficeShapedGlyph {
    /// <summary>Creates a glyph that uses the font's nominal advance.</summary>
    public OfficeShapedGlyph(int glyphId, string unicodeText, int textIndex)
        : this(glyphId, unicodeText, textIndex, advanceWidth: null, offsetX: 0, offsetY: 0) {
    }

    /// <summary>Creates a positioned glyph expressed in the request font's design units.</summary>
    public OfficeShapedGlyph(
        int glyphId,
        string unicodeText,
        int textIndex,
        int advanceWidth,
        int offsetX = 0,
        int offsetY = 0)
        : this(glyphId, unicodeText, textIndex, (int?)advanceWidth, offsetX, offsetY) {
    }

    private OfficeShapedGlyph(
        int glyphId,
        string unicodeText,
        int textIndex,
        int? advanceWidth,
        int offsetX,
        int offsetY) {
        if (glyphId <= 0) {
            throw new ArgumentOutOfRangeException(nameof(glyphId), "Shaped glyph identifiers must be positive.");
        }
        if (string.IsNullOrEmpty(unicodeText)) {
            throw new ArgumentException("Shaped glyphs must preserve non-empty Unicode extraction text.", nameof(unicodeText));
        }
        if (textIndex < 0) {
            throw new ArgumentOutOfRangeException(nameof(textIndex), "Shaped glyph text indexes cannot be negative.");
        }

        GlyphId = glyphId;
        UnicodeText = unicodeText;
        TextIndex = textIndex;
        AdvanceWidth = advanceWidth;
        OffsetX = offsetX;
        OffsetY = offsetY;
    }

    internal static OfficeShapedGlyph CreatePositionedUsingNominalAdvance(
        int glyphId,
        string unicodeText,
        int textIndex,
        int offsetX,
        int offsetY) => new OfficeShapedGlyph(glyphId, unicodeText, textIndex, advanceWidth: null, offsetX, offsetY);

    /// <summary>Font glyph identifier.</summary>
    public int GlyphId { get; }

    /// <summary>Logical Unicode text represented by this glyph.</summary>
    public string UnicodeText { get; }

    /// <summary>UTF-16 source index where <see cref="UnicodeText"/> begins.</summary>
    public int TextIndex { get; }

    /// <summary>Optional shaped advance in font design units; null uses the nominal glyph width.</summary>
    public int? AdvanceWidth { get; }

    /// <summary>Horizontal placement offset in font design units.</summary>
    public int OffsetX { get; }

    /// <summary>Vertical placement offset in font design units.</summary>
    public int OffsetY { get; }
}
