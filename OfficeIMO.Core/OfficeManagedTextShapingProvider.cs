using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Drawing;

/// <summary>
/// Dependency-light shaping provider for bounded Arabic joining, bidirectional text, and common
/// OpenType substitutions that can be represented by a TrueType-outline font.
/// </summary>
/// <remarks>
/// The provider deliberately declines scripts and lookup types that require shaping beyond the
/// bounded managed core. Callers then retain their normal scalar fallback and diagnostics. This
/// keeps <see cref="IOfficeTextShapingProvider"/> as the single shaping contract used by Drawing and PDF.
/// </remarks>
public sealed class OfficeManagedTextShapingProvider : IOfficeTextShapingProvider {
    /// <summary>Shared stateless provider instance.</summary>
    public static OfficeManagedTextShapingProvider Instance { get; } = new OfficeManagedTextShapingProvider();

    private OfficeManagedTextShapingProvider() {
    }

    /// <inheritdoc />
    public OfficeTextShapingResult? ShapeText(OfficeTextShapingRequest request) {
        if (request == null) throw new ArgumentNullException(nameof(request));
        request.CancellationToken.ThrowIfCancellationRequested();
        if (string.IsNullOrEmpty(request.Text) ||
            !OfficeManagedTextShaper.RequiresComplexLayout(request.Text) && request.FeatureSettings.IsDefault ||
            OfficeTextElements.ContainsVariationSelector(request.Text) ||
            OfficeTextElements.ContainsZeroWidthJoinerSequence(request.Text) ||
            OfficeTextElements.ContainsShapingRequiredScript(request.Text) ||
            (OfficeTextElements.ContainsJoiningScript(request.Text) &&
             !OfficeArabicTextShaper.CanShapeAllJoiningCharacters(request.Text))) {
            return null;
        }

        IOfficeFontProgram? font = request.IsOpenTypeCff
            ? OfficeOpenTypeCffFont.TryLoad(request.FontDataForShaping, request.VariationCoordinatesForShaping, out _)
            : OfficeTrueTypeFont.TryLoad(request.FontDataForShaping, request.FontCollectionIndex);
        if (font == null) return null;

        string contextual = OfficeArabicTextShaper.Shape(request.Text);
        IReadOnlyList<VisualTextElement> visualElements = MapVisualElements(
            request.Text,
            contextual,
            request.Direction,
            request.CancellationToken);
        if (visualElements.Count == 0) return null;
        string visual = string.Concat(visualElements.Select(static element => element.VisualText));
        if (!font.HasGlyphs(visual)) return null;
        var tokens = new List<OfficeOpenTypeSubstitution.GlyphToken>();
        foreach (VisualTextElement element in visualElements) {
            request.CancellationToken.ThrowIfCancellationRequested();
            if (!TryAddElementGlyphs(font, element, tokens)) return null;
        }

        OfficeOpenTypeSubstitution? substitution = OfficeOpenTypeSubstitution.TryCreate(request.FontDataForShaping);
        substitution?.Apply(tokens, request.FeatureSettings, request.CancellationToken);
        bool kerningEnabled = !request.FeatureSettings.TryGetValue("kern", out int kerningValue) || kerningValue != 0;
        var glyphIds = new int[tokens.Count];
        var scalars = new int[tokens.Count];
        for (int index = 0; index < tokens.Count; index++) {
            glyphIds[index] = tokens[index].GlyphId;
            scalars[index] = tokens[index].Scalar;
        }
        OfficeOpenTypeGlyphPositioning[] positioning = kerningEnabled
            ? PositionGlyphRun(font, glyphIds, scalars)
            : new OfficeOpenTypeGlyphPositioning[tokens.Count];
        var glyphs = new List<OfficeShapedGlyph>(tokens.Count);
        var advanceAdjustments = new List<int>(tokens.Count);
        foreach ((OfficeOpenTypeSubstitution.GlyphToken token, int index) in tokens.Select((token, index) => (token, index))) {
            request.CancellationToken.ThrowIfCancellationRequested();
            glyphs.Add(OfficeShapedGlyph.CreatePositionedUsingNominalAdvance(
                token.GlyphId,
                token.UnicodeText,
                token.TextIndex,
                positioning[index].XPlacement,
                offsetY: 0));
            advanceAdjustments.Add(positioning[index].XAdvance);
        }

        return glyphs.Count == 0 ? null : new OfficeTextShapingResult(glyphs, advanceAdjustments);
    }

    private static IReadOnlyList<VisualTextElement> MapVisualElements(
        string logical,
        string contextual,
        OfficeTextDirection direction,
        System.Threading.CancellationToken cancellationToken) {
        var logicalElements = new List<VisualTextElement>();
        int logicalIndex = 0;
        foreach (string contextualElement in OfficeTextElements.Enumerate(contextual)) {
            cancellationToken.ThrowIfCancellationRequested();
            int length = contextualElement.Length;
            string logicalElement = logical.Substring(logicalIndex, Math.Min(length, logical.Length - logicalIndex));
            if (!IsBidiControlElement(logicalElement)) {
                logicalElements.Add(new VisualTextElement(contextualElement, logicalElement, logicalIndex));
            }
            logicalIndex += length;
        }

        return OfficeBidiTextResolver.ToVisualOrder(
            contextual,
            logicalElements,
            direction,
            cancellationToken,
            static element => element.WithVisualText(OfficeBidiTextResolver.MirrorText(element.VisualText)));
    }

    private static bool TryAddElementGlyphs(
        IOfficeFontProgram font,
        VisualTextElement element,
        List<OfficeOpenTypeSubstitution.GlyphToken> glyphs) {
        int visualIndex = 0;
        int logicalOffset = 0;
        while (visualIndex < element.VisualText.Length) {
            int visualScalar = ReadScalar(element.VisualText, ref visualIndex);
            int logicalStart = logicalOffset;
            int logicalScalar = ReadScalar(element.LogicalText, ref logicalOffset);
            if (!font.TryGetGlyphMetrics(visualScalar, out int glyphId, out _)) {
                return false;
            }

            string unicodeText = char.ConvertFromUtf32(logicalScalar);
            glyphs.Add(new OfficeOpenTypeSubstitution.GlyphToken(
                glyphId,
                unicodeText,
                element.LogicalIndex + logicalStart,
                logicalScalar));
        }

        return true;
    }

    private static OfficeOpenTypeGlyphPositioning[] PositionGlyphRun(
        IOfficeFontProgram font,
        IReadOnlyList<int> glyphIds,
        IReadOnlyList<int> scalars) {
        if (font is OfficeTrueTypeFont trueType) return trueType.PositionGlyphRun(glyphIds, scalars);
        if (font is OfficeOpenTypeCffFont cff) return cff.PositionGlyphRun(glyphIds, scalars);
        return new OfficeOpenTypeGlyphPositioning[glyphIds.Count];
    }

    private static bool IsBidiControlElement(string value) =>
        value.Length > 0 && OfficeTextElements.ContainsBidiControl(value);

    private static int ReadScalar(string text, ref int index) {
        char first = text[index++];
        return char.IsHighSurrogate(first) &&
               index < text.Length &&
               char.IsLowSurrogate(text[index])
            ? char.ConvertToUtf32(first, text[index++])
            : first;
    }

    private readonly struct VisualTextElement {
        internal VisualTextElement(string visualText, string logicalText, int logicalIndex) {
            VisualText = visualText;
            LogicalText = logicalText;
            LogicalIndex = logicalIndex;
        }

        internal string VisualText { get; }
        internal string LogicalText { get; }
        internal int LogicalIndex { get; }

        internal VisualTextElement WithVisualText(string visualText) =>
            new VisualTextElement(visualText, LogicalText, LogicalIndex);
    }
}
