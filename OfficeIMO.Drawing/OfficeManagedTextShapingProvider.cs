using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Drawing;

/// <summary>
/// Dependency-light shaping provider for bounded Arabic joining and bidirectional text that can be
/// represented by a TrueType-outline font.
/// </summary>
/// <remarks>
/// The provider deliberately declines OpenType/CFF fonts and scripts that require GSUB/GPOS shaping
/// beyond the managed core. Callers then retain their normal scalar fallback and diagnostics. This
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
        if (request.IsOpenTypeCff ||
            string.IsNullOrEmpty(request.Text) ||
            !OfficeManagedTextShaper.RequiresComplexLayout(request.Text) ||
            OfficeTextElements.ContainsBidiControl(request.Text) ||
            OfficeTextElements.ContainsShapingRequiredScript(request.Text) ||
            (OfficeTextElements.ContainsJoiningScript(request.Text) &&
             !OfficeArabicTextShaper.CanShapeAllJoiningCharacters(request.Text))) {
            return null;
        }

        OfficeTrueTypeFont? font = OfficeTrueTypeFont.TryLoad(request.FontData, request.FontCollectionIndex);
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
        var glyphs = new List<OfficeShapedGlyph>();
        foreach (VisualTextElement element in visualElements) {
            request.CancellationToken.ThrowIfCancellationRequested();
            if (!TryAddElementGlyphs(font, element, glyphs)) return null;
        }

        return glyphs.Count == 0 ? null : new OfficeTextShapingResult(glyphs);
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

        return OfficeManagedTextShaper.ToVisualOrder(
            logicalElements,
            static element => element.VisualText,
            direction,
            cancellationToken);
    }

    private static bool TryAddElementGlyphs(
        OfficeTrueTypeFont font,
        VisualTextElement element,
        List<OfficeShapedGlyph> glyphs) {
        int visualIndex = 0;
        int logicalOffset = 0;
        while (visualIndex < element.VisualText.Length) {
            int visualScalar = ReadScalar(element.VisualText, ref visualIndex);
            int logicalStart = logicalOffset;
            int logicalScalar = ReadScalar(element.LogicalText, ref logicalOffset);
            if (!font.TryGetGlyphMetrics(visualScalar, out int glyphId, out int advanceWidth)) {
                return false;
            }

            string unicodeText = char.ConvertFromUtf32(logicalScalar);
            glyphs.Add(new OfficeShapedGlyph(
                glyphId,
                unicodeText,
                element.LogicalIndex + logicalStart,
                advanceWidth));
        }

        return true;
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
    }
}
