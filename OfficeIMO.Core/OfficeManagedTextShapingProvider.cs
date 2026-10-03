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
public sealed partial class OfficeManagedTextShapingProvider : IOfficeTextShapingProvider, IOfficeTextShapingProviderMetadata {
    private static readonly System.Runtime.CompilerServices.ConditionalWeakTable<byte[], LatinFont> LatinFonts = new();
    /// <summary>Shared stateless provider instance.</summary>
    public static OfficeManagedTextShapingProvider Instance { get; } = new OfficeManagedTextShapingProvider();

    private OfficeManagedTextShapingProvider() {
    }

    /// <inheritdoc />
    public OfficeTextShapingBackend Backend => OfficeTextShapingBackend.Managed;

    /// <inheritdoc />
    public OfficeTextShapingResult? ShapeText(OfficeTextShapingRequest request) {
        if (!TryShapeTokens(request, out IOfficeFontProgram font, out List<OfficeOpenTypeSubstitution.GlyphToken> tokens,
            out OfficeTextDirection resolvedDirection)) return null;
        bool kerningEnabled = request.FeatureSettings.TryGetValue("kern", out int kerningValue)
            ? kerningValue != 0 : !request.ApplyDefaultLatinLigatures;
        var glyphIds = new int[tokens.Count];
        var scalars = new int[tokens.Count];
        for (int index = 0; index < tokens.Count; index++) {
            glyphIds[index] = tokens[index].GlyphId;
            scalars[index] = tokens[index].Scalar;
        }
        OfficeOpenTypeGlyphPositioning[] positioning = kerningEnabled
            ? PositionGlyphRun(font, glyphIds, scalars, request.CancellationToken)
            : new OfficeOpenTypeGlyphPositioning[tokens.Count];
        var glyphs = new List<OfficeShapedGlyph>(tokens.Count);
        var advanceAdjustments = new List<int>(tokens.Count);
        for (int index = 0; index < tokens.Count; index++) {
            OfficeOpenTypeSubstitution.GlyphToken token = tokens[index];
            request.CancellationToken.ThrowIfCancellationRequested();
            glyphs.Add(token.IsUnicodeContinuation
                ? OfficeShapedGlyph.CreateUnicodeContinuation(
                    token.GlyphId,
                    token.TextIndex,
                    positioning[index].XPlacement,
                    offsetY: 0)
                : OfficeShapedGlyph.CreatePositionedUsingNominalAdvance(
                    token.GlyphId,
                    token.UnicodeText,
                    token.TextIndex,
                    positioning[index].XPlacement,
                    offsetY: 0));
            advanceAdjustments.Add(positioning[index].XAdvance);
        }

        int[]? clusterStarts = tokens.Any(token => token.ClusterStart != token.TextIndex)
            ? tokens.Select(token => token.ClusterStart).ToArray() : null;
        return glyphs.Count == 0 ? null : new OfficeTextShapingResult(glyphs, advanceAdjustments, resolvedDirection, clusterStarts);
    }

    private static IReadOnlyList<VisualTextElement> MapVisualElements(
        string logical,
        string contextual,
        OfficeTextDirection direction,
        System.Threading.CancellationToken cancellationToken,
        bool reorder = true) {
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

        if (!reorder) return logicalElements;
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
        IReadOnlyList<int> scalars,
        System.Threading.CancellationToken cancellationToken) {
        if (font is OfficeTrueTypeFont trueType) return trueType.PositionGlyphRun(glyphIds, scalars, cancellationToken);
        if (font is OfficeOpenTypeCffFont cff) return cff.PositionGlyphRun(glyphIds, scalars, cancellationToken);
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

    private sealed class LatinFont {
        internal LatinFont(byte[] data, bool isCff) {
            Substitution = OfficeOpenTypeSubstitution.TryCreate(data);
            Font = isCff ? OfficeOpenTypeCffFont.TryLoad(data, null, out _) : OfficeTrueTypeFont.TryLoad(data);
        }
        internal IOfficeFontProgram? Font { get; }
        internal OfficeOpenTypeSubstitution? Substitution { get; }
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
