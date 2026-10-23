using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeManagedTextShapingProvider {
    // Default Latin consumers can project the canonical substituted tokens directly. These
    // tokens include continuation ownership and logical clusters, not just nominal cmap glyphs.
    // Default Latin shaping also applies the font's default pair kerning (GPOS kern feature,
    // with the legacy kern table as fallback). Positions are null when no pair adjusts the run.
    internal bool TryShapeDefaultLatinTokens(OfficeTextShapingRequest request,
        out List<OfficeOpenTypeSubstitution.GlyphToken> tokens, out OfficeTextDirection direction,
        out OfficeOpenTypeGlyphPositioning[]? positions) {
        if (request == null) throw new ArgumentNullException(nameof(request));
        tokens = null!;
        positions = null;
        direction = request.Direction;
        if (!request.ApplyDefaultLatinLigatures || !request.FeatureSettings.IsDefault ||
            !TryShapeTokens(request, out IOfficeFontProgram font, out tokens, out direction)) return false;
        positions = PositionDefaultTokens(font, tokens, request.CancellationToken);
        return true;
    }

    // Width consumers can visit the same substituted and kerned tokens without allocating public
    // shaped glyphs and clusters. scaleAdjustment converts font-unit kerning into the consumer's
    // units with the same rounding its drawing path uses.
    internal bool TryMeasureDefaultLatinText(OfficeTextShapingRequest request, Func<int, string, int> measureGlyph,
        Func<int, int> scaleAdjustment, out int advance) {
        if (request == null) throw new ArgumentNullException(nameof(request));
        if (measureGlyph == null) throw new ArgumentNullException(nameof(measureGlyph));
        if (scaleAdjustment == null) throw new ArgumentNullException(nameof(scaleAdjustment));
        advance = 0;
        if (!TryShapeDefaultLatinTokens(request, out List<OfficeOpenTypeSubstitution.GlyphToken> tokens, out _,
            out OfficeOpenTypeGlyphPositioning[]? positions)) return false;
        for (int index = 0; index < tokens.Count; index++) {
            request.CancellationToken.ThrowIfCancellationRequested();
            OfficeOpenTypeSubstitution.GlyphToken token = tokens[index];
            advance = checked(advance + measureGlyph(token.GlyphId, token.UnicodeText));
            if (positions != null && positions[index].XAdvance != 0) advance = checked(advance + scaleAdjustment(positions[index].XAdvance));
        }
        return true;
    }

    private static OfficeOpenTypeGlyphPositioning[]? PositionDefaultTokens(IOfficeFontProgram font,
        List<OfficeOpenTypeSubstitution.GlyphToken> tokens, System.Threading.CancellationToken cancellationToken) {
        if (tokens.Count < 2) return null;
        var glyphIds = new int[tokens.Count];
        var scalars = new int[tokens.Count];
        for (int index = 0; index < tokens.Count; index++) {
            glyphIds[index] = tokens[index].GlyphId;
            scalars[index] = tokens[index].Scalar;
        }
        OfficeOpenTypeGlyphPositioning[] positions = PositionGlyphRun(font, glyphIds, scalars, cancellationToken);
        foreach (OfficeOpenTypeGlyphPositioning position in positions) {
            if (position.XAdvance != 0 || position.XPlacement != 0) return positions;
        }
        return null;
    }

    private bool TryShapeTokens(OfficeTextShapingRequest request, out IOfficeFontProgram font,
        out List<OfficeOpenTypeSubstitution.GlyphToken> tokens, out OfficeTextDirection resolvedDirection,
        bool allowDiagnosedFallback = false) {
        if (request == null) throw new ArgumentNullException(nameof(request));
        font = null!;
        tokens = null!;
        resolvedDirection = request.Direction;
        request.CancellationToken.ThrowIfCancellationRequested();
        if (request.Direction == OfficeTextDirection.TopToBottom ||
            string.IsNullOrEmpty(request.Text) ||
            !OfficeManagedTextShaper.RequiresComplexLayout(request.Text) && request.FeatureSettings.IsDefault && !request.ApplyDefaultLatinLigatures ||
            // Default Latin GSUB and its diagnostics own supported composition. The other
            // managed path cannot provide mark attachment or required Arabic ligatures.
            !allowDiagnosedFallback && !request.ApplyDefaultLatinLigatures &&
            (OfficeTextElements.ContainsCombiningMark(request.Text) ||
             OfficeArabicTextShaper.RequiresLamAlefLigatureShaping(request.Text, request.FeatureSettings)) ||
            !request.ApplyDefaultLatinLigatures && (OfficeTextElements.ContainsVariationSelector(request.Text) ||
            OfficeTextElements.ContainsZeroWidthJoinerSequence(request.Text) ||
            OfficeTextElements.ContainsShapingRequiredScript(request.Text) ||
            (OfficeTextElements.ContainsJoiningScript(request.Text) &&
             !OfficeArabicTextShaper.CanShapeAllJoiningCharacters(request.Text)))) return false;

        LatinFont? latinFont = ResolveLatinFont(request);
        IOfficeFontProgram? loadedFont = latinFont != null ? latinFont.Font : LoadShapingFont(request);
        if (loadedFont == null) return false;
        font = loadedFont;
        resolvedDirection = request.Direction == OfficeTextDirection.Auto
            ? OfficeTextElements.ResolveBaseDirection(request.Text) : request.Direction;
        OfficeOpenTypeSubstitution? substitution = latinFont?.Substitution ?? OfficeOpenTypeSubstitution.TryCreate(request.FontDataForShaping);
        bool ascii = request.ApplyDefaultLatinLigatures &&
            (resolvedDirection == OfficeTextDirection.LeftToRight || resolvedDirection == OfficeTextDirection.Auto) &&
            IsPrintableAscii(request.Text, request.CancellationToken);
        if (ascii) {
            // The existing default lookup preflight can decline this font. Do it before
            // glyph mapping when a Latin letter establishes eligibility. Numeric-only
            // text remains eligible for nominal shaping even with an unsupported GSUB.
            bool hasLatin = false;
            for (int index = 0; index < request.Text.Length; index++) {
                if ((index & 4095) == 0) request.CancellationToken.ThrowIfCancellationRequested();
                char character = request.Text[index];
                hasLatin |= character >= 'A' && character <= 'Z' || character >= 'a' && character <= 'z';
            }
            if (hasLatin && request.FeatureSettings.IsDefault && substitution != null &&
                !substitution.CanApplyLatinDefaults(request.FeatureSettings, request.CancellationToken)) return false;
            tokens = new List<OfficeOpenTypeSubstitution.GlyphToken>(request.Text.Length);
            for (int index = 0; index < request.Text.Length; index++) {
                request.CancellationToken.ThrowIfCancellationRequested();
                char character = request.Text[index];
                if (!font.TryGetGlyphMetrics(character, out int glyphId, out _)) return false;
                tokens.Add(new OfficeOpenTypeSubstitution.GlyphToken(glyphId, AsciiCharacters[character - 32], index, character));
            }
        } else {
            // Direct default-Latin and diagnosed callers reach this owner too.
            // Only complete typography can retain a proven blank control glyph.
            if (OfficeManagedTextShaper.ContainsNonPaintingControl(request.Text, request.CancellationToken)) return false;
            string contextual = request.ApplyDefaultLatinLigatures ? request.Text : OfficeArabicTextShaper.Shape(request.Text);
            IReadOnlyList<VisualTextElement> visualElements = MapVisualElements(request.Text, contextual, resolvedDirection,
                request.CancellationToken, reorder: !request.ApplyDefaultLatinLigatures || request.Direction != OfficeTextDirection.Auto);
            if (visualElements.Count == 0) return false;
            string visual = string.Concat(visualElements.Select(static element => element.VisualText));
            if (!font.HasGlyphs(visual)) return false;
            tokens = new List<OfficeOpenTypeSubstitution.GlyphToken>(visualElements.Count);
            foreach (VisualTextElement element in visualElements) {
                request.CancellationToken.ThrowIfCancellationRequested();
                if (!TryAddElementGlyphs(font, element, tokens)) return false;
            }
        }

        if (request.ApplyDefaultLatinLigatures) {
            if (substitution != null && !substitution.ApplyLatinDefaults(tokens, request.FeatureSettings, request.CancellationToken, request.Text)) return false;
        } else {
            if (substitution != null && !substitution.CanApply(request.FeatureSettings, request.CancellationToken)) return false;
            substitution?.Apply(tokens, request.FeatureSettings, request.CancellationToken);
        }
        return tokens.Count > 0;
    }
}
