using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeManagedTextShapingProvider {
    // Default Latin consumers can project the canonical substituted tokens directly. These
    // tokens include continuation ownership and logical clusters, not just nominal cmap glyphs.
    internal bool TryShapeDefaultLatinTokens(OfficeTextShapingRequest request,
        out List<OfficeOpenTypeSubstitution.GlyphToken> tokens, out OfficeTextDirection direction) {
        if (request == null) throw new ArgumentNullException(nameof(request));
        tokens = null!;
        direction = request.Direction;
        return request.ApplyDefaultLatinLigatures && request.FeatureSettings.IsDefault &&
            TryShapeTokens(request, out _, out tokens, out direction);
    }

    // Default Latin shaping uses nominal advances and no positioning. Width consumers can
    // visit the same substituted tokens without allocating public shaped glyphs and clusters.
    // Explicit features and providers retain the complete shaping-result contract.
    internal bool TryMeasureDefaultLatinText(OfficeTextShapingRequest request, Func<int, string, int> measureGlyph, out int advance) {
        if (request == null) throw new ArgumentNullException(nameof(request));
        if (measureGlyph == null) throw new ArgumentNullException(nameof(measureGlyph));
        advance = 0;
        if (!TryShapeDefaultLatinTokens(request, out List<OfficeOpenTypeSubstitution.GlyphToken> tokens, out _)) return false;
        foreach (OfficeOpenTypeSubstitution.GlyphToken token in tokens) {
            request.CancellationToken.ThrowIfCancellationRequested();
            advance = checked(advance + measureGlyph(token.GlyphId, token.UnicodeText));
        }
        return true;
    }

    private bool TryShapeTokens(OfficeTextShapingRequest request, out IOfficeFontProgram font,
        out List<OfficeOpenTypeSubstitution.GlyphToken> tokens, out OfficeTextDirection resolvedDirection) {
        if (request == null) throw new ArgumentNullException(nameof(request));
        font = null!;
        tokens = null!;
        resolvedDirection = request.Direction;
        request.CancellationToken.ThrowIfCancellationRequested();
        if (request.Direction == OfficeTextDirection.TopToBottom ||
            string.IsNullOrEmpty(request.Text) ||
            !OfficeManagedTextShaper.RequiresComplexLayout(request.Text) && request.FeatureSettings.IsDefault && !request.ApplyDefaultLatinLigatures ||
            !request.ApplyDefaultLatinLigatures && (OfficeTextElements.ContainsVariationSelector(request.Text) ||
            OfficeTextElements.ContainsZeroWidthJoinerSequence(request.Text) ||
            OfficeTextElements.ContainsShapingRequiredScript(request.Text) ||
            (OfficeTextElements.ContainsJoiningScript(request.Text) &&
             !OfficeArabicTextShaper.CanShapeAllJoiningCharacters(request.Text)))) return false;

        LatinFont? latinFont = request.ApplyDefaultLatinLigatures
            ? LatinFonts.GetValue(request.FontDataForShaping, data => new LatinFont(data, request.IsOpenTypeCff)) : null;
        IOfficeFontProgram? loadedFont = latinFont?.Font ?? (request.IsOpenTypeCff
            ? OfficeOpenTypeCffFont.TryLoad(request.FontDataForShaping, request.VariationCoordinatesForShaping, out _)
            : OfficeTrueTypeFont.TryLoad(request.FontDataForShaping, request.FontCollectionIndex));
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
