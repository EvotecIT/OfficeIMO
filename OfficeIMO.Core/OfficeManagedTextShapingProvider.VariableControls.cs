using System;
using System.Collections.Generic;
using System.IO;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeManagedTextShapingProvider {
    // Preserve the existing selected-instance nominal Latin contract. This path
    // applies neither automatic GSUB nor variable mark positioning, and never
    // caches an instance by its source bytes. Visible advances stay nominal so
    // the consuming selected font remains the HVAR/phantom-point metrics owner.
    private static OfficeTextShapingResult? ShapeNominalVariableControls(OfficeTextShapingRequest request) {
        if (request.IsOpenTypeCff || request.Direction == OfficeTextDirection.TopToBottom
            || !request.FeatureSettings.IsDefault || request.ApplyDefaultLatinLigatures
            || request.Text.Length == 0 || request.Text.Length > MaximumNormalizationInput) return null;
        try {
            int[] scalars = OfficeUnicodeBidi.ToScalars(request.Text, out int[] starts);
            foreach (int scalar in scalars) {
                request.CancellationToken.ThrowIfCancellationRequested();
                if (scalar is >= 0x20 and <= 0x7E) continue;
                if (scalar is 0x034F or 0x200C or 0x200D || !OfficeManagedTextShaper.IsNonPaintingControl(scalar)) return null;
            }
            OfficeOpenTypeReader? reader = OfficeOpenTypeReader.TryCreate(request.FontDataForShaping);
            if (reader == null || !reader.TryGetTable("fvar", out _, out _)) return null;
            if (reader.TryGetTable("GSUB", out _, out _)) {
                OfficeOpenTypeSubstitution? substitution = OfficeOpenTypeSubstitution.TryCreate(request.FontDataForShaping);
                if (substitution == null || !substitution.CanRetainNominalLatin(request.Text,
                    request.Language, request.CancellationToken)) return null;
            }
            if (LoadShapingFont(request) is not OfficeTrueTypeFont font) return null;
            byte[] levels = ResolveTypographyLevels(request, scalars, out int paragraphLevel);
            if (!TryMapTypographyTokens(request, font, reader, scalars, starts, levels, out var tokens)) return null;
            var attachments = new OfficeOpenTypeMarkAttachment[tokens.Count];
            for (int index = 0; index < attachments.Length; index++)
                attachments[index] = new OfficeOpenTypeMarkAttachment(-1, 0, 0, false);
            List<int> visual = ResolveTypographyVisualOrder(request, tokens, attachments, starts, levels, out int[] logicalClusters);
            var glyphIds = new int[tokens.Count]; var visualScalars = new int[tokens.Count];
            for (int index = 0; index < visual.Count; index++) {
                request.CancellationToken.ThrowIfCancellationRequested();
                glyphIds[index] = tokens[visual[index]].GlyphId;
                visualScalars[index] = tokens[visual[index]].Scalar;
            }
            OfficeOpenTypeGlyphPositioning[] positions = PositionGlyphRun(font, glyphIds, visualScalars, request.CancellationToken);
            var glyphs = new List<OfficeShapedGlyph>(tokens.Count); var adjustments = new List<int>(tokens.Count);
            var clusters = new int[tokens.Count];
            for (int index = 0; index < visual.Count; index++) {
                request.CancellationToken.ThrowIfCancellationRequested();
                OfficeOpenTypeSubstitution.GlyphToken token = tokens[visual[index]];
                OfficeOpenTypeGlyphPositioning position = positions[index];
                if (OfficeManagedTextShaper.IsNonPaintingControl(token.Scalar)) {
                    if (position.XPlacement != 0 || position.XAdvance != 0
                        || position.YPlacement != 0 || position.YAdvance != 0) return null;
                    glyphs.Add(token.IsUnicodeContinuation
                        ? OfficeShapedGlyph.CreatePositionedUnicodeContinuation(token.GlyphId, token.TextIndex, 0, 0, 0, 0)
                        : new OfficeShapedGlyph(token.GlyphId, token.UnicodeText, token.TextIndex, 0, 0, 0, 0));
                } else {
                    glyphs.Add(token.IsUnicodeContinuation
                        ? OfficeShapedGlyph.CreateUnicodeContinuation(token.GlyphId, token.TextIndex, position.XPlacement, 0)
                        : OfficeShapedGlyph.CreatePositionedUsingNominalAdvance(token.GlyphId, token.UnicodeText,
                            token.TextIndex, position.XPlacement, 0));
                }
                adjustments.Add(position.XAdvance);
                clusters[index] = logicalClusters[visual[index]];
            }
            return new OfficeTextShapingResult(glyphs, adjustments,
                (paragraphLevel & 1) == 1 ? OfficeTextDirection.RightToLeft : OfficeTextDirection.LeftToRight, clusters);
        } catch (Exception exception) when (exception is InvalidDataException or OverflowException
            or ArgumentOutOfRangeException or IndexOutOfRangeException) { return null; }
    }
}
