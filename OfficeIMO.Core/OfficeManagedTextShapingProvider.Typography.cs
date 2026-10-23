using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeManagedTextShapingProvider {
    private static readonly ConditionalWeakTable<object, TypographyFont> TypographyFonts = new();

    // This path keeps nominal glyphs in logical order until the font has completed its
    // substitutions and attachments. Reversing scalars first loses ligature components.
    private OfficeTextShapingResult? ShapeTypography(OfficeTextShapingRequest request, out bool fontDirected) {
        fontDirected = false;
        request.CancellationToken.ThrowIfCancellationRequested();
        if (request.Direction == OfficeTextDirection.TopToBottom || request.Text.Length == 0 ||
            OfficeTextElements.ContainsVariationSelector(request.Text) ||
            OfficeTextElements.ContainsShapingRequiredScript(request.Text) ||
            !OfficeArabicTextShaper.CanShapeAllJoiningCharacters(request.Text) ||
            request.VariationCoordinatesForShaping.Count != 0) return null;
        try {
            object key = request.FontProgramCacheKeyForShaping ?? request.FontDataForShaping;
            TypographyFont cached = TypographyFonts.GetValue(key, _ => new TypographyFont(request));
            fontDirected = cached.Substitution != null && OfficeTextElements.ContainsJoiningScript(request.Text);
            IOfficeFontProgram? font = cached.Font;
            OfficeOpenTypeReader? reader = cached.Reader;
            // Variable anchors, face extraction and variation substitution have their own
            // provider contract; static success must not silently discard those inputs.
            if (font == null || reader == null || reader.TryGetTable("fvar", out _, out _)) return null;
            int[] scalars = OfficeUnicodeBidi.ToScalars(request.Text, out int[] starts);
            int? baseLevel = request.Direction == OfficeTextDirection.RightToLeft ? 1
                : request.Direction == OfficeTextDirection.LeftToRight ? 0 : (int?)null;
            OfficeUnicodeBidi.Paragraph paragraph = OfficeUnicodeBidi.Resolve(scalars, baseLevel, request.CancellationToken);
            byte[] levels = paragraph.GetLineLevels(0, scalars.Length);
            if (!TryMapTypographyTokens(request, font, reader, scalars, starts, levels,
                out List<OfficeOpenTypeSubstitution.GlyphToken> tokens)) return null;
            if (tokens.Count == 0 || cached.Substitution != null && HasReversedLatinTypography(tokens, starts, levels, request)) return null;
            if (!ApplyTypographySubstitutions(cached.Substitution, font, tokens, request, starts, levels)) return null;
            if (!OfficeOpenTypeMarkPositioning.TryGetAttachments(request.FontDataForShaping, tokens,
                request.Language, request.FeatureSettings, request.CancellationToken, out OfficeOpenTypeMarkAttachment[] attachments)) return null;
            List<int> visual = ResolveTypographyVisualOrder(request, tokens, attachments, starts, levels, out int[] logicalClusters);
            OfficeOpenTypeGlyphPositioning[] positions;
            if (request.FeatureSettings.TryGetValue("kern", out int kern) && kern == 0)
                positions = new OfficeOpenTypeGlyphPositioning[tokens.Count];
            else if (cached.Substitution == null && !attachments.Any(item => item.IsMark)) {
                // Preserve the established simple-font visual positioning route;
                // source controls now remain explicit zero-advance glyphs.
                if (!TryPositionVisualTypography(font, tokens, visual, request, out positions)) return null;
            } else if (!TryPositionTypography(font, tokens, request, out positions)) return null;
            return CompleteTypography(request, reader, tokens, attachments, positions, logicalClusters, visual,
                (paragraph.ParagraphLevel & 1) == 1 ? OfficeTextDirection.RightToLeft : OfficeTextDirection.LeftToRight);
        } catch (Exception exception) when (exception is InvalidDataException or OverflowException or
            ArgumentOutOfRangeException or IndexOutOfRangeException) { return null; }
    }

    private static bool ApplyTypographySubstitutions(OfficeOpenTypeSubstitution? substitution,
        IOfficeFontProgram font, List<OfficeOpenTypeSubstitution.GlyphToken> tokens, OfficeTextShapingRequest request,
        int[] scalarStarts, byte[] levels) {
        if (substitution == null) return TryApplySimpleTypographyJoining(font, tokens, request);
        OfficeOpenTypeSubstitution.ScriptShapingContext context = substitution.CreateScriptShapingContext(
            request.Text, request.FeatureSettings, request.Language, request.CancellationToken);
        var output = new List<OfficeOpenTypeSubstitution.GlyphToken>(tokens.Count);
        var scripts = new string[tokens.Count];
        string inherited = string.Empty;
        for (int index = 0; index < tokens.Count; index++) {
            context.Charge();
            int scalar = tokens[index].Scalar;
            string script = OfficeOpenTypeGlyphFilter.IsUnicodeMark(scalar) ? string.Empty
                : OfficeTextElements.IsJoiningScriptScalar(scalar) ? "arab"
                : scalar is >= 0x0590 and <= 0x05FF ? "hebr"
                : scalar is >= 0x0041 and <= 0x024F or >= 0x1E00 and <= 0x1EFF ? "latn" : string.Empty;
            if (script.Length > 0) inherited = script;
            scripts[index] = inherited;
        }
        string following = "latn";
        for (int index = tokens.Count - 1; index >= 0; index--) {
            context.Charge();
            if (scripts[index].Length == 0) scripts[index] = following;
            else following = scripts[index];
        }
        for (int start = 0; start < tokens.Count;) {
            context.Charge();
            string script = scripts[start];
            byte level = TypographyLevel(tokens[start].TextIndex, scalarStarts, levels);
            int end = start + 1;
            while (end < tokens.Count && scripts[end] == script
                && TypographyLevel(tokens[end].TextIndex, scalarStarts, levels) == level) { context.Charge(); end++; }
            var segment = tokens.GetRange(start, end - start);
            if (script == "arab") {
                if (!substitution.TryApplyArabicDefaults(segment, request.FeatureSettings, request.Language,
                    request.CancellationToken, request.Text, context)) return false;
            } else if (!substitution.TryApplyMarkScriptDefaults(segment, request.FeatureSettings, script,
                request.Language, request.CancellationToken, request.Text, context)) return false;
            output.AddRange(segment);
            start = end;
        }
        tokens.Clear(); tokens.AddRange(output);
        return true;
    }

    private static bool TryPositionTypography(IOfficeFontProgram font,
        IReadOnlyList<OfficeOpenTypeSubstitution.GlyphToken> tokens, OfficeTextShapingRequest request,
        out OfficeOpenTypeGlyphPositioning[] positions) {
        if (font is OfficeTrueTypeFont trueType)
            return trueType.TryPositionGlyphRun(tokens, request.Language, request.CancellationToken, out positions, rejectMarkAdjustments: true);
        if (font is OfficeOpenTypeCffFont cff)
            return cff.TryPositionGlyphRun(tokens, request.Language, request.CancellationToken, out positions, rejectMarkAdjustments: true);
        positions = Array.Empty<OfficeOpenTypeGlyphPositioning>();
        return false;
    }

    // Resolve attachment origins in logical dependency order, after allocating the visual
    // pens. Anchor deltas have no bidi-dependent sign and never consume additional advance.
    private static OfficeTextShapingResult CompleteTypography(OfficeTextShapingRequest request,
        OfficeOpenTypeReader reader, List<OfficeOpenTypeSubstitution.GlyphToken> tokens,
        OfficeOpenTypeMarkAttachment[] attachments, OfficeOpenTypeGlyphPositioning[] positions,
        int[] logicalClusters, List<int> visual, OfficeTextDirection direction) {
        int[] advances = new int[tokens.Count], pen = new int[tokens.Count], penY = new int[tokens.Count];
        int cursor = 0, cursorY = 0;
        foreach (int index in visual) {
            request.CancellationToken.ThrowIfCancellationRequested();
            if ((attachments[index].IsMark || OfficeManagedTextShaper.IsNonPaintingControl(tokens[index].Scalar))
                && (positions[index].XPlacement != 0 || positions[index].YPlacement != 0
                    || positions[index].XAdvance != 0 || positions[index].YAdvance != 0))
                throw new InvalidDataException("A zero-advance mark cannot discard font positioning.");
            pen[index] = cursor; penY[index] = cursorY;
            advances[index] = attachments[index].IsMark || OfficeManagedTextShaper.IsNonPaintingControl(tokens[index].Scalar) ? 0
                : checked(reader.AdvanceWidth(tokens[index].GlyphId) + positions[index].XAdvance);
            cursor = checked(cursor + advances[index]);
            cursorY = checked(cursorY + (attachments[index].IsMark || OfficeManagedTextShaper.IsNonPaintingControl(tokens[index].Scalar)
                ? 0 : positions[index].YAdvance));
        }
        var originX = new int[tokens.Count]; var originY = new int[tokens.Count];
        for (int index = 0; index < tokens.Count; index++) {
            request.CancellationToken.ThrowIfCancellationRequested();
            OfficeOpenTypeMarkAttachment attachment = attachments[index];
            int parent = attachment.ParentGlyphIndex;
            if (parent >= index) throw new InvalidDataException("A mark attachment must refer to an earlier logical glyph.");
            originX[index] = parent < 0 ? checked(pen[index] + positions[index].XPlacement)
                : checked(originX[parent] + attachment.DeltaX);
            originY[index] = parent < 0 ? checked(penY[index] + positions[index].YPlacement)
                : checked(originY[parent] + attachment.DeltaY);
        }
        var glyphs = new List<OfficeShapedGlyph>(tokens.Count); var clusters = new int[tokens.Count];
        foreach (int index in visual) {
            OfficeOpenTypeSubstitution.GlyphToken token = tokens[index];
            int offsetX = checked(originX[index] - pen[index]);
            int offsetY = checked(originY[index] - penY[index]);
            int advanceY = attachments[index].IsMark || OfficeManagedTextShaper.IsNonPaintingControl(token.Scalar)
                ? 0 : positions[index].YAdvance;
            clusters[glyphs.Count] = logicalClusters[index];
            glyphs.Add(token.IsUnicodeContinuation
                ? OfficeShapedGlyph.CreatePositionedUnicodeContinuation(token.GlyphId, token.TextIndex,
                    advances[index], advanceY, offsetX, offsetY)
                : new OfficeShapedGlyph(token.GlyphId, token.UnicodeText, token.TextIndex,
                    advances[index], advanceY, offsetX, offsetY));
        }
        return new OfficeTextShapingResult(glyphs, advanceAdjustments: null, direction, clusters);
    }

    private sealed class TypographyFont {
        internal TypographyFont(OfficeTextShapingRequest request) {
            Reader = OfficeOpenTypeReader.TryCreate(request.FontDataForShaping);
            Font = LoadShapingFont(request);
            Substitution = OfficeOpenTypeSubstitution.TryCreate(request.FontDataForShaping);
        }
        internal OfficeOpenTypeReader? Reader { get; }
        internal IOfficeFontProgram? Font { get; }
        internal OfficeOpenTypeSubstitution? Substitution { get; }
    }
}
