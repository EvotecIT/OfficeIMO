using System.Runtime.CompilerServices;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeManagedTextShapingProvider {
    private static readonly ConditionalWeakTable<object, TypographyFont> TypographyFonts = new();

    // This path keeps nominal glyphs in logical order until the font has completed its
    // substitutions and attachments. Reversing scalars first loses ligature components.
    private OfficeTextShapingResult? ShapeTypography(OfficeTextShapingRequest request, out bool fontDirected) {
        fontDirected = false;
        request.CancellationToken.ThrowIfCancellationRequested();
        if (request.Direction == OfficeTextDirection.TopToBottom || request.Text.Length == 0 || request.Text.Length > MaximumNormalizationInput ||
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
            byte[] levels = ResolveTypographyLevels(request, scalars, out int paragraphLevel);
            if (!TryMapTypographyTokens(request, font, reader, scalars, starts, levels,
                out List<OfficeOpenTypeSubstitution.GlyphToken> tokens)) return null;
            if (tokens.Count == 0 || cached.Substitution != null && HasReversedLatinTypography(tokens, starts, levels, request)) return null;
            if (!ApplyTypographySubstitutions(cached.Substitution, font, tokens, request, starts, levels)) return null;
            MergeLigatureClusters(tokens, request.CancellationToken);
            if (!HideTypographyControls(request, font, reader, tokens)) return null;
            if (!OfficeOpenTypeMarkPositioning.TryGetAttachments(request.FontDataForShaping, tokens,
                request.Language, request.FeatureSettings, request.CancellationToken, out OfficeOpenTypeMarkAttachment[] attachments)) return null;
            List<int> visual = ResolveTypographyVisualOrder(request, tokens, attachments, starts, levels, out int[] logicalClusters);
            OfficeOpenTypeGlyphPositioning[] positions;
            bool kerning = !request.FeatureSettings.TryGetValue("kern", out int kern) || kern != 0;
            bool cursive = !request.FeatureSettings.TryGetValue("curs", out int curs) || curs != 0;
            bool distances = !request.FeatureSettings.TryGetValue("dist", out int dist) || dist != 0;
            var rightToLeft = new bool[tokens.Count];
            for (int index = 0; index < tokens.Count; index++)
                rightToLeft[index] = (TypographyLevel(tokens[index].TextIndex, starts, levels) & 1) == 1;
            var options = new OfficeOpenTypePositioningOptions(kerning, cursive, rightToLeft, distances, attachments);
            if (cached.Substitution == null && !attachments.Any(item => item.IsMark) && kerning) {
                // Preserve the established simple-font visual positioning route;
                // source controls now remain explicit zero-advance glyphs.
                if (!TryPositionVisualTypography(font, tokens, visual, request, out positions)) return null;
                // That route reads pair kerning only. A font with cursive attachment or
                // distances takes the logical GPOS route, which owns curs and dist alongside kern.
                if ((cursive && cached.HasCursiveFeature || distances && cached.HasDistancesFeature)
                    && !TryPositionTypography(font, tokens, request, out positions, options)) return null;
            } else if (!TryPositionTypography(font, tokens, request, out positions, options)) return null;
            return CompleteTypography(request, reader, tokens, attachments, positions, logicalClusters, visual,
                (paragraphLevel & 1) == 1 ? OfficeTextDirection.RightToLeft : OfficeTextDirection.LeftToRight);
        } catch (Exception exception) when (exception is InvalidDataException or OverflowException or
            ArgumentOutOfRangeException or IndexOutOfRangeException) { return null; }
    }

    private static bool RequiresCursiveAttachment(OfficeTextShapingRequest request) {
        if (request.FeatureSettings.TryGetValue("curs", out int curs) && curs == 0) return false;
        object key = request.FontProgramCacheKeyForShaping ?? request.FontDataForShaping;
        TypographyFont cached = TypographyFonts.GetValue(key, _ => new TypographyFont(request));
        return cached.HasCursiveFeature;
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

    // A ligature that consumed a later grapheme's base owns that grapheme's cluster, so
    // the grapheme's remaining marks join the ligature's cluster, as reference shaping
    // merges clusters over a ligature's source span.
    private static void MergeLigatureClusters(List<OfficeOpenTypeSubstitution.GlyphToken> tokens,
        CancellationToken cancellationToken) {
        int ownerCluster = -1, ownerEnd = -1;
        for (int index = 0; index < tokens.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            OfficeOpenTypeSubstitution.GlyphToken token = tokens[index];
            if (ownerCluster >= 0 && token.ClusterStart > ownerCluster && token.ClusterStart < ownerEnd)
                tokens[index] = token = token.WithClusterStart(ownerCluster);
            int end = token.TextIndex + token.UnicodeText.Length;
            if (token.ClusterStart != ownerCluster) { ownerCluster = token.ClusterStart; ownerEnd = end; }
            else ownerEnd = Math.Max(ownerEnd, end);
        }
    }

    private static bool TryPositionTypography(IOfficeFontProgram font,
        IReadOnlyList<OfficeOpenTypeSubstitution.GlyphToken> tokens, OfficeTextShapingRequest request,
        out OfficeOpenTypeGlyphPositioning[] positions, OfficeOpenTypePositioningOptions? options = null) {
        if (font is OfficeTrueTypeFont trueType)
            return trueType.TryPositionGlyphRun(tokens, request.Language, request.CancellationToken, out positions, rejectMarkAdjustments: true, options);
        if (font is OfficeOpenTypeCffFont cff)
            return cff.TryPositionGlyphRun(tokens, request.Language, request.CancellationToken, out positions, rejectMarkAdjustments: true, options);
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
            // Positioning admits only placement on marks, from lookups after attachment.
            bool control = OfficeManagedTextShaper.IsNonPaintingControl(tokens[index].Scalar);
            if ((attachments[index].IsMark || control)
                && (control && (positions[index].XPlacement != 0 || positions[index].YPlacement != 0)
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
                : checked(originX[parent] + attachment.DeltaX + positions[index].XPlacement);
            originY[index] = parent < 0 ? checked(penY[index] + positions[index].YPlacement)
                : checked(originY[parent] + attachment.DeltaY + positions[index].YPlacement);
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
            HasCursiveFeature = Reader != null && OfficeOpenTypeLayoutFeatures.HasFeature(Reader, "GPOS", "curs");
            HasDistancesFeature = Reader != null && OfficeOpenTypeLayoutFeatures.HasFeature(Reader, "GPOS", "dist");
        }
        internal OfficeOpenTypeReader? Reader { get; }
        internal bool HasCursiveFeature { get; }
        internal bool HasDistancesFeature { get; }
        internal IOfficeFontProgram? Font { get; }
        internal OfficeOpenTypeSubstitution? Substitution { get; }
    }
}
