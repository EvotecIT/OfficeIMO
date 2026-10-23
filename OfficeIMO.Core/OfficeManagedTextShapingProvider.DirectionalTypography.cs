using System;
using System.Collections.Generic;
using System.IO;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeManagedTextShapingProvider {
    // Direction is the UAX #9 paragraph base. Only resolved scalar levels tell
    // whether an override actually reverses the font's natural Latin order.
    private static byte TypographyLevel(int sourceIndex, int[] starts, byte[] levels) {
        int scalar = Array.BinarySearch(starts, sourceIndex);
        if (scalar < 0) scalar = Math.Max(0, ~scalar - 1);
        return levels[scalar];
    }

    private static bool HasReversedLatinTypography(IReadOnlyList<OfficeOpenTypeSubstitution.GlyphToken> tokens,
        int[] starts, byte[] levels, OfficeTextShapingRequest request) {
        foreach (OfficeOpenTypeSubstitution.GlyphToken token in tokens) {
            request.CancellationToken.ThrowIfCancellationRequested();
            bool latin = token.Scalar is >= 0x0041 and <= 0x024F or >= 0x1E00 and <= 0x1EFF;
            if (latin && OfficeUnicodeBidi.GetClass(token.Scalar) == OfficeBidiClass.L
                && (TypographyLevel(token.TextIndex, starts, levels) & 1) != 0) return true;
        }
        return false;
    }

    // Fonts without GSUB retain the already supported length-preserving Arabic
    // forms. Marks and required ligatures remain outside this simple profile.
    private static bool TryApplySimpleTypographyJoining(IOfficeFontProgram font,
        List<OfficeOpenTypeSubstitution.GlyphToken> tokens, OfficeTextShapingRequest request) {
        if (!OfficeTextElements.ContainsJoiningScript(request.Text)) return true;
        if (OfficeArabicTextShaper.RequiresLamAlefLigatureShaping(request.Text, request.FeatureSettings)) return false;
        foreach (OfficeOpenTypeSubstitution.GlyphToken token in tokens) {
            request.CancellationToken.ThrowIfCancellationRequested();
            if (OfficeOpenTypeGlyphFilter.IsUnicodeMark(token.Scalar)
                && !OfficeManagedTextShaper.IsNonPaintingControl(token.Scalar)) return false;
        }
        string?[]? joining = OfficeArabicTextShaper.GetOpenTypeJoiningForms(request.Text,
            request.CancellationToken.ThrowIfCancellationRequested);
        if (joining == null) return false;
        for (int index = 0; index < tokens.Count; index++) {
            request.CancellationToken.ThrowIfCancellationRequested();
            OfficeOpenTypeSubstitution.GlyphToken token = tokens[index];
            if (!OfficeArabicTextShaper.TryGetContextualPresentationScalar(token.Scalar, joining[token.TextIndex], out int scalar)) continue;
            if (!font.TryGetGlyphMetrics(scalar, out int glyph, out _)) return false;
            tokens[index] = token.WithGlyph(glyph);
        }
        return true;
    }

    private static bool TryPositionVisualTypography(IOfficeFontProgram font,
        IReadOnlyList<OfficeOpenTypeSubstitution.GlyphToken> tokens, List<int> visual,
        OfficeTextShapingRequest request, out OfficeOpenTypeGlyphPositioning[] positions) {
        var ordered = new OfficeOpenTypeSubstitution.GlyphToken[tokens.Count];
        for (int index = 0; index < ordered.Length; index++) {
            request.CancellationToken.ThrowIfCancellationRequested(); ordered[index] = tokens[visual[index]];
        }
        if (!TryPositionTypography(font, ordered, request, out OfficeOpenTypeGlyphPositioning[] selected)) {
            positions = Array.Empty<OfficeOpenTypeGlyphPositioning>(); return false;
        }
        positions = new OfficeOpenTypeGlyphPositioning[tokens.Count];
        for (int index = 0; index < selected.Length; index++) {
            request.CancellationToken.ThrowIfCancellationRequested(); positions[visual[index]] = selected[index];
        }
        return true;
    }

    private static List<int> ResolveTypographyVisualOrder(OfficeTextShapingRequest request,
        List<OfficeOpenTypeSubstitution.GlyphToken> tokens, OfficeOpenTypeMarkAttachment[] attachments,
        int[] scalarStarts, byte[] levels, out int[] logicalClusters) {
        var groups = new List<List<int>>(); var groupLevels = new List<byte>();
        var byCluster = new Dictionary<int, int>(); logicalClusters = new int[tokens.Count];
        for (int index = 0; index < tokens.Count; index++) {
            request.CancellationToken.ThrowIfCancellationRequested();
            int parent = attachments[index].ParentGlyphIndex;
            if (parent >= index) throw new InvalidDataException("A mark attachment must refer to an earlier logical glyph.");
            int cluster = parent < 0 ? tokens[index].ClusterStart : logicalClusters[parent];
            logicalClusters[index] = cluster;
            if (!byCluster.TryGetValue(cluster, out int group)) {
                group = groups.Count; byCluster.Add(cluster, group); groups.Add(new List<int>());
                groupLevels.Add(TypographyLevel(cluster, scalarStarts, levels));
            }
            groups[group].Add(index);
        }
        int[] order = OfficeUnicodeBidi.ReorderLine(groupLevels.ToArray());
        var visual = new List<int>(tokens.Count);
        foreach (int group in order) for (int offset = 0; offset < groups[group].Count; offset++) {
            request.CancellationToken.ThrowIfCancellationRequested();
            visual.Add(groups[group][(groupLevels[group] & 1) == 1 ? groups[group].Count - offset - 1 : offset]);
        }
        return visual;
    }
}
