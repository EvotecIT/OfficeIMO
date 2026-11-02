using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    /// <summary>Uses shared cluster/visual ownership for generic and imported Word character spacing.</summary>
    private static RichLineJustification PlanRichCharacterDistribution(List<RichSeg> segments,
        double naturalWidth, double availableWidth, PdfCreateOptions options, bool useWordJustificationSpacing) {
        var text = new StringBuilder();
        var indexes = new List<int>();
        var owners = new List<(int Segment, int Glyph)>();
        var commands = new PdfTextShowCommand?[segments.Count];
        for (int index = 0; index < segments.Count; index++) {
            RichSeg segment = segments[index];
            if (segment.LeadingSpace && !segment.LeadingIsTab) {
                for (int space = 0; space < segment.LeadingSpaceCount; space++) {
                    indexes.Add(text.Length);
                    owners.Add((index, -1));
                    text.Append(' ');
                }
            }
            if (segment.InlineElement != null) {
                indexes.Add(text.Length);
                owners.Add((index, -2));
                text.Append('\uFFFC');
                continue;
            }
            if (segment.Text.Length == 0) continue;
            PdfTextShowCommand ownership = segment.KerningCommand is { VisualGlyphs: not null } prepared ? prepared :
                EncodeTextShowCommand(segment.Text, segment.Font, segment.NamedFont, options,
                    segment.FeatureSettings, segment.TextDirection, segment.FontMetricScale,
                    segment.HorizontalTextScaling, segment.PreservesPaintedGlyphs, requireGlyphClusters: true);
            PdfTextShowCommand command = segment.KerningCommand ?? ownership;
            if (command.GlyphCount != ownership.GlyphCount || command.GlyphHex != ownership.GlyphHex)
                throw new InvalidOperationException("Character distribution must retain the prepared encoded glyphs.");
            commands[index] = command;
            int start = text.Length;
            // Encoded glyph ownership is relative to painted text, without optional breaks.
            string paintedText = WithoutOptionalBreaks(segment.Text);
            text.Append(paintedText);
            if (ownership.VisualGlyphs is { } glyphs) {
                int[] sourceIndexes = ResolveRichGlyphClusterStarts(glyphs);
                for (int glyph = 0; glyph < glyphs.Count; glyph++) {
                    indexes.Add(start + sourceIndexes[glyph]);
                    owners.Add((index, glyph));
                }
            } else {
                // The only cluster-free encoding here is a standard one-byte WinAnsi font.
                // Embedded encoders were explicitly requested to return shaped ownership.
                if (command.GlyphCount != paintedText.Length)
                    throw new InvalidOperationException("Character distribution requires encoded glyph source ownership.");
                for (int glyph = 0; glyph < command.GlyphCount; glyph++) {
                    indexes.Add(start + glyph);
                    owners.Add((index, glyph));
                }
            }
        }
        if (indexes.Count < 2) return new RichLineJustification(segments, 0D, naturalWidth);
        bool[] boundaries = OfficeOpenTypeTracking.GetBoundaries(text.ToString(), indexes);
        // Source-only controls can follow the final visible cluster. Character
        // distribution expands between visible clusters, never after the last.
        for (int index = boundaries.Length - 1; index >= 0; index--) {
            if (!boundaries[index]) continue;
            boundaries[index] = false;
            break;
        }
        int opportunities = boundaries.Count(boundary => boundary);
        if (opportunities == 0) return new RichLineJustification(segments, 0D, naturalWidth);
        if (useWordJustificationSpacing)
            return PlanRichWordCharacterDistribution(segments, naturalWidth, availableWidth, options,
                boundaries, owners, commands, opportunities);
        double spacing = (availableWidth - naturalWidth) / opportunities;
        var leading = new int[segments.Count];
        var inline = new int[segments.Count];
        var glyphBoundaries = new bool[segments.Count][];
        for (int boundary = 0; boundary < boundaries.Length; boundary++) {
            if (!boundaries[boundary]) continue;
            (int segment, int glyph) = owners[boundary];
            if (glyph == -1) leading[segment]++;
            else if (glyph == -2) inline[segment]++;
            else {
                glyphBoundaries[segment] ??= new bool[commands[segment]!.GlyphCount];
                glyphBoundaries[segment][glyph] = true;
            }
        }
        var distributed = new List<RichSeg>(segments.Count);
        for (int index = 0; index < segments.Count; index++) {
            RichSeg segment = segments[index];
            bool[]? selected = glyphBoundaries[index];
            int count = (selected?.Count(boundary => boundary) ?? 0) + inline[index];
            distributed.Add(count == 0 && leading[index] == 0 ? segment : segment.WithJustification(
                count * spacing, segment.LeadingAdvance + leading[index] * spacing,
                selected == null ? null : commands[index]!.WithJustificationAdvances(Array.ConvertAll(selected, boundary => boundary ? spacing : 0D))));
        }
        return new RichLineJustification(distributed, 0D, availableWidth);
    }

    /// <summary>Consumes ordinary-space capacity, then source-average capacity, before equal residual expansion.</summary>
    private static RichLineJustification PlanRichWordCharacterDistribution(List<RichSeg> segments,
        double naturalWidth, double availableWidth, PdfCreateOptions options, bool[] boundaries,
        List<(int Segment, int Glyph)> owners, PdfTextShowCommand?[] commands, int opportunities) {
        double adjustment = availableWidth - naturalWidth;
        var preferredCharacters = new double?[segments.Count];
        for (int index = 0; index < segments.Count; index++) {
            RichSeg segment = segments[index];
            if (segment.InlineElement == null && segment.Text.Length > 0)
                preferredCharacters[index] = GetWordPreferredCharacterCapacity(segment, options);
        }

        var resources = new List<OfficeTextSpacingResource>();
        for (int boundary = 0; boundary < boundaries.Length; boundary++) {
            if (!boundaries[boundary]) continue;
            (int index, int glyph) = owners[boundary];
            RichSeg segment = segments[index];
            if (glyph == -1 && IsWordExpansionSpace(segment)) {
                // A preserved group retains one opportunity per source space.
                // Its aggregate source advance supplies their preferred shares.
                double spaceCapacity = Math.Max(0D, segment.LeadingAdvance / segment.LeadingSpaceCount / 2D);
                resources.Add(new OfficeTextSpacingResource(boundary, 0, 0D, spaceCapacity));
            } else if (glyph >= 0 && preferredCharacters[index] is double characterCapacity) {
                resources.Add(new OfficeTextSpacingResource(boundary, 2, 0D, characterCapacity));
            }
            // Objects and unavailable source metrics retain their existing
            // boundary opportunity without inventing a preferred font average.
            resources.Add(new OfficeTextSpacingResource(boundary, 4, 0D, adjustment / opportunities));
        }
        OfficeTextSpacingAllocation allocation = OfficeTextSpacingAllocator.Allocate(boundaries.Length, resources, adjustment);
        var leading = new double[segments.Count];
        var glyphTotals = new double[segments.Count];
        var glyphAdvances = new double[segments.Count][];
        for (int boundary = 0; boundary < boundaries.Length; boundary++) {
            double delta = allocation.Deltas[boundary];
            if (delta == 0D) continue;
            (int index, int glyph) = owners[boundary];
            if (glyph == -1) leading[index] += delta;
            else {
                glyphTotals[index] += delta;
                if (glyph < 0) continue;
                glyphAdvances[index] ??= new double[commands[index]!.GlyphCount];
                glyphAdvances[index][glyph] += delta;
            }
        }
        var distributed = new List<RichSeg>(segments.Count);
        for (int index = 0; index < segments.Count; index++) {
            RichSeg segment = segments[index];
            double[]? advances = glyphAdvances[index];
            distributed.Add(glyphTotals[index] == 0D && leading[index] == 0D ? segment : segment.WithJustification(
                glyphTotals[index], segment.LeadingAdvance + leading[index],
                advances == null ? null : commands[index]!.WithJustificationAdvances(advances)));
        }
        return new RichLineJustification(distributed, 0D, naturalWidth + allocation.AppliedAdjustment);
    }

    /// <summary>Reads the resolved source face's optional average at the segment's effective page size.</summary>
    private static double? GetWordPreferredCharacterCapacity(RichSeg segment, PdfCreateOptions options) {
        OfficeOpenTypeAverageAdvanceMetrics? average = null;
        if (segment.NamedFont is PdfNamedFontFace named) {
            if (options.TryGetNamedFontProgram(named, out PdfTrueTypeFontProgram? trueType) && trueType != null)
                average = trueType.AverageAdvanceMetrics;
            else if (options.TryGetNamedOpenTypeCffFontProgram(named, out PdfOpenTypeCffFontProgram? cff) && cff != null)
                average = cff.AverageAdvanceMetrics;
        } else {
            if (options.TryGetEmbeddedStandardFontProgram(segment.Font, out PdfTrueTypeFontProgram? trueType) && trueType != null)
                average = trueType.AverageAdvanceMetrics;
            else if (options.TryGetEmbeddedStandardOpenTypeCffFontProgram(segment.Font, out PdfOpenTypeCffFontProgram? cff) && cff != null)
                average = cff.AverageAdvanceMetrics;
        }
        if (average is not OfficeOpenTypeAverageAdvanceMetrics metric) return null;
        // Keep the source metric continuous in points. Authored tracking and
        // character spacing belong to natural advances; PDF Tf/width rounding
        // does not redefine this optional preferred-character resource.
        return metric.AverageAdvance / (double)metric.UnitsPerEm
            * EffectiveRichFontSize(segment.FontSize, segment.Baseline)
            * (segment.HorizontalTextScaling / 100D) / 32D;
    }
}
