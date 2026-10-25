namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    // Bar declarations share the paragraph stop snapshot, but never enter
    // text-stop selection, leader generation or font/line-box measurement.
    private static PdfTabStop[]? SnapshotBarTabStops(IReadOnlyList<PdfTabStop>? stops, double frameWidth) {
        if (stops == null) return null;
        PdfTabStop[] bars = stops.Where(stop => stop.Alignment == PdfTabAlignment.Bar)
            .Select(stop => stop.ResolveFramePosition(frameWidth)).ToArray();
        return bars.Length == 0 ? null : bars;
    }

    /// <summary>Retains the paragraph anchor and logical interval before physical line slicing.</summary>
    private static void BindRichLineBarTabs(RichLine line, PdfTabStop[]? stops, double lineOrigin, double height) {
        if (stops == null) return;
        line.BarTabStops = stops;
        // The painter receives a line origin, which may include a hanging
        // indent or floating-object exclusion. Rules keep the paragraph origin.
        line.BarOriginOffset = -lineOrigin;
        line.BarLineHeight = height;
    }

    /// <summary>
    /// Paints stationary paragraph rules over each final logical line interval.
    /// Only the first and last logical lines receive their own actually applied
    /// paragraph before/after spacing. Float clearance is excluded. Cell layout
    /// retains both extents separately from flattened inter-paragraph heights.
    /// </summary>
    private static void AppendRichBarTabDecorations(StringBuilder sb, List<List<RichSeg>> lines,
        List<double> lineHeights, PdfOptions options, double startY, double fontSize, double leading,
        double xOrigin, double? firstLineX, IReadOnlyList<double>? lineXOffsets, IReadOnlyList<double>? lineTopGaps,
        PdfStandardFont? baselineFont, double paragraphSpacingAfter, double paragraphSpacingBefore, bool tagged, string? graphicsState,
        TableCellContentFrame? clipFrame) {
        if (!lines.Any(line => line is RichLine { BarTabStops: not null })) return;
        var content = new ContentStreamBuilder(sb).SaveState();
        if (graphicsState != null) content.GraphicsState(graphicsState);
        // Glyph clipping deliberately allows font ink to bleed outside some
        // cell frames. Stationary rules own no such allowance, including caps.
        if (clipFrame is { } clip)
            content.Rectangle(clip.Left, clip.Top - clip.Height, clip.Width, clip.Height).ClipPath().EndPath();
        content.LineCap(2).LineJoin(1).StrokeDash(Array.Empty<double>());
        AppendArtifactBegin(sb, tagged);
        double yOffset = 0D;
        double baseAscent = GetAscenderForOptions(baselineFont ?? ChooseNormal(options.DefaultFont), fontSize, options);
        for (int index = 0; index < lines.Count; index++) {
            if (lines[index] is RichLine { BarTabStops: { } stops } line) {
                double top = startY - yOffset - (lineTopGaps?[index] ?? 0D) + baseAscent;
                double before = line.IsFirstParagraphLine ? line.BarSpacingBefore + paragraphSpacingBefore : 0D;
                double after = line.IsLastParagraphLine ? line.BarSpacingAfter + paragraphSpacingAfter : 0D;
                double bottom = top - line.BarLineHeight - after;
                // The text top already follows consumed before-spacing. Extend
                // only the first rule back to the paragraph's flow boundary.
                top += before;
                if (clipFrame is { } frame) {
                    top = Math.Min(top, frame.Top);
                    bottom = Math.Max(bottom, frame.Top - frame.Height);
                }
                double origin = ResolveRichLineXOrigin(xOrigin, firstLineX, lineXOffsets, index) + line.BarOriginOffset;
                if (top > bottom) {
                    foreach (PdfTabStop stop in stops) {
                        double x = origin + stop.Position;
                        content.StrokeColor(stop.BarColor).LineWidth(stop.BarWidth)
                            .MoveTo(x, top).LineTo(x, bottom).StrokePath();
                    }
                }
            }
            yOffset += index < lineHeights.Count ? lineHeights[index] : leading;
        }
        AppendArtifactEnd(sb, tagged);
        content.RestoreState();
    }
}
