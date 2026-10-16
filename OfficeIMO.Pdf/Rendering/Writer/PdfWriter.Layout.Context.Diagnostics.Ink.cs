using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private readonly List<OfficeTextInkClip> _diagnosticCanvasClips = new();
        private readonly List<bool> _diagnosticCanvasSourceClips = new();
        private readonly Dictionary<byte[], IOfficeBoundedFontProgram> _diagnosticFonts = new();
        private bool _diagnosticPaintHidden;
        private PdfOptions? _diagnosticTextOptions;
        private PdfOptions? _diagnosticTextSource;

        // Never register speculative leader/omitted glyphs in the output document's subset.
        private PdfOptions DiagnosticTextOptions {
            get {
                if (_diagnosticTextOptions == null || !ReferenceEquals(_diagnosticTextSource, currentOpts) || !_diagnosticTextOptions.HasSameMeasurementFonts(currentOpts)) {
                    _diagnosticTextOptions = currentOpts.CreateTextMeasurementSnapshot();
                    _diagnosticTextSource = currentOpts;
                }
                return _diagnosticTextOptions;
            }
        }

        /// <summary>Inspects final composed paint without changing text operators or line admission.</summary>
        private void ReportRichPaintDiagnostics(string code, string source, RichParagraphBlock block,
            List<List<RichSeg>> lines, IReadOnlyList<double> heights, double firstBaseline,
            double fontSize, double leading, double textX, double textWidth,
            TableCellContentFrame frame, bool clipsToFrame = true,
            IReadOnlyList<PdfAlign?>? alignments = null, IReadOnlyList<double>? offsets = null,
            IReadOnlyList<double>? widths = null, PdfStandardFont? baselineFont = null,
            OfficeTransform? localTransform = null, TableCellContentFrame? secondClip = null,
            double? firstLineX = null, double? firstLineWidth = null, List<double>? topGaps = null) {
            if (!currentOpts.HasDiagnosticsReport || _diagnosticPaintHidden) return;
            OfficeTransform transform = (localTransform ?? OfficeTransform.Identity).Then(_canvasEffectToPage);
            if (!transform.TryInvert(out _)) return; // Singular paint has no filled area.
            var sourceClips = new List<OfficeTextInkClip>();
            var targetClips = new List<OfficeTextInkClip>();
            for (int index = 0; index < _diagnosticCanvasClips.Count; index++) {
                (_diagnosticCanvasSourceClips[index] ? sourceClips : targetClips).Add(_diagnosticCanvasClips[index]);
            }
            OfficeTextInkClip pageClip = DiagnosticRectangleClip(new TableCellContentFrame(0D, currentOpts.PageHeight,
                currentOpts.PageWidth, currentOpts.PageHeight), OfficeTransform.Identity);
            bool clipped = false, outsidePage = false, overflow = false, outerClipped = false;
            string? unmeasuredReason = null;
            double baseline = firstBaseline;
            int? bidiLevel = ResolveRichParagraphBidiLevel(lines);
            try {
                for (int index = 0; index < lines.Count; index++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    List<RichSeg> line = lines[index];
                    double lineY = AdjustRichLineBaseline(baseline - (topGaps?[index] ?? 0D), line, currentOpts, fontSize, baselineFont);
                    double available = ResolveRichLineWidth(textWidth, firstLineWidth, widths, index);
                    double origin = ResolveRichLineXOrigin(textX, firstLineX, offsets, index);
                    PdfAlign alignment = ResolveRichLineAlignment(block.Align, alignments, index);
                    double advance = MeasureRichLineWidth(line);
                    int gaps = line.Where(s => s.LeadingSpace && s.LeadingSpaceIsExpandable).Sum(s => s.LeadingSpaceCount);
                    bool justify = alignment == PdfAlign.Justify && index != lines.Count - 1 &&
                        !line.Any(s => s.EndsWithHardBreak) && gaps > 0 && available > advance;
                    double wordSpacing = justify ? (available - advance) / gaps : 0D;
                    double dx = alignment == PdfAlign.Center ? Math.Max(0D, (available - advance) / 2D)
                        : alignment == PdfAlign.Right ? Math.Max(0D, available - advance) : 0D;
                    BidiLineLayout? bidi = bidiLevel.HasValue ? CreateBidiLineLayout(line, bidiLevel.Value, wordSpacing) : null;
                    double cursor = dx;
                    // Paragraph frames may differ on the first line or alongside floats.
                    TableCellContentFrame effectiveFrame = clipsToFrame ? frame
                        : new TableCellContentFrame(origin, frame.Top, available, frame.Height);
                    OfficeTextInkClip frameClip = DiagnosticRectangleClip(effectiveFrame, transform);
                    foreach (var part in line.Select((segment, partIndex) => (segment, partIndex))) {
                        RichSeg segment = part.segment;
                        double gap = segment.LeadingSpace ? segment.LeadingAdvance +
                            (segment.LeadingSpaceIsExpandable ? wordSpacing * segment.LeadingSpaceCount : 0D) : 0D;
                        double gapX = origin + (bidi == null ? cursor : dx + bidi.GapX[part.partIndex]);
                        cursor += gap;
                        if (bidi != null) cursor = dx + bidi.VisualX[part.partIndex];
                        double x = origin + cursor;
                        double size = EffectiveRichFontSize(segment.FontSize, segment.Baseline);
                        double runY = lineY + TextRiseForBaseline(segment.FontSize, segment.Baseline) + segment.LineAlignmentRise;
                        double width = GetRichSegmentWidth(segment);
                        if (segment.BackgroundColor.HasValue && (width > 0D || gap > 0D)) {
                            double px = Math.Max(1.4D, size * .14D), py = Math.Max(.45D, size * .05D);
                            double ascent = GetAscenderForOptions(segment.Font, segment.NamedFont, size, currentOpts);
                            double descent = GetDescenderForOptions(segment.Font, segment.NamedFont, size, currentOpts);
                            Paint(DiagnosticRectangle(x - (bidi == null ? gap : 0D) - px, runY - descent - py,
                                width + (bidi == null ? gap : 0D) + 2D * px, ascent + descent + 2D * py));
                            if (bidi != null && gap > 0D) Paint(DiagnosticRectangle(gapX, runY - descent - py, gap, ascent + descent + 2D * py));
                        }
                        if (segment.LeadingSpace && segment.LeadingUnderlineStyle != OfficeTextDecorationStyle.None &&
                            segment.LeadingTabLeader == PdfTabLeaderStyle.None && segment.LeadingSpaceIsExpandable)
                            Decoration(gapX, gap, lineY + segment.LeadingDecorationTextRise - segment.LeadingDecorationFontSize * .15D,
                                segment.LeadingUnderlineStyle);
                        if (segment.LeadingTabLeader != PdfTabLeaderStyle.None) {
                            string leader = BuildTabLeaderText(gap, segment, DiagnosticTextOptions);
                            foreach (var glyph in DiagnosticGlyphContours(segment, leader, gapX, runY)) Paint(glyph.Contours, nominalGlyphBounds: glyph.NominalBounds);
                        }
                        if (segment.InlineElement is PdfInlineBox box) {
                            double bottom = lineY + box.BaselineOffset;
                            if (box.Background.HasValue) Paint(DiagnosticRectangle(x, bottom, box.Width, box.Height));
                            if (box.BorderColor.HasValue && box.BorderWidth > 0D) {
                                double half = box.BorderWidth / 2D;
                                var ring = DiagnosticRectangle(x - half, bottom - half, box.Width + box.BorderWidth, box.Height + box.BorderWidth);
                                if (box.Width > box.BorderWidth && box.Height > box.BorderWidth) {
                                    var hole = DiagnosticRectangle(x + half, bottom + half, box.Width - box.BorderWidth, box.Height - box.BorderWidth)[0];
                                    hole.Reverse(); ring.Add(hole);
                                }
                                Paint(ring);
                            }
                        } else if (segment.InlineElement is PdfInlineImage image) {
                            double imageX = image.TableTextRotation > 0 ? origin + available - cursor - image.Width : x;
                            foreach (var paint in DiagnosticInlineImagePaint(image, imageX, lineY)) Paint(paint.Contours, paint.Clips);
                        } else if (segment.InlineElement == null && !segment.IsLineBoxStrut) {
                            foreach (var glyph in DiagnosticGlyphContours(segment, segment.Text, x, runY)) Paint(glyph.Contours, nominalGlyphBounds: glyph.NominalBounds);
                            if (segment.Underline) {
                                if (segment.UnderlineStyle == OfficeTextDecorationStyle.Words)
                                    VisitWordDecorationAdvances(segment.Text,
                                        span => MeasurePositionedTextWidth(span, segment.Font, segment.NamedFont, segment.FontSize, segment.Baseline,
                                            DiagnosticTextOptions, segment.FeatureSettings, segment.TextDirection, segment.HorizontalTextScaling, segment.CharacterSpacing, segment.FontMetricScale),
                                        (start, end) => Decoration(x + start, end - start, runY - size * .15D, OfficeTextDecorationStyle.Single));
                                else Decoration(x, width, runY - size * .15D, segment.UnderlineStyle);
                            }
                            if (segment.Strike) Decoration(x, width, runY + size * .32D, segment.StrikeStyle);
                        }
                        cursor += width;
                    }
                    baseline -= GetRichLineHeight(heights, index, leading);

                    void Decoration(double x, double width, double y, OfficeTextDecorationStyle style) {
                        if (width <= 0D || style == OfficeTextDecorationStyle.None) return;
                        // Use the writer's own stroke contours, including patterned gaps and round caps.
                        foreach (var paint in DiagnosticDecorationContours(x, x + width, y, style)) Paint(paint);
                    }
                    void Paint(List<List<OfficePoint>> contours, List<OfficeTextInkClip>? imageClips = null, bool nominalGlyphBounds = false) {
                        foreach (var contour in contours) for (int p = 0; p < contour.Count; p++) contour[p] = transform.TransformPoint(contour[p]);
                        var intrinsic = new List<OfficeTextInkClip>();
                        if (imageClips != null) foreach (var imageClip in imageClips) intrinsic.Add(DiagnosticMapClip(imageClip, transform));
                        // The source already paints through these clips. Only extra target
                        // clipping can remove ink relative to that source appearance. Keep
                        // the reference clips as a prefix for additional-loss measurement.
                        var sourcePaintClips = new List<OfficeTextInkClip>(intrinsic);
                        sourcePaintClips.AddRange(sourceClips);
                        var admittedClips = new List<OfficeTextInkClip>(sourcePaintClips);
                        admittedClips.AddRange(targetClips);
                        var admitted = OfficeRasterCanvas.MeasureNominalFilledContourBounds(contours, OfficeFillRule.NonZero, admittedClips, cancellationToken);
                        RequireMeasured(admitted.IsMeasured);
                        if (targetClips.Count > 0) outerClipped |= Loss(sourcePaintClips, admittedClips);
                        if (!admitted.HasInk) return;
                        var withFrame = new List<OfficeTextInkClip>(admittedClips) { frameClip };
                        if (secondClip.HasValue) withFrame.Add(DiagnosticRectangleClip(secondClip.Value, transform));
                        if (clipsToFrame) clipped |= Loss(admittedClips, withFrame);
                        else overflow |= Loss(admittedClips, withFrame);
                        var paintClips = clipsToFrame ? withFrame : admittedClips;
                        outsidePage |= Loss(paintClips, new List<OfficeTextInkClip>(paintClips) { pageClip });

                        bool Loss(List<OfficeTextInkClip> before, List<OfficeTextInkClip> after) {
                            if (!HasAdditionalLoss(contours, before, after)) return false;
                            if (!nominalGlyphBounds) return true;
                            var initial = OfficeRasterCanvas.MeasureNominalFilledContourBounds(contours, OfficeFillRule.NonZero, before, cancellationToken);
                            var final = OfficeRasterCanvas.MeasureNominalFilledContourBounds(contours, OfficeFillRule.NonZero, after, cancellationToken);
                            RequireMeasured(initial.IsMeasured && final.IsMeasured);
                            // Full rejection proves loss. Partial rejection proves loss from AFM
                            // extrema only when the envelope was intact and clip edges share its axes.
                            if (!initial.IsClipped && (!final.HasInk || after.Skip(before.Count).All(clip => DiagnosticAlignedRectangle(clip, transform)))) return true;
                            unmeasuredReason = "Only nominal standard-font glyph bounds are available; partial path or transformed clipping requires the actual glyph outline.";
                            return false;
                        }
                    }
                }
            } catch (NotSupportedException exception) {
                unmeasuredReason = exception.Message;
            }
            var bounds = transform.TransformRectangleBounds(frame.Left, frame.Top - frame.Height, frame.Width, frame.Height);
            void Report(string suffix, string message, PdfLayoutDiagnosticKind kind, PdfConversionWarningSeverity severity = PdfConversionWarningSeverity.Warning) =>
                ReportAuthoredLayoutDiagnostic(code + suffix, source, message, kind, bounds.Left, bounds.Bottom,
                    bounds.Right - bounds.Left, bounds.Bottom - bounds.Top, severity);
            if (unmeasuredReason != null) currentOpts.AddUnassessedLayoutWarning(code + "PaintBoundsUnmeasured", source, unmeasuredReason);
            if (clipped) Report("PaintClipped", "Paint extends outside the text clip and was clipped or omitted.", PdfLayoutDiagnosticKind.ClippedContent);
            if (outerClipped) Report("CanvasClipPaintClipped", "Text or inline paint was clipped by an enclosing canvas path.", PdfLayoutDiagnosticKind.ClippedContent);
            if (outsidePage) Report("PageBoundsPaintClipped", "Text or inline paint extends outside the physical page.", PdfLayoutDiagnosticKind.ClippedContent);
            if (overflow && !outsidePage) Report("PaintOverflow", "Paint extends outside the paragraph frame and was preserved.", PdfLayoutDiagnosticKind.Overflow, PdfConversionWarningSeverity.Information);
        }

        private static void RequireMeasured(bool measured) {
            if (!measured) throw new NotSupportedException("Authoring paint diagnostics exceeded the bounded contour inspection limit.");
        }

        /// <summary>Measures removed filled ink directly; equal bounding boxes do not establish equal paint.</summary>
        private bool HasAdditionalLoss(List<List<OfficePoint>> paint, List<OfficeTextInkClip> before, List<OfficeTextInkClip> after) {
            var bounds = OfficeRasterCanvas.MeasureNominalAdditionalInkLoss(paint, OfficeFillRule.NonZero, after, before.Count, cancellationToken);
            RequireMeasured(bounds.IsMeasured);
            return bounds.IsClipped;
        }

        private static List<List<OfficePoint>> DiagnosticRectangle(double x, double bottom, double width, double height) =>
            width <= 0D || height <= 0D ? new() : new() { new() {
                new OfficePoint(x, bottom), new OfficePoint(x + width, bottom),
                new OfficePoint(x + width, bottom + height), new OfficePoint(x, bottom + height) } };

        /// <summary>Proves a rectangular clip uses the nominal glyph envelope's axes; corner-only rejection is otherwise unqualified.</summary>
        private static bool DiagnosticAlignedRectangle(OfficeTextInkClip clip, OfficeTransform toPage) {
            if (clip.FilledContours?.Count != 1 || !toPage.TryInvert(out OfficeTransform inverse)) return false;
            var ring = clip.FilledContours[0];
            int count = ring.Count;
            if (count == 5 && ring[0] == ring[4]) count--;
            if (count != 4) return false;
            var points = new OfficePoint[4];
            for (int index = 0; index < 4; index++) points[index] = inverse.TransformPoint(ring[index]);
            bool? previousHorizontal = null;
            for (int index = 0; index < 4; index++) {
                var start = points[index]; var end = points[(index + 1) % 4];
                bool horizontal = Math.Abs(start.Y - end.Y) <= 1E-9D;
                bool vertical = Math.Abs(start.X - end.X) <= 1E-9D;
                if (horizontal == vertical || previousHorizontal == horizontal) return false;
                previousHorizontal = horizontal;
            }
            return true;
        }

        private static OfficeTextInkClip DiagnosticMapClip(OfficeTextInkClip clip, OfficeTransform transform) {
            var mapped = new List<IReadOnlyList<OfficePoint>>();
            foreach (var ring in clip.FilledContours!) {
                var points = new List<OfficePoint>(ring.Count);
                foreach (OfficePoint point in ring) points.Add(transform.TransformPoint(point));
                mapped.Add(points);
            }
            return OfficeTextInkClip.FromFilledContours(mapped, clip.FillRule);
        }

        private static OfficeTextInkClip DiagnosticRectangleClip(TableCellContentFrame frame, OfficeTransform transform) {
            var contours = DiagnosticRectangle(frame.Left, frame.Top - frame.Height, frame.Width, frame.Height);
            foreach (var ring in contours) for (int index = 0; index < ring.Count; index++) ring[index] = transform.TransformPoint(ring[index]);
            return OfficeTextInkClip.FromFilledContours(contours, OfficeFillRule.NonZero);
        }
    }
}
