namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private void AppendFootnoteParagraph(RichParagraphBlock paragraph, PdfParagraphStyle? style,
            double frameX, double frameWidth, FootnotePanel[] panels, List<FootnoteUnit> units) {
            double size = style?.FontSize ?? currentOpts.DefaultFontSize;
            double leading = GetParagraphLeading(style, size);
            var frame = GetParagraphTextFrame(style, frameX, frameWidth);
            var wrapped = WrapRichRunsCoreWithFirstLineOrigin(paragraph.Runs, frame.Width, size, ChooseNormal(currentOpts.DefaultFont), leading,
                frame.FirstLineWidth, frame.FirstLineX - frame.X, GetParagraphTabStopWidth(style), currentOpts,
                style?.TabStops.ToArray(), lineSpacing: style?.LineSpacing, textWhitespaceMode: style?.TextWhitespaceMode, wordJustificationMode: style?.UseWordJustificationSpacing == true && paragraph.Align == PdfAlign.Justify ? style.JustificationMode : null);
            var captured = new RichParagraphBlock(paragraph.Runs, paragraph.Align, paragraph.DefaultColor, style);
            for (int index = 0; index < wrapped.Lines.Count; index++) {
                if (units.Count >= MaximumTextLayoutLines) throw new InvalidDataException("Footnote flow exceeded the 100,000-fragment limit.");
                double advance = Math.Max(wrapped.LineHeights[index], GetRichLineFrameFitHeight(wrapped.Lines[index], wrapped.LineHeights[index]));
                if (wrapped.Lines[index] is RichLine barLine) barLine.BarLineHeight = advance;
                units.Add(new FootnoteUnit {
                    Block = captured, Line = wrapped.Lines[index],
                    Advance = advance,
                    Before = index == 0 ? GetParagraphSpacingBefore(style) : 0D,
                    After = index == wrapped.Lines.Count - 1 ? GetParagraphSpacingAfter(style, leading) : 0D,
                    Size = size, Leading = leading, X = index == 0 ? frame.FirstLineX : frame.X,
                    Width = index == 0 ? frame.FirstLineWidth : frame.Width,
                    ParagraphLine = index, ParagraphLines = wrapped.Lines.Count,
                    MinimumOrphanLines = style == null ? 1 : Math.Max(1, ResolveMinimumOrphanLines(style)),
                    MinimumWidowLines = style == null ? 1 : Math.Max(1, ResolveMinimumWidowLines(style)),
                    KeepParagraphTogether = style?.KeepTogether == true, Panels = panels
                });
            }
        }

        // Paint one actual page fragment through the ordinary paragraph owner.
        // Its start remains the clip boundary for later exact-height picture
        // lines, and a continued fragment receives a fresh start on its page.
        private void PaintFootnoteParagraphFragment(IReadOnlyList<FootnoteUnit> units, int start, int count, ref double top, double? originX = null, int? structureParent = null, string structureType = "Note") {
            FootnoteUnit first = units[start];
            FootnoteUnit last = units[start + count - 1];
            var paragraph = (RichParagraphBlock)first.Block!;
            var lines = new List<List<RichSeg>>(count);
            var heights = new List<double>(count);
            double advance = 0D;
            for (int index = start; index < start + count; index++) {
                cancellationToken.ThrowIfCancellationRequested();
                lines.Add(units[index].Line!);
                heights.Add(units[index].Advance);
                advance += units[index].Advance;
            }
            top -= first.Before;
            int? markedContentId = RegisterTextStructureElement(structureType, structureParent);
            if (first.List != null && first.ListItem != null && first.ParagraphLine == 0) PaintFootnoteListMarker(first, top, originX);
            MarkRichFonts(paragraph.Runs);
            bool firstLineFrame = first.ParagraphLine == 0 && count > 1;
            FootnoteUnit normalFrame = firstLineFrame ? units[start + 1] : first;
            WriteRichParagraph(sb, paragraph, lines, heights, currentOpts,
                FirstTextBaselineFromTop(ChooseNormal(currentOpts.DefaultFont), first.Size, top), first.Size, first.Leading,
                currentPage!.Annotations, (originX ?? currentOpts.MarginLeft) + normalFrame.X, normalFrame.Width,
                firstLineXOverride: firstLineFrame ? (originX ?? currentOpts.MarginLeft) + first.X : null,
                firstLineWidthOverride: firstLineFrame ? first.Width : null,
                structureType: structureType, markedContentId: markedContentId, structurePage: currentPage,
                hasFollowingParagraphLines: last.ParagraphLine + 1 < last.ParagraphLines, paragraphSpacingAfter: last.After,
                paragraphSpacingBefore: first.Before, suppressInlineAccessibility: _suppressCanvasAccessibilityWrappers);
            pageDirty = true;
            top -= advance + last.After;
        }
    }
}
