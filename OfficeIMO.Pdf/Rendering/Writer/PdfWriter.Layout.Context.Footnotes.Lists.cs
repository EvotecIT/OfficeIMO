namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private bool AppendFootnoteList(PdfListBlock block, double frameX, double frameWidth,
            FootnotePanel[] panels, List<FootnoteUnit> units) {
            PreparedListLayout prepared = PrepareListLayout(block, frameWidth, currentOpts.DefaultFontSize.ToPoints(), topLevelSpacing: false);
            int start = units.Count;
            for (int itemIndex = 0; itemIndex < prepared.Items.Count; itemIndex++) {
                PreparedListItem item = prepared.Items[itemIndex];
                var paragraph = new RichParagraphBlock(item.Item.Runs, block.Align, block.Color ?? prepared.Style?.Color);
                for (int line = 0; line < item.TextLayout.Lines.Count; line++) {
                    if (units.Count >= MaximumTextLayoutLines) throw new InvalidDataException("Footnote flow exceeded the 100,000-fragment limit.");
                    double advance = GetRichLineHeight(item.TextLayout.LineHeights, line, prepared.Leading);
                    units.Add(new FootnoteUnit {
                        Block = paragraph, Line = item.TextLayout.Lines[line],
                        Advance = Math.Max(advance, GetRichLineFrameFitHeight(item.TextLayout.Lines[line], advance)),
                        Before = itemIndex == 0 && line == 0 ? prepared.SpacingBefore : 0D,
                        After = line == item.TextLayout.Lines.Count - 1 ?
                            itemIndex == prepared.Items.Count - 1 ? prepared.SpacingAfter : prepared.ItemSpacing : 0D,
                        X = frameX + prepared.ListLeftIndent + prepared.MarkerWidth + prepared.MarkerGap,
                        FrameX = frameX, Width = prepared.AlignmentWidth, Size = prepared.Size, Leading = prepared.Leading,
                        ParagraphLine = line, ParagraphLines = item.TextLayout.Lines.Count,
                        Panels = panels, List = prepared, ListItem = item
                    });
                }
            }
            if (prepared.Style?.KeepTogether == true)
                AddFootnoteKeepRange(units, start, units.Count,
                    relaxIfOversized: prepared.Style.KeepTogetherOverflow == PdfKeepTogetherOverflowBehavior.SplitAcrossPages);
            return prepared.Style?.KeepWithNext == true;
        }

        private void PaintFootnoteListMarker(FootnoteUnit unit, double top, double? originX = null) {
            PreparedListLayout prepared = unit.List!;
            PreparedListItem item = unit.ListItem!;
            if (!string.IsNullOrEmpty(item.Item.BookmarkName)) AddNamedDestinationName(item.Item.BookmarkName!, top);
            foreach (string alias in item.Item.BookmarkAliases) AddNamedDestinationName(alias, top);
            var placement = GetListMarkerPlacement(prepared, item);
            double markerX = (originX ?? currentOpts.MarginLeft.ToPoints()) + unit.FrameX + placement.X;
            double baseline = AdjustRichLineBaseline(FirstTextBaselineFromTop(ChooseNormal(currentOpts.DefaultFont), unit.Size, top),
                unit.Line!, currentOpts, unit.Size);
            global::OfficeIMO.OfficeColor? color = prepared.Style?.MarkerColor ?? prepared.Block.Color ?? prepared.Style?.Color;
            int? markedContentId = RegisterTextStructureElement("Lbl");
            if (WriteListMarkerFallback(item.MarkerFallback, markerX, placement.Width, baseline, prepared.MarkerSize,
                    prepared.Leading, placement.Align, color, markedContentId)) return;
            if (prepared.MarkerNamedFont.HasValue) currentPage!.UsedNamedFonts.Add(prepared.MarkerNamedFont.Value);
            else MarkSimpleFont(prepared.MarkerFont);
            WriteLinesInternal(GetFontResourceName(prepared.MarkerFont, prepared.MarkerNamedFont, ChooseNormal(currentOpts.DefaultFont)),
                prepared.MarkerSize, prepared.Leading, markerX, placement.Width,
                baseline + ResolveFontLineAlignmentRise(prepared.MarkerFont, prepared.MarkerNamedFont, prepared.MarkerSize,
                    currentOpts, prepared.Style?.LineSpacing), new List<string> { item.Marker }, placement.Align, color,
                applyBaselineTweak: true, structureType: "Lbl", markedContentId: markedContentId,
                namedFont: prepared.MarkerNamedFont, featureSettings: prepared.Style?.MarkerFeatureSettings);
        }
    }
}
