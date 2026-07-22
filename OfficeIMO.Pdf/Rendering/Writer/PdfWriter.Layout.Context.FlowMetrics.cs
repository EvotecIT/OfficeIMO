using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private void EnsureFixedFlowBlockFits(string blockName, double blockWidth, double blockHeight, double availableWidth) {
            if (blockWidth > availableWidth + 0.001) {
                throw new ArgumentException(blockName + " width exceeds the available page content width.");
            }

            double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
            if (blockHeight > availableHeight + 0.001) {
                throw new ArgumentException(blockName + " height exceeds the available page content height.");
            }
        }

        private (double Width, double Height) ResolveImageFlowBox(ImageBlock image, PdfImageStyle style, double frameWidth, double spacingBefore, double spacingAfter) {
            double imageWidth = image.Width;
            double imageHeight = image.Height;
            if (!style.ScaleDownToFit) {
                return (imageWidth, imageHeight);
            }

            double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom - spacingBefore - spacingAfter;
            double scale = 1D;
            if (imageWidth > frameWidth) {
                scale = Math.Min(scale, frameWidth / imageWidth);
            }

            if (availableHeight > 0D && imageHeight * scale > availableHeight) {
                scale = Math.Min(scale, availableHeight / imageHeight);
            }

            if (scale >= 1D) {
                return (imageWidth, imageHeight);
            }

            if (scale <= 0D || double.IsNaN(scale) || double.IsInfinity(scale)) {
                return (imageWidth, imageHeight);
            }

            return (imageWidth * scale, imageHeight * scale);
        }

        private static void ValidateHorizontalRule(PdfHorizontalRuleStyle rule) {
            if (rule.Thickness <= 0 || double.IsNaN(rule.Thickness) || double.IsInfinity(rule.Thickness)) {
                throw new ArgumentException("Horizontal rule thickness must be a positive finite value.");
            }

            if (rule.SpacingBefore < 0 || double.IsNaN(rule.SpacingBefore) || double.IsInfinity(rule.SpacingBefore)) {
                throw new ArgumentException("Horizontal rule spacing before must be a non-negative finite value.");
            }

            if (rule.SpacingAfter < 0 || double.IsNaN(rule.SpacingAfter) || double.IsInfinity(rule.SpacingAfter)) {
                throw new ArgumentException("Horizontal rule spacing after must be a non-negative finite value.");
            }
        }

        private static void ValidatePanelStyle(PanelStyle style, double panelWidth) {
            Guard.LeftCenterRightAlign(style.Align, nameof(style.Align), "Panel box");

            if (style.BorderWidth < 0 || double.IsNaN(style.BorderWidth) || double.IsInfinity(style.BorderWidth)) {
                throw new ArgumentException("Panel border width must be a non-negative finite value.");
            }

            if (style.PaddingX < 0 || double.IsNaN(style.PaddingX) || double.IsInfinity(style.PaddingX)) {
                throw new ArgumentException("Panel horizontal padding must be a non-negative finite value.");
            }

            if (style.PaddingY < 0 || double.IsNaN(style.PaddingY) || double.IsInfinity(style.PaddingY)) {
                throw new ArgumentException("Panel vertical padding must be a non-negative finite value.");
            }

            if (style.MaxWidth.HasValue && (style.MaxWidth.Value <= 0 || double.IsNaN(style.MaxWidth.Value) || double.IsInfinity(style.MaxWidth.Value))) {
                throw new ArgumentException("Panel maximum width must be a positive finite value.");
            }

            if (style.SpacingBefore < 0 || double.IsNaN(style.SpacingBefore) || double.IsInfinity(style.SpacingBefore)) {
                throw new ArgumentException("Panel spacing before must be a non-negative finite value.");
            }

            if (style.SpacingAfter < 0 || double.IsNaN(style.SpacingAfter) || double.IsInfinity(style.SpacingAfter)) {
                throw new ArgumentException("Panel spacing after must be a non-negative finite value.");
            }

            if (panelWidth - 2 * style.PaddingX <= 0) {
                throw new ArgumentException("Panel horizontal padding must leave a positive text width.");
            }
        }

        private void EnsurePanelSegmentCanFitLine(double topPadding, double lineHeight) {
            double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
            if (topPadding + lineHeight > availableHeight + 0.001D) {
                throw new ArgumentException("Panel vertical padding and first line height exceed the available page content height.");
            }
        }

        private void RenderListItem(System.Collections.Generic.IReadOnlyList<TextRun> runs, System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>> lines, System.Collections.Generic.List<double> lineHeights, string marker, PdfStandardFont markerFont, PdfNamedFontFace? markerNamedFont, double markerSize, PdfColor? markerColor, double markerX, double markerWidth, PdfAlign markerAlign, double textX, double textWidth, PdfAlign textAlign, PdfColor? color, double size, double leading, double spacingBefore, double spacingAfter, string? bookmarkName, ref int? listStructureElementIndex, ref LayoutResult.Page? listStructurePage) {
            int lineIndex = 0;
            bool firstSegment = true;
            var listFont = ChooseNormal(currentOpts.DefaultFont);
            PageStructElement? listItemElement = null;
            spacingBefore = ResolveTopLevelSpacingBefore(spacingBefore);
            if (spacingBefore > 0) {
                if (y - spacingBefore < currentOpts.MarginBottom) {
                    NewPage();
                    spacingBefore = 0D;
                }

                if (spacingBefore > 0) y -= spacingBefore;
            }

            while (lineIndex < lines.Count) {
                double available = y - currentOpts.MarginBottom;
                double firstLineHeight = GetRichLineHeight(lineHeights, lineIndex, leading);
                if (available < firstLineHeight) {
                    NewPage();
                    available = y - currentOpts.MarginBottom;
                    if (available < firstLineHeight) {
                        break;
                    }
                }

                int take = 0;
                double heightSum = 0;
                for (int k = lineIndex; k < lines.Count; k++) {
                    double lineHeight = GetRichLineHeight(lineHeights, k, leading);
                    if (heightSum + lineHeight > available) {
                        break;
                    }

                    heightSum += lineHeight;
                    take++;
                }

                if (take == 0) {
                    NewPage();
                    continue;
                }

                var segmentLines = new System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>>(take);
                var segmentHeights = new System.Collections.Generic.List<double>(take);
                for (int k = 0; k < take; k++) {
                    segmentLines.Add(lines[lineIndex + k]);
                    segmentHeights.Add(GetRichLineHeight(lineHeights, lineIndex + k, leading));
                }

                double baselineY = FirstTextBaselineFromTop(listFont, size, y);
                int? listElementIndex = firstSegment ? EnsurePageStructureContainer("L", ref listStructureElementIndex, ref listStructurePage) : null;
                int? listItemElementIndex = firstSegment ? RegisterStructureContainer("LI", listElementIndex) : null;
                if (firstSegment) {
                    if (listItemElementIndex.HasValue && currentPage != null) {
                        listItemElement = currentPage.StructElements[listItemElementIndex.Value];
                    }

                    if (!string.IsNullOrEmpty(bookmarkName)) {
                        AddNamedDestinationName(bookmarkName!, y);
                    }

                    var markerLines = new System.Collections.Generic.List<string>(1) { marker };
                    int? labelMarkedContentId = RegisterTextStructureElement("Lbl", listItemElementIndex);
                    if (markerNamedFont.HasValue) {
                        currentPage!.UsedNamedFonts.Add(markerNamedFont.Value);
                    } else {
                        MarkSimpleFont(markerFont);
                    }

                    WriteLinesInternal(
                        GetFontResourceName(markerFont, markerNamedFont, ChooseNormal(currentOpts.DefaultFont)),
                        markerSize,
                        leading,
                        markerX,
                        markerWidth,
                        baselineY,
                        markerLines,
                        markerAlign,
                        markerColor ?? color,
                        applyBaselineTweak: true,
                        structureType: "Lbl",
                        markedContentId: labelMarkedContentId,
                        namedFont: markerNamedFont);
                }

                pageDirty = true;
                int? bodyMarkedContentId = firstSegment || listItemElement == null
                    ? RegisterTextStructureElement("LBody", listItemElementIndex)
                    : RegisterTextStructureElement("LBody", listItemElement);
                WriteRichParagraph(sb, new RichParagraphBlock(runs, textAlign, color), segmentLines, segmentHeights, currentOpts, baselineY, size, leading, currentPage!.Annotations, textX, textWidth, structureType: "LBody", markedContentId: bodyMarkedContentId, structurePage: currentPage);
                MarkRichFonts(runs);
                y -= heightSum;
                lineIndex += take;
                firstSegment = false;
                if (lineIndex < lines.Count) {
                    NewPage();
                } else {
                    y -= spacingAfter;
                }
            }
        }

        private static double MeasureListKeepTogetherHeight(System.Collections.Generic.IReadOnlyList<TableCellTextLayout> itemLayouts, double leading, double spacingBefore, double itemSpacing, double spacingAfter) {
            double total = 0D;
            for (int itemIndex = 0; itemIndex < itemLayouts.Count; itemIndex++) {
                total += itemIndex == 0 ? spacingBefore : 0D;
                total += MeasureRichLinesHeight(itemLayouts[itemIndex].LineHeights, itemLayouts[itemIndex].LineCount, leading);
                total += itemIndex == itemLayouts.Count - 1 ? spacingAfter : itemSpacing;
            }

            return total;
        }

        private PdfParagraphStyle? EffectiveParagraphStyle(RichParagraphBlock paragraph) => paragraph.Style ?? currentOpts.DefaultParagraphStyleSnapshot;

        private double MeasureNextParagraphFirstVisualHeight(RichParagraphBlock paragraph, double frameX, double frameWidth, double fontSize) {
            PdfParagraphStyle? paragraphStyle = EffectiveParagraphStyle(paragraph);
            double leading = GetParagraphLeading(paragraphStyle, fontSize);
            double spacingBefore = GetParagraphSpacingBefore(paragraphStyle);
            var textFrame = GetParagraphTextFrame(paragraphStyle, frameX, frameWidth);
            var wrap = WrapRichRunsCoreWithFirstLineOrigin(paragraph.Runs, textFrame.Width, fontSize, ChooseNormal(currentOpts.DefaultFont), leading, textFrame.FirstLineWidth, textFrame.FirstLineX - textFrame.X, GetParagraphTabStopWidth(paragraphStyle), currentOpts, paragraphStyle?.TabStops.ToArray());
            if (wrap.LineHeights.Count == 0) {
                return spacingBefore;
            }

            int linesToReserve = 1;
            if (paragraphStyle?.KeepTogether == true) {
                linesToReserve = wrap.LineHeights.Count;
            } else if (paragraphStyle != null) {
                int minimumOrphanLines = ResolveMinimumOrphanLines(paragraphStyle);
                if (minimumOrphanLines > 1 && wrap.LineHeights.Count > 1) {
                    linesToReserve = Math.Min(minimumOrphanLines, wrap.LineHeights.Count);
                }
            }

            double height = spacingBefore;
            for (int i = 0; i < linesToReserve; i++) {
                height += GetRichLineHeight(wrap.LineHeights, i, leading);
            }

            return height;
        }

        private const int MaxKeepWithNextChainBlocks = 256;

        private double MeasureKeepWithNextChainHeight(System.Collections.Generic.IList<IPdfBlock> blocks, int startIndex, double frameX, double frameWidth, double fontSize) {
            double height = 0D;
            int inspectedBlocks = 0;
            for (int blockIndex = startIndex; blockIndex < blocks.Count; blockIndex++) {
                IPdfBlock block = blocks[blockIndex];
                if (IsNonVisualFlowMarker(block)) {
                    continue;
                }
                if (inspectedBlocks >= MaxKeepWithNextChainBlocks) {
                    break;
                }
                inspectedBlocks++;

                bool keepWithNext = KeepsWithNext(block);
                height += keepWithNext
                    ? MeasureKeepWithNextBlockHeight(block, frameX, frameWidth, fontSize)
                    : MeasureNextBlockFirstVisualHeight(block, frameX, frameWidth, fontSize);

                if (!keepWithNext) {
                    break;
                }
            }

            return height;
        }

        private static bool IsNonVisualFlowMarker(IPdfBlock block) =>
            block is BookmarkBlock;

        private double MeasureKeepWithNextBlockHeight(IPdfBlock block, double frameX, double frameWidth, double fontSize) {
            if (block is HeadingBlock heading) {
                return MeasureHeadingBlockHeight(heading, frameWidth);
            }

            if (block is RichParagraphBlock paragraph) {
                return MeasureParagraphBlockHeight(paragraph, frameX, frameWidth, fontSize);
            }

            if (block is BulletListBlock bullets) {
                return MeasureBulletListBlockHeight(bullets, frameWidth, fontSize);
            }

            if (block is NumberedListBlock numbered) {
                return MeasureNumberedListBlockHeight(numbered, frameWidth, fontSize);
            }

            if (block is TableBlock table) {
                return MeasureTableBlockHeight(table, frameWidth, fontSize, firstVisualOnly: false);
            }

            if (block is HorizontalRuleBlock rule) {
                PdfHorizontalRuleStyle style = ResolveHorizontalRuleStyle(rule, currentOpts);
                return style.SpacingBefore + style.Thickness + style.SpacingAfter;
            }

            if (block is ImageBlock image) {
                return MeasureImageBlockHeight(image, frameWidth);
            }

            if (block is ShapeBlock shape) {
                return MeasureShapeBlockHeight(shape);
            }

            if (block is DrawingBlock drawing) {
                return MeasureDrawingBlockHeight(drawing);
            }

            if (block is PanelParagraphBlock panel) {
                return MeasurePanelBlockHeight(panel, frameWidth, fontSize, firstVisualOnly: false);
            }

            if (block is RowBlock row) {
                return MeasureRowBlockHeight(row, frameX, frameWidth, fontSize, firstVisualOnly: false);
            }

            return MeasureNextBlockFirstVisualHeight(block, frameX, frameWidth, fontSize);
        }

        private double MeasureHeadingBlockHeight(HeadingBlock heading, double frameWidth) {
            PdfHeadingStyle? headingStyle = ResolveHeadingStyle(heading, currentOpts);
            double headingSize = GetHeadingFontSize(heading, headingStyle);
            double headingLeading = GetHeadingLeading(headingStyle, headingSize);
            double spacingBefore = headingStyle?.SpacingBefore ?? 0D;
            double spacingAfter = GetHeadingSpacingAfter(headingStyle, headingLeading);
            PdfColor? headingColor = heading.Color ?? headingStyle?.Color;
            System.Collections.Generic.IReadOnlyList<TextRun> headingRuns = CreateHeadingTextRuns(heading, headingStyle, headingColor);
            var wrap = WrapRichRunsCore(headingRuns, frameWidth, headingSize, ChooseNormal(currentOpts.DefaultFont), headingLeading, null, DefaultParagraphTabStopWidth, currentOpts);
            return spacingBefore + MeasureRichLinesHeight(wrap.LineHeights, wrap.Lines.Count, headingLeading) + spacingAfter;
        }

        private double MeasureParagraphBlockHeight(RichParagraphBlock paragraph, double frameX, double frameWidth, double fontSize) {
            PdfParagraphStyle? paragraphStyle = EffectiveParagraphStyle(paragraph);
            double leading = GetParagraphLeading(paragraphStyle, fontSize);
            double spacingBefore = GetParagraphSpacingBefore(paragraphStyle);
            double spacingAfter = GetParagraphSpacingAfter(paragraphStyle, leading);
            var textFrame = GetParagraphTextFrame(paragraphStyle, frameX, frameWidth);
            var wrap = WrapRichRunsCoreWithFirstLineOrigin(paragraph.Runs, textFrame.Width, fontSize, ChooseNormal(currentOpts.DefaultFont), leading, textFrame.FirstLineWidth, textFrame.FirstLineX - textFrame.X, GetParagraphTabStopWidth(paragraphStyle), currentOpts, paragraphStyle?.TabStops.ToArray());
            return spacingBefore + wrap.LineHeights.Sum() + spacingAfter;
        }

        private double MeasureBulletListBlockHeight(BulletListBlock bullets, double frameWidth, double fontSize) {
            PdfListStyle? listStyle = ResolveListStyle(bullets, currentOpts);
            double size = GetListFontSize(listStyle, fontSize);
            double markerSize = GetListMarkerFontSize(listStyle, size);
            double leading = Math.Max(GetListLeading(listStyle, size), GetListLeading(listStyle, markerSize));
            var baseFont = ChooseNormal(currentOpts.DefaultFont);
            PdfStandardFont markerFont = GetListMarkerFont(listStyle, currentOpts.DefaultFont);
            PdfNamedFontFace? markerNamedFont = GetListMarkerNamedFont(listStyle, currentOpts);
            const string bulletGlyph = "•";
            double estimatedBulletWidth = bullets.RichItems.Count == 0
                ? EstimateSimpleTextWidthForOptions(bulletGlyph, markerFont, markerNamedFont, markerSize, currentOpts)
                : bullets.RichItems.Max(item => EstimateSimpleTextWidthForOptions(item.Marker ?? bulletGlyph, markerFont, markerNamedFont, markerSize, currentOpts));
            double bulletWidth = GetListMarkerWidth(listStyle, estimatedBulletWidth);
            double spaceAdvance = EstimateSimpleTextWidthForOptions(" ", markerFont, markerNamedFont, markerSize, currentOpts);
            double markerGap = GetListMarkerGap(listStyle, spaceAdvance);
            double rawTextWidth = frameWidth - (listStyle?.LeftIndent ?? 0D) - bulletWidth - markerGap;
            double availableWidth = Math.Max(rawTextWidth, EstimateSimpleTextWidthForOptions("WW", baseFont, size, currentOpts));
            double itemSpacing = GetListItemSpacing(listStyle, leading);
            var wrappedItems = new System.Collections.Generic.List<TableCellTextLayout>(bullets.RichItems.Count);
            for (int itemIndex = 0; itemIndex < bullets.RichItems.Count; itemIndex++) {
                wrappedItems.Add(CreateListItemTextLayout(bullets.RichItems[itemIndex], availableWidth, baseFont, size, leading, currentOpts));
            }

            double listSpacingBefore = ResolveTopLevelSpacingBefore(listStyle?.SpacingBefore ?? 0D);
            double listSpacingAfter = listStyle?.GetSpacingAfter(itemSpacing) ?? itemSpacing;
            return MeasureListKeepTogetherHeight(wrappedItems, leading, listSpacingBefore, itemSpacing, listSpacingAfter);
        }

        private double MeasureNumberedListBlockHeight(NumberedListBlock numbered, double frameWidth, double fontSize) {
            PdfListStyle? listStyle = ResolveListStyle(numbered, currentOpts);
            double size = GetListFontSize(listStyle, fontSize);
            double markerSize = GetListMarkerFontSize(listStyle, size);
            double leading = Math.Max(GetListLeading(listStyle, size), GetListLeading(listStyle, markerSize));
            var baseFont = ChooseNormal(currentOpts.DefaultFont);
            PdfStandardFont markerFont = GetListMarkerFont(listStyle, currentOpts.DefaultFont);
            PdfNamedFontFace? markerNamedFont = GetListMarkerNamedFont(listStyle, currentOpts);
            int lastNumber = numbered.StartNumber + Math.Max(0, numbered.RichItems.Count - 1);
            string widestMarker = lastNumber.ToString(CultureInfo.InvariantCulture) + ".";
            double estimatedMarkerWidth = numbered.RichItems.Count == 0
                ? EstimateSimpleTextWidthForOptions(widestMarker, markerFont, markerNamedFont, markerSize, currentOpts)
                : numbered.RichItems
                    .Select((item, itemIndex) => item.Marker ?? ((numbered.StartNumber + itemIndex).ToString(CultureInfo.InvariantCulture) + "."))
                    .Max(marker => EstimateSimpleTextWidthForOptions(marker, markerFont, markerNamedFont, markerSize, currentOpts));
            double markerWidth = GetListMarkerWidth(listStyle, estimatedMarkerWidth);
            double spaceAdvance = EstimateSimpleTextWidthForOptions(" ", markerFont, markerNamedFont, markerSize, currentOpts);
            double markerGap = GetListMarkerGap(listStyle, spaceAdvance);
            double rawTextWidth = frameWidth - (listStyle?.LeftIndent ?? 0D) - markerWidth - markerGap;
            double availableWidth = Math.Max(rawTextWidth, EstimateSimpleTextWidthForOptions("WW", baseFont, size, currentOpts));
            double itemSpacing = GetListItemSpacing(listStyle, leading);
            var wrappedItems = new System.Collections.Generic.List<TableCellTextLayout>(numbered.RichItems.Count);
            for (int itemIndex = 0; itemIndex < numbered.RichItems.Count; itemIndex++) {
                wrappedItems.Add(CreateListItemTextLayout(numbered.RichItems[itemIndex], availableWidth, baseFont, size, leading, currentOpts));
            }

            double listSpacingBefore = ResolveTopLevelSpacingBefore(listStyle?.SpacingBefore ?? 0D);
            double listSpacingAfter = listStyle?.GetSpacingAfter(itemSpacing) ?? itemSpacing;
            return MeasureListKeepTogetherHeight(wrappedItems, leading, listSpacingBefore, itemSpacing, listSpacingAfter);
        }

        private bool KeepsWithNext(IPdfBlock block) {
            if (block is HeadingBlock heading) {
                return ResolveHeadingStyle(heading, currentOpts)?.KeepWithNext ?? true;
            }

            if (block is RichParagraphBlock paragraph) {
                return EffectiveParagraphStyle(paragraph)?.KeepWithNext == true;
            }

            if (block is BulletListBlock bullets) {
                return ResolveListStyle(bullets, currentOpts)?.KeepWithNext == true;
            }

            if (block is NumberedListBlock numbered) {
                return ResolveListStyle(numbered, currentOpts)?.KeepWithNext == true;
            }

            if (block is TableBlock table) {
                return (table.Style ?? currentOpts.DefaultTableStyleSnapshot ?? TableStyles.Light()).KeepWithNext;
            }

            if (block is HorizontalRuleBlock rule) {
                return ResolveHorizontalRuleStyle(rule, currentOpts).KeepWithNext;
            }

            if (block is ImageBlock image) {
                return ResolveImageStyle(image, currentOpts).KeepWithNext;
            }

            if (block is ShapeBlock shape) {
                return ResolveDrawingStyle(shape, currentOpts).KeepWithNext;
            }

            if (block is DrawingBlock drawing) {
                return ResolveDrawingStyle(drawing, currentOpts).KeepWithNext;
            }

            if (block is PanelParagraphBlock panel) {
                return ResolvePanelStyle(panel, currentOpts).KeepWithNext;
            }

            if (block is RowBlock row) {
                return (row.StyleSnapshot ?? currentOpts.DefaultRowStyleSnapshot)?.KeepWithNext == true;
            }

            return false;
        }

        private double MeasureNextBlockFirstVisualHeight(IPdfBlock block, double frameX, double frameWidth, double fontSize) {
            if (block is RichParagraphBlock paragraph) {
                return MeasureNextParagraphFirstVisualHeight(paragraph, frameX, frameWidth, fontSize);
            }

            if (block is HeadingBlock heading) {
                PdfHeadingStyle? headingStyle = ResolveHeadingStyle(heading, currentOpts);
                double headingSize = GetHeadingFontSize(heading, headingStyle);
                double headingLeading = GetHeadingLeading(headingStyle, headingSize);
                return (headingStyle?.SpacingBefore ?? 0D) + headingLeading;
            }

            if (block is SpacerBlock spacer) {
                return spacer.Height;
            }

            if (block is BulletListBlock bullets) {
                PdfListStyle? listStyle = ResolveListStyle(bullets, currentOpts);
                double size = GetListFontSize(listStyle, fontSize);
                double markerSize = GetListMarkerFontSize(listStyle, size);
                double leading = Math.Max(GetListLeading(listStyle, size), GetListLeading(listStyle, markerSize));
                string? firstItem = bullets.Items.Count > 0 ? bullets.Items[0] : null;
                if (firstItem == null) {
                    return listStyle?.SpacingBefore ?? 0D;
                }

                return (listStyle?.SpacingBefore ?? 0D) + leading;
            }

            if (block is NumberedListBlock numbered) {
                PdfListStyle? listStyle = ResolveListStyle(numbered, currentOpts);
                double size = GetListFontSize(listStyle, fontSize);
                double markerSize = GetListMarkerFontSize(listStyle, size);
                double leading = Math.Max(GetListLeading(listStyle, size), GetListLeading(listStyle, markerSize));
                string? firstItem = numbered.Items.Count > 0 ? numbered.Items[0] : null;
                if (firstItem == null) {
                    return listStyle?.SpacingBefore ?? 0D;
                }

                return (listStyle?.SpacingBefore ?? 0D) + leading;
            }

            if (block is PanelParagraphBlock panel) {
                return MeasurePanelBlockHeight(panel, frameWidth, fontSize, firstVisualOnly: true);
            }

            if (block is TableBlock table) {
                return MeasureTableBlockHeight(table, frameWidth, fontSize, firstVisualOnly: true);
            }

            if (block is DeferredTableBlock deferredTable) {
                return MeasureDeferredTableFirstVisualHeight(deferredTable, frameWidth, fontSize);
            }

            if (block is HorizontalRuleBlock rule) {
                PdfHorizontalRuleStyle style = ResolveHorizontalRuleStyle(rule, currentOpts);
                return style.SpacingBefore + style.Thickness + style.SpacingAfter;
            }

            if (block is TextFieldBlock textField) {
                return textField.SpacingBefore + textField.Height + textField.SpacingAfter;
            }

            if (block is CheckBoxBlock checkBox) {
                return checkBox.SpacingBefore + checkBox.Size + checkBox.SpacingAfter;
            }

            if (block is ChoiceFieldBlock choiceField) {
                return choiceField.SpacingBefore + choiceField.Height + choiceField.SpacingAfter;
            }

            if (block is RadioButtonGroupBlock radioButtonGroup) {
                return radioButtonGroup.SpacingBefore + radioButtonGroup.Height + radioButtonGroup.SpacingAfter;
            }

            if (block is ImageBlock image) {
                return MeasureImageBlockHeight(image, frameWidth);
            }

            if (block is ShapeBlock shape) {
                return MeasureShapeBlockHeight(shape);
            }

            if (block is DrawingBlock drawing) {
                return MeasureDrawingBlockHeight(drawing);
            }

            if (block is RowBlock row) {
                return MeasureRowBlockHeight(row, frameX, frameWidth, fontSize, firstVisualOnly: true);
            }

            return 0D;
        }

        private double MeasureDeferredTableFirstVisualHeight(DeferredTableBlock table, double frameWidth, double fontSize) {
            PdfTableStyle style = table.Style ?? currentOpts.DefaultTableStyleSnapshot ?? TableStyles.Light();
            using System.Collections.Generic.IEnumerator<DeferredTableBatch> batches = table.CreateBatches(style).GetEnumerator();
            return batches.MoveNext()
                ? MeasureTableBlockHeight(batches.Current.Table, frameWidth, fontSize, firstVisualOnly: true)
                : 0D;
        }

        private double MeasureImageBlockHeight(ImageBlock image, double frameWidth) {
            PdfImageStyle style = ResolveImageStyle(image, currentOpts);
            double spacingBefore = ResolveTopLevelSpacingBefore(style.SpacingBefore);
            var box = ResolveImageFlowBox(image, style, frameWidth, spacingBefore, style.SpacingAfter);
            return spacingBefore + box.Height + style.SpacingAfter;
        }

        private double MeasureShapeBlockHeight(ShapeBlock shape) {
            PdfDrawingStyle style = ResolveDrawingStyle(shape, currentOpts);
            return style.SpacingBefore + shape.Shape.Height + style.SpacingAfter;
        }

        private double MeasureDrawingBlockHeight(DrawingBlock drawing) {
            PdfDrawingStyle style = ResolveDrawingStyle(drawing, currentOpts);
            return style.SpacingBefore + drawing.Drawing.Height + style.SpacingAfter;
        }

        private double MeasurePanelBlockHeight(PanelParagraphBlock panel, double frameWidth, double fontSize, bool firstVisualOnly) {
            PanelStyle panelStyle = ResolvePanelStyle(panel, currentOpts);
            double innerWidth = panelStyle.MaxWidth.HasValue ? Math.Min(frameWidth, panelStyle.MaxWidth.Value) : frameWidth;
            ValidatePanelStyle(panelStyle, innerWidth);
            double size = fontSize;
            double leading = size * 1.4;
            double textWidth = innerWidth - 2 * panelStyle.PaddingX;
            var wrap = WrapRichRunsCore(panel.Runs, textWidth, size, ChooseNormal(currentOpts.DefaultFont), leading, null, DefaultParagraphTabStopWidth, currentOpts);
            int lineCount = firstVisualOnly ? Math.Min(1, wrap.LineHeights.Count) : wrap.LineHeights.Count;
            double spacingBefore = ResolveTopLevelSpacingBefore(panelStyle.SpacingBefore);
            double textHeight = MeasureRichLinesHeight(wrap.LineHeights, lineCount, leading);
            double spacingAfter = firstVisualOnly ? 0D : panelStyle.SpacingAfter;
            return spacingBefore + panelStyle.PaddingY + textHeight + panelStyle.PaddingY + spacingAfter;
        }

        private double MeasureRowBlockHeight(RowBlock row, double frameX, double frameWidth, double fontSize, bool firstVisualOnly) {
            int columns = row.Columns.Count;
            PdfRowStyle? rowStyle = row.StyleSnapshot ?? currentOpts.DefaultRowStyleSnapshot;
            double spacingBefore = ResolveTopLevelSpacingBefore(rowStyle?.SpacingBefore ?? 0D);
            if (columns == 0) {
                double spacingAfter = firstVisualOnly ? 0D : rowStyle?.SpacingAfter ?? 0D;
                return spacingBefore + spacingAfter;
            }

            double rowGap = row.GapOverride ?? rowStyle?.Gap ?? PdfRowStyle.DefaultGap;
            double totalGap = rowGap * Math.Max(0, columns - 1);
            if (totalGap >= frameWidth) {
                return spacingBefore;
            }

            double columnAreaWidth = frameWidth - totalGap;
            var columnWidths = new double[columns];
            double tallestFirstVisual = 0D;
            for (int columnIndex = 0; columnIndex < columns; columnIndex++) {
                RowColumn column = row.Columns[columnIndex];
                double columnWidth = Math.Max(0D, columnAreaWidth * (column.WidthPercent / 100D));
                columnWidths[columnIndex] = columnWidth;
                if (firstVisualOnly && column.Blocks.Count > 0) {
                    tallestFirstVisual = Math.Max(tallestFirstVisual, MeasureNextBlockFirstVisualHeight(column.Blocks[0], frameX, columnWidth, fontSize));
                }
            }

            if (firstVisualOnly) {
                return spacingBefore + tallestFirstVisual;
            }

            var columnItems = BuildRowColumnItems(row, columnWidths);
            double contentHeight = 0D;
            foreach (var items in columnItems) {
                contentHeight = Math.Max(contentHeight, MeasureRowKeepTogetherHeight(items));
            }

            return spacingBefore + contentHeight + (rowStyle?.SpacingAfter ?? 0D);
        }

        private double MeasureTableBlockHeight(TableBlock table, double frameWidth, double fontSize, bool firstVisualOnly) {
            PdfTableStyle style = table.Style ?? currentOpts.DefaultTableStyleSnapshot ?? TableStyles.Light();
            int columns = GetTableColumnCount(table);
            if (columns == 0 || table.Rows.Count == 0) {
                return style.SpacingBefore + (firstVisualOnly ? 0D : style.SpacingAfter);
            }

            double columnGap = GetTableCellSpacing(style);
            double rowGap = columnGap;
            ValidateTableRoleRowCounts(style, table.Rows.Count);
            int headerRowCount = style.HeaderRowCount;
            int footerRowCount = style.FooterRowCount;
            int footerStartRowIndex = table.Rows.Count - footerRowCount;
            ValidateTableCellStyleCoordinates(style, table, columns);
            ValidateTableColumnStyleBounds(style, columns);
            ValidateTableRowStyleBounds(style, table.Rows.Count);
            ValidateTableRowSpansWithinRoleBoundaries(table, columns, headerRowCount, footerStartRowIndex);
            double tableFontSize = GetTableBodyFontSize(style, fontSize);
            TableColumnLayout columnLayout = ResolveTableColumnLayout(table, currentOpts, style, columns, frameWidth, tableFontSize, headerRowCount, footerStartRowIndex);

            var rowLines = new TableCellTextLayout[table.Rows.Count][];
            var rowHeights = new double[table.Rows.Count];
            var rowLeadings = new double[table.Rows.Count];
            for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++) {
                bool rowUsesBold = GetTableRowBold(style, rowIndex, headerRowCount, footerStartRowIndex);
                double originalRowSize = GetTableRowFontSize(style, rowIndex, headerRowCount, footerStartRowIndex, fontSize);
                double rowSize = ResolveTableRowShrinkFontSize(table, style, rowIndex, columns, columnLayout.Widths, columnGap, originalRowSize, rowUsesBold, currentOpts);
                double runFontSizeScale = GetTableRunFontSizeScale(table, style, rowIndex, columns, columnLayout.Widths, columnGap, originalRowSize, rowSize, rowUsesBold, currentOpts);
                double rowLeading = GetTableLeading(style, rowSize);
                rowLeadings[rowIndex] = rowLeading;
                rowLines[rowIndex] = new TableCellTextLayout[columns];
                double maxRequiredHeight = rowLeading + GetTableRowMaxPaddingTop(table, style, rowIndex, columns) + GetTableRowMaxPaddingBottom(table, style, rowIndex, columns);
                for (int columnIndex = 0; columnIndex < columns; columnIndex++) {
                    rowLines[rowIndex][columnIndex] = new TableCellTextLayout(new System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>> { new() }, new System.Collections.Generic.List<double> { rowLeading });
                }

                var cells = GetTableCellLayouts(table, rowIndex, columns);
                for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++) {
                    TableCellLayout cell = cells[cellIndex];
                    PdfStandardFont cellFont = GetTableRowFont(currentOpts, rowUsesBold);
                    double cellWidth = GetTableCellWidth(columnLayout.Widths, cell.Column, cell.ColumnSpan, columnGap);
                    double innerWidth = Math.Max(1D, cellWidth - GetTableCellPaddingLeft(style, rowIndex, cell.Column) - GetTableCellPaddingRight(style, rowIndex, cell.Column));
                    TableCellTextLayout lines = CreateTableCellTextLayout(cell, innerWidth, cellFont, rowSize, rowLeading, currentOpts, runFontSizeScale, style.MinimumShrinkFontSize ?? 6D);
                    rowLines[rowIndex][cell.Column] = lines;
                    if (cell.RowSpan <= 1) {
                        maxRequiredHeight = Math.Max(maxRequiredHeight, MeasureTableCellContentHeight(cell, lines, 0, lines.LineCount, rowLeading, innerWidth) + GetTableCellPaddingTop(style, rowIndex, cell.Column) + GetTableCellPaddingBottom(style, rowIndex, cell.Column));
                    }
                }

                rowHeights[rowIndex] = ResolveTableRowHeight(style, rowIndex, maxRequiredHeight);
            }

            ApplyTableRowSpanHeights(table, style, columns, columnLayout.Widths, rowLines, rowHeights, rowLeadings, columnGap, rowGap);

            double captionHeight = 0D;
            if (!string.IsNullOrWhiteSpace(style.Caption)) {
                double captionSize = style.CaptionFontSize ?? fontSize;
                double captionLeading = captionSize * 1.25D;
                var captionRuns = new[] { TextRun.Normal(style.Caption!, style.CaptionColor, captionSize) };
                var captionWrap = WrapRichRunsCore(captionRuns, columnLayout.Width, captionSize, ChooseNormal(currentOpts.DefaultFont), captionLeading, null, DefaultParagraphTabStopWidth, currentOpts);
                captionHeight = MeasureRichLinesHeight(captionWrap.LineHeights, captionWrap.Lines.Count, captionLeading) + style.CaptionSpacingAfter;
            }

            int measuredRowCount = firstVisualOnly ? 1 : rowHeights.Length;
            double tableHeight = style.SpacingBefore + captionHeight + GetTableRowsHeight(rowHeights, 0, measuredRowCount, rowGap);
            return firstVisualOnly ? tableHeight : tableHeight + style.SpacingAfter;
        }

        private void ConsumeSpacer(double height) {
            double remaining = height;
            while (remaining > 0.001D) {
                double available = y - currentOpts.MarginBottom;
                if (available <= 0.5D) {
                    NewPage();
                    continue;
                }

                double consumed = Math.Min(remaining, available);
                y -= consumed;
                remaining -= consumed;
                if (remaining > 0.001D) {
                    NewPage();
                }
            }
        }

        private void RenderHorizontalRuleBlock(HorizontalRuleBlock block, double containerX, double containerWidth) {
            PdfHorizontalRuleStyle ruleStyle = ResolveHorizontalRuleStyle(block, currentOpts);
            ValidateHorizontalRule(ruleStyle);
            double spacingBefore = ResolveTopLevelSpacingBefore(ruleStyle.SpacingBefore);
            double needed = spacingBefore + ruleStyle.Thickness + ruleStyle.SpacingAfter;
            EnsureFixedFlowBlockFits("Horizontal rule", containerWidth, needed, containerWidth);
            if (y - needed < currentOpts.MarginBottom) {
                NewPage();
                spacingBefore = 0D;
            }
            if (spacingBefore > 0) y -= spacingBefore;
            double yLine = y - ruleStyle.Thickness * 0.5;
            DrawHLine(sb, ruleStyle.Color, ruleStyle.Thickness, containerX, containerX + containerWidth, yLine, emitGeneratedStructure);
            pageDirty = true;
            y -= ruleStyle.Thickness + ruleStyle.SpacingAfter;
        }

    }
}
