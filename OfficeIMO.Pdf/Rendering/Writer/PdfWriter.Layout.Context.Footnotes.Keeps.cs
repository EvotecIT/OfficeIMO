namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private static IEnumerable<RichSeg> FootnoteReferenceSegments(IEnumerable<PdfFootnoteDefinition> definitions) {
            foreach (PdfFootnoteDefinition definition in definitions.Distinct())
                yield return new RichSeg(string.Empty, false, false, false, false, null, null, null, null, null,
                    PdfStandardFont.Helvetica, 1D, PdfTextBaseline.Normal, 0D, footnoteReference: definition);
        }

        // A keep chain admits its guaranteed body prefix as one atomic fragment. A terminal
        // paragraph contributes only its legal first fragment; later references remain unclaimed.
        private void ReserveFootnoteKeepChain(IList<IPdfBlock> blocks, int start) {
            if (isRunningContent || !KeepsWithNext(blocks[start])) return;
            bool hasReferences = false;
            for (int index = start; index < blocks.Count && index < start + MaxKeepWithNextChainBlocks; index++) {
                if (PdfFootnoteTraversal.References(blocks[index]).Any(note => !scheduledFootnotes.Contains(note))) { hasReferences = true; break; }
                if (!IsNonVisualFlowMarker(blocks[index]) && !KeepsWithNext(blocks[index])) break;
            }
            if (!hasReferences) return;
            var references = new List<RichSeg>();
            int inspected = 0;
            for (int index = start; index < blocks.Count && inspected < MaxKeepWithNextChainBlocks; index++) {
                IPdfBlock block = blocks[index];
                if (IsNonVisualFlowMarker(block)) continue;
                inspected++;
                if (KeepsWithNext(block)) references.AddRange(FootnoteReferenceSegments(PdfFootnoteTraversal.References(block)));
                else { references.AddRange(FirstVisualFootnoteSegments(block, currentOpts.MarginLeft, width)); break; }
            }
            if (!HasUnscheduledFootnote(references)) return;
            double height = MeasureKeepWithNextChainHeight(blocks, start, currentOpts.MarginLeft, width, currentOpts.DefaultFontSize, 0D);
            if (height > GetMaximumBlockContinuationHeight() + .001D) return;
            while (!TryReserveFootnotes(references, y - height, commit: false)) {
                if (!HasCurrentPageBodyContent())
                    throw new ArgumentException("The kept body group and first footnote fragments exceed the available page body height.");
                NewBlockFrame();
                height = MeasureKeepWithNextChainHeight(blocks, start, currentOpts.MarginLeft, width, currentOpts.DefaultFontSize, 0D);
                if (height > GetMaximumBlockContinuationHeight() + .001D) return;
            }
            TryReserveFootnotes(references, y - height, commit: true);
        }

        private IEnumerable<RichSeg> FirstVisualFootnoteSegments(IPdfBlock block, double frameX, double frameWidth) {
            if (block is RichParagraphBlock paragraph) {
                PdfParagraphStyle? style = EffectiveParagraphStyle(paragraph);
                double size = style?.FontSize ?? currentOpts.DefaultFontSize;
                double leading = GetParagraphLeading(style, size);
                var frame = GetParagraphTextFrame(style, frameX, frameWidth);
                var wrap = WrapRichRunsCoreWithFirstLineOrigin(paragraph.Runs, frame.Width, size, ChooseNormal(currentOpts.DefaultFont), leading,
                    frame.FirstLineWidth, frame.FirstLineX - frame.X, GetParagraphTabStopWidth(style), currentOpts,
                    style?.TabStops.ToArray(), lineSpacing: style?.LineSpacing, textWhitespaceMode: style?.TextWhitespaceMode, wordJustificationMode: style?.UseWordJustificationSpacing == true && paragraph.Align == PdfAlign.Justify ? style.JustificationMode : null, cancellationToken: cancellationToken);
                int count = style?.KeepTogether == true ? wrap.Lines.Count : Math.Min(wrap.Lines.Count, style == null ? 1 : Math.Max(1, ResolveMinimumOrphanLines(style)));
                if (style != null && wrap.Lines.Count < Math.Max(1, ResolveMinimumOrphanLines(style)) + Math.Max(1, ResolveMinimumWidowLines(style))) count = wrap.Lines.Count;
                return wrap.Lines.Take(count).SelectMany(line => line);
            }
            if (block is HeadingBlock heading) {
                PdfHeadingStyle? style = ResolveHeadingStyle(heading, currentOpts);
                double size = GetHeadingFontSize(heading, style);
                double before = y < GetCurrentFramePageStartY() - .001D || style?.ApplySpacingBeforeAtTop == true
                    ? style?.SpacingBefore ?? 0D : 0D;
                var savedPictures = pictureAnchors.Working.ToArray();
                try {
                    var wrap = WrapHeadingForFrame(heading, CreateHeadingTextRuns(heading, style, heading.Color ?? style?.Color),
                        style, frameX, frameWidth, size, GetHeadingLeading(style, size), y - before);
                    return wrap.Lines.Take(1).SelectMany(line => line).ToArray();
                } finally {
                    pictureAnchors.Working.Clear();
                    foreach (var entry in savedPictures) pictureAnchors.Working.Add(entry.Key, entry.Value);
                }
            }
            if (block is PdfListBlock list) {
                PreparedListLayout prepared = PrepareListLayout(list, frameWidth, currentOpts.DefaultFontSize, topLevelSpacing: true);
                return prepared.Items.Count == 0 ? Array.Empty<RichSeg>() : prepared.Items[0].TextLayout.Lines.Take(1).SelectMany(line => line);
            }
            if (block is ContainerBlock container) {
                PdfPanelStyle style = ResolveContainerStyle(container);
                var frame = ResolveContainerFrame(container, style, frameX, frameWidth);
                return FirstNestedFootnoteSegments(container.Blocks, frame.X + style.PaddingX, frame.ContentWidth);
            }
            if (block is TableBlock table) return GetFirstTableFootnoteSegments(table);
            if (block is FlowBlock { StaticBlocks: { } flow }) return FirstNestedFootnoteSegments(flow, frameX, frameWidth);
            if (block is SemanticBlock semantic) return FirstNestedFootnoteSegments(semantic.Blocks, frameX, frameWidth);
            if (block is LayerBlock layer) return FirstNestedFootnoteSegments(layer.Blocks, frameX, frameWidth);
            if (block is SectionBlock { Options.IncludeHeading: false } section) return FirstNestedFootnoteSegments(section.Blocks, frameX, frameWidth);
            return Array.Empty<RichSeg>();
        }

        private IEnumerable<RichSeg> FirstNestedFootnoteSegments(IReadOnlyList<IPdfBlock> blocks, double frameX, double frameWidth) {
            foreach (IPdfBlock block in blocks) {
                if (IsNonVisualFlowMarker(block)) continue;
                return FirstVisualFootnoteSegments(block, frameX, frameWidth);
            }
            return Array.Empty<RichSeg>();
        }

        private void ReserveAtomicFootnoteGroup(IEnumerable<IPdfBlock> blocks, Func<double> measureBodyHeight, Action advance) {
            RichSeg[] references = FootnoteReferenceSegments(blocks.SelectMany(PdfFootnoteTraversal.References)).ToArray();
            if (!HasUnscheduledFootnote(references)) return;
            double bodyHeight = measureBodyHeight();
            while (!TryReserveFootnotes(references, y - bodyHeight, commit: false)) {
                if (!HasCurrentPageBodyContent())
                    throw new ArgumentException("The kept body group and first footnote fragments exceed the available page body height.");
                advance();
                bodyHeight = measureBodyHeight();
            }
            TryReserveFootnotes(references, y - bodyHeight, commit: true);
        }
    }
}
