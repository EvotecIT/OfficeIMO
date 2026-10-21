namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private sealed class FootnoteUnit {
            public IPdfBlock? Block { get; init; }
            public List<RichSeg>? Line { get; init; }
            public double Advance { get; init; }
            public double Before { get; init; }
            public double After { get; init; }
            public double Size { get; init; }
            public double Leading { get; init; }
            public double X { get; init; }
            public double Width { get; init; }
            public int ParagraphLine { get; init; }
            public int ParagraphLines { get; init; }
            public int MinimumOrphanLines { get; init; }
            public int MinimumWidowLines { get; init; }
            public bool KeepParagraphTogether { get; init; }
            public double Height => Before + Advance + After;
            public bool IsVisual => Line != null || Block is ImageBlock or HorizontalRuleBlock;
        }

        private PreparedFootnote PrepareFootnote(PdfFootnoteDefinition definition) {
            PdfOptions noteOptions = footnotePageOptions ?? currentOpts;
            // Page defaults can differ from those at definition capture. Reject any
            // effective constraint that this fragment planner does not implement.
            PdfDocument.ValidateFootnoteBlocks(definition.Blocks, noteOptions);
            PdfDocument.ValidateFootnoteBlocks(definition.Separator, noteOptions);
            PdfDocument.ValidateFootnoteBlocks(definition.ContinuationSeparator, noteOptions);
            PdfDocument.ValidateFootnoteBlocks(definition.ContinuationNotice, noteOptions);
            var key = (definition, currentPageBaseOptions, noteOptions.PageWidth - noteOptions.MarginLeft - noteOptions.MarginRight);
            if (preparedFootnotes.TryGetValue(key, out PreparedFootnote? prepared)) return prepared;
            PdfOptions saved = currentOpts;
            currentOpts = noteOptions;
            try {
                var units = PrepareFootnoteUnits(definition.Blocks);
                if (!units.Any(unit => unit.IsVisual || unit.Height > 0D)) throw new ArgumentException("A footnote must have measurable content.");
                prepared = new PreparedFootnote(definition, units, PrepareFootnoteUnits(definition.Separator),
                    PrepareFootnoteUnits(definition.ContinuationSeparator), PrepareFootnoteUnits(definition.ContinuationNotice), key.Item3);
                preparedFootnotes.Add(key, prepared);
                return prepared;
            } finally { currentOpts = saved; }
        }

        private List<FootnoteUnit> PrepareFootnoteUnits(IReadOnlyList<IPdfBlock> blocks) {
            var units = new List<FootnoteUnit>();
            double frameWidth = currentOpts.PageWidth - currentOpts.MarginLeft - currentOpts.MarginRight;
            foreach (IPdfBlock block in blocks) {
                cancellationToken.ThrowIfCancellationRequested();
                switch (block) {
                    case RichParagraphBlock paragraph:
                        PdfParagraphStyle? style = EffectiveParagraphStyle(paragraph);
                        double size = style?.FontSize ?? currentOpts.DefaultFontSize;
                        double leading = GetParagraphLeading(style, size);
                        var frame = GetParagraphTextFrame(style, 0D, frameWidth);
                        var wrapped = WrapRichRunsCoreWithFirstLineOrigin(paragraph.Runs, frame.Width, size, ChooseNormal(currentOpts.DefaultFont), leading,
                            frame.FirstLineWidth, frame.FirstLineX - frame.X, GetParagraphTabStopWidth(style), currentOpts,
                            style?.TabStops.ToArray(), lineSpacing: style?.LineSpacing, textWhitespaceMode: style?.TextWhitespaceMode);
                        for (int index = 0; index < wrapped.Lines.Count; index++) {
                            if (units.Count >= MaximumTextLayoutLines) throw new InvalidDataException("Footnote flow exceeded the 100,000-fragment limit.");
                            units.Add(new FootnoteUnit {
                                Block = paragraph, Line = wrapped.Lines[index], Advance = Math.Max(wrapped.LineHeights[index], GetRichLineFrameFitHeight(wrapped.Lines[index], wrapped.LineHeights[index])),
                                Before = index == 0 ? GetParagraphSpacingBefore(style) : 0D,
                                After = index == wrapped.Lines.Count - 1 ? GetParagraphSpacingAfter(style, leading) : 0D,
                                Size = size, Leading = leading, X = index == 0 ? frame.FirstLineX : frame.X,
                                Width = index == 0 ? frame.FirstLineWidth : frame.Width,
                                ParagraphLine = index, ParagraphLines = wrapped.Lines.Count,
                                MinimumOrphanLines = style == null ? 1 : Math.Max(1, ResolveMinimumOrphanLines(style)),
                                MinimumWidowLines = style == null ? 1 : Math.Max(1, ResolveMinimumWidowLines(style)),
                                KeepParagraphTogether = style?.KeepTogether == true
                            });
                        }
                        break;
                    case FlowBlock { IsReplayable: false, StaticBlocks: { } children }:
                        units.AddRange(PrepareFootnoteUnits(children));
                        break;
                    case SpacerBlock spacer:
                        units.Add(new FootnoteUnit { Advance = spacer.Height });
                        break;
                    case BookmarkBlock bookmark:
                        units.Add(new FootnoteUnit { Block = bookmark });
                        break;
                    case HorizontalRuleBlock rule:
                        PdfHorizontalRuleStyle ruleStyle = ResolveHorizontalRuleStyle(rule, currentOpts);
                        ValidateHorizontalRule(ruleStyle);
                        units.Add(new FootnoteUnit { Block = rule, Before = ruleStyle.SpacingBefore, Advance = ruleStyle.Thickness, After = ruleStyle.SpacingAfter, Width = frameWidth });
                        break;
                    case ImageBlock image:
                        PdfImageStyle imageStyle = ResolveImageStyle(image, currentOpts);
                        PdfDocument.ValidateImageStyleForBox(imageStyle, image.Width, image.Height, nameof(imageStyle.ClipPath));
                        double scale = imageStyle.ScaleDownToFit ? Math.Min(1D, Math.Min(frameWidth / image.Width,
                            Math.Max(0D, currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom - imageStyle.SpacingBefore - imageStyle.SpacingAfter) / image.Height)) : 1D;
                        if (scale <= 0D) throw new ArgumentException("Footnote image spacing leaves no available page body height.");
                        PdfDocument.ValidateImageFitDimensions(image.Info, imageStyle.Fit, nameof(imageStyle.Fit));
                        var box = (Width: image.Width * scale, Height: image.Height * scale);
                        if (box.Width > frameWidth + .001D) throw new ArgumentException("Footnote image exceeds the available page body width.");
                        units.Add(new FootnoteUnit { Block = image, Before = imageStyle.SpacingBefore, Advance = box.Height, After = imageStyle.SpacingAfter, Width = box.Width });
                        break;
                    default:
                        throw new NotSupportedException("Unsupported footnote flow block " + block.GetType().Name + ".");
                }
                if (units.Count > MaximumTextLayoutLines) throw new InvalidDataException("Footnote flow exceeded the 100,000-fragment limit.");
            }
            return units;
        }

        private void PaintFootnoteUnit(FootnoteUnit unit, ref double top) {
            cancellationToken.ThrowIfCancellationRequested();
            top -= unit.Before;
            if (unit.Block is BookmarkBlock bookmark) AddNamedDestination(bookmark, top);
            if (unit.Block is RichParagraphBlock paragraph && unit.Line != null) {
                var lines = new List<List<RichSeg>> { unit.Line };
                var heights = new List<double> { unit.Advance };
                int? markedContentId = RegisterTextStructureElement("Note");
                MarkRichFonts(paragraph.Runs);
                WriteRichParagraph(sb, paragraph, lines, heights, currentOpts,
                    FirstTextBaselineFromTop(ChooseNormal(currentOpts.DefaultFont), unit.Size, top), unit.Size, unit.Leading,
                    currentPage!.Annotations, currentOpts.MarginLeft + unit.X, unit.Width,
                    structureType: "Note", markedContentId: markedContentId, structurePage: currentPage,
                    hasFollowingParagraphLines: unit.ParagraphLine + 1 < unit.ParagraphLines);
            } else if (unit.Block is ImageBlock image) {
                PdfImageStyle style = ResolveImageStyle(image, currentOpts);
                double frameWidth = currentOpts.PageWidth - currentOpts.MarginLeft - currentOpts.MarginRight;
                double left = GetAlignedObjectX(currentOpts.MarginLeft, frameWidth, unit.Width, style.Align);
                PageImage placed = CreatePageImage(image, style, left, top - unit.Advance, unit.Width, unit.Advance);
                currentPage!.Images.Add(placed);
                AddImageLinkAnnotation(image, style, placed, left, top - unit.Advance, unit.Width, unit.Advance);
                if (!string.IsNullOrWhiteSpace(placed.AlternativeText)) {
                    int? id = RegisterFigureStructureElement(placed.AlternativeText!);
                    placed.MarkedContentId = id;
                    placed.StructElementIndex = FindStructElementIndex(currentPage, id, "Figure");
                }
            } else if (unit.Block is HorizontalRuleBlock rule) {
                PdfHorizontalRuleStyle style = ResolveHorizontalRuleStyle(rule, currentOpts);
                DrawHLine(sb, style.Color, style.Thickness, currentOpts.MarginLeft,
                    currentOpts.MarginLeft + unit.Width, top - style.Thickness / 2D, emitGeneratedStructure);
            }
            pageDirty = true;
            top -= unit.Advance + unit.After;
        }
    }
}
