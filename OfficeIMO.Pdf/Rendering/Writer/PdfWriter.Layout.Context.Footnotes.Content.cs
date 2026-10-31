namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private PreparedFootnote PrepareFootnote(PdfFootnoteDefinition definition) {
            PdfCreateOptions noteOptions = footnotePageOptions ?? currentOpts;
            // Page defaults can differ from those at definition capture. Reject any
            // effective constraint that this fragment planner does not implement.
            PdfDocument.ValidateFootnoteBlocks(definition.Blocks, noteOptions);
            PdfDocument.ValidateFootnoteBlocks(definition.Separator, noteOptions, allowTables: false);
            PdfDocument.ValidateFootnoteBlocks(definition.ContinuationSeparator, noteOptions, allowTables: false);
            PdfDocument.ValidateFootnoteBlocks(definition.ContinuationNotice, noteOptions, allowTables: false);
            var key = (definition, currentPageBaseOptions, noteOptions.PageWidth.ToPoints() - noteOptions.MarginLeft.ToPoints() - noteOptions.MarginRight.ToPoints());
            if (preparedFootnotes.TryGetValue(key, out PreparedFootnote? prepared)) return prepared;
            PdfCreateOptions saved = currentOpts;
            currentOpts = noteOptions;
            try {
                var units = PrepareFootnoteUnits(definition.Blocks);
                if (!units.Any(unit => unit.IsVisual || unit.Height > 0D)) throw new ArgumentException("A footnote must have measurable content.");
                var separator = PrepareFootnoteUnits(definition.Separator);
                var continuation = PrepareFootnoteUnits(definition.ContinuationSeparator);
                var notice = PrepareFootnoteUnits(definition.ContinuationNotice);
                double maximum = currentOpts.PageHeight.ToPoints() - currentOpts.MarginTop.ToPoints() - currentOpts.MarginBottom.ToPoints() -
                    Math.Max(MeasureFootnoteUnits(separator), MeasureFootnoteUnits(continuation)) - MeasureFootnoteUnits(notice);
                foreach (FootnoteUnit unit in units) unit.MaximumKeepHeight = maximum;
                prepared = new PreparedFootnote(definition, units, separator, continuation, notice, key.Item3);
                preparedFootnotes.Add(key, prepared);
                return prepared;
            } finally { currentOpts = saved; }
        }

        private List<FootnoteUnit> PrepareFootnoteUnits(IReadOnlyList<IPdfBlock> blocks) {
            var units = new List<FootnoteUnit>();
            double frameWidth = currentOpts.PageWidth.ToPoints() - currentOpts.MarginLeft.ToPoints() - currentOpts.MarginRight.ToPoints();
            AppendFootnoteUnits(blocks, 0D, frameWidth, Array.Empty<FootnotePanel>(), units);
            CompleteFootnoteUnitMeasurements(units);
            return units;
        }

        private void AppendFootnoteUnits(IReadOnlyList<IPdfBlock> blocks, double frameX, double frameWidth,
            FootnotePanel[] panels, List<FootnoteUnit> units) {
            var spans = new List<FootnoteBlockSpan>();
            foreach (IPdfBlock block in blocks) {
                cancellationToken.ThrowIfCancellationRequested();
                int start = units.Count;
                bool keepWithNext = false;
                switch (block) {
                    case RichParagraphBlock paragraph:
                        PdfParagraphStyle? style = EffectiveParagraphStyle(paragraph);
                        AppendFootnoteParagraph(paragraph, style, frameX, frameWidth, panels, units);
                        keepWithNext = style?.KeepWithNext == true;
                        break;
                    case FlowBlock { IsReplayable: false, StaticBlocks: { } children }:
                        AppendFootnoteUnits(children, frameX, frameWidth, panels, units);
                        break;
                    case ContainerBlock container:
                        keepWithNext = AppendFootnotePanel(container, frameX, frameWidth, panels, units);
                        break;
                    case PdfListBlock list:
                        keepWithNext = AppendFootnoteList(list, frameX, frameWidth, panels, units);
                        break;
                    case TableBlock { Style.BorderFrame: not null } framedTable:
                        PdfTableStyle framedStyle = ResolveContainedTableWidth(framedTable.Style!, frameWidth);
                        var capturedTable = new TableBlock(framedTable.Cells.Select(row => row.ToArray()), framedTable.Align, framedStyle);
                        foreach (var link in framedTable.Links) capturedTable.AddLink(link.Key, link.Value);
                        keepWithNext = AppendFootnotePanel(new ContainerBlock(capturedTable), frameX, frameWidth, panels, units);
                        break;
                    case TableBlock table:
                        if (GetTableColumnCount(table) == 0) break;
                        PdfTableStyle tableStyle = PreparePairedTableBorders(table,
                            table.Style ?? currentOpts.DefaultTableStyleSnapshot ?? PdfTableStyles.Light());
                        PreparedTableGeometry geometry = PrepareTableGeometry(table, tableStyle, frameWidth);
                        double tableHeight = tableStyle.SpacingBefore.ToPoints() + GetTableRowsHeight(geometry.Rows.Heights,
                            0, table.Rows.Count, geometry.Gap) + tableStyle.SpacingAfter.ToPoints();
                        units.Add(new FootnoteUnit { Block = table, Table = geometry, Advance = tableHeight,
                            FrameX = frameX, FrameWidth = frameWidth, Panels = panels });
                        keepWithNext = tableStyle.KeepWithNext;
                        break;
                    case SpacerBlock spacer:
                        units.Add(new FootnoteUnit { Advance = spacer.Height, Panels = panels });
                        break;
                    case BookmarkBlock bookmark:
                        units.Add(new FootnoteUnit { Block = bookmark, Panels = panels, X = frameX });
                        break;
                    case HorizontalRuleBlock rule:
                        PdfHorizontalRuleStyle ruleStyle = ResolveHorizontalRuleStyle(rule, currentOpts);
                        ValidateHorizontalRule(ruleStyle);
                        units.Add(new FootnoteUnit { Block = rule, Before = ruleStyle.SpacingBefore.ToPoints(), Advance = ruleStyle.Thickness.ToPoints(),
                            After = ruleStyle.SpacingAfter.ToPoints(), X = frameX, Width = frameWidth, Panels = panels });
                        keepWithNext = ruleStyle.KeepWithNext;
                        break;
                    case ImageBlock image:
                        PdfImageStyle imageStyle = ResolveImageStyle(image, currentOpts);
                        PdfDocument.ValidateImageStyleForBox(imageStyle, image.Width, image.Height, nameof(imageStyle.ClipPath));
                        double scale = imageStyle.ScaleDownToFit ? Math.Min(1D, Math.Min(frameWidth / image.Width,
                            Math.Max(0D, currentOpts.PageHeight.ToPoints() - currentOpts.MarginTop.ToPoints() - currentOpts.MarginBottom.ToPoints() - imageStyle.SpacingBefore.ToPoints() - imageStyle.SpacingAfter.ToPoints()) / image.Height)) : 1D;
                        if (scale <= 0D) throw new ArgumentException("Footnote image spacing leaves no available page body height.");
                        PdfDocument.ValidateImageFitDimensions(image.Info, imageStyle.Fit, nameof(imageStyle.Fit));
                        var box = (Width: image.Width * scale, Height: image.Height * scale);
                        if (box.Width > frameWidth + .001D) throw new ArgumentException("Footnote image exceeds the available page body width.");
                        units.Add(new FootnoteUnit { Block = image, Before = imageStyle.SpacingBefore.ToPoints(), Advance = box.Height,
                            After = imageStyle.SpacingAfter.ToPoints(), Width = box.Width, FrameX = frameX, FrameWidth = frameWidth, Panels = panels });
                        keepWithNext = imageStyle.KeepWithNext;
                        break;
                    default:
                        throw new NotSupportedException("Unsupported footnote flow block " + block.GetType().Name + ".");
                }
                if (units.Count > MaximumTextLayoutLines) throw new InvalidDataException("Footnote flow exceeded the 100,000-fragment limit.");
                if (units.Count > start) spans.Add(new FootnoteBlockSpan(start, units.Count, keepWithNext, IsNonVisualFlowMarker(block)));
            }
            AddFootnoteKeepChains(units, spans);
        }

        private void PaintFootnoteUnit(FootnoteUnit unit, ref double top, double? originX = null) {
            cancellationToken.ThrowIfCancellationRequested();
            top -= unit.Before;
            if (unit.Block is BookmarkBlock bookmark) AddNamedDestination(bookmark, top);
            if (unit.Block is ImageBlock image) {
                PdfImageStyle style = ResolveImageStyle(image, currentOpts);
                double left = GetAlignedObjectX((originX ?? currentOpts.MarginLeft.ToPoints()) + unit.FrameX, unit.FrameWidth, unit.Width, style.Align);
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
                DrawHLine(sb, style.Color, style.Thickness.ToPoints(), (originX ?? currentOpts.MarginLeft.ToPoints()) + unit.X,
                    (originX ?? currentOpts.MarginLeft.ToPoints()) + unit.X + unit.Width, top - style.Thickness.ToPoints() / 2D, emitGeneratedStructure);
            }
            pageDirty = true;
            top -= unit.Advance + unit.After;
        }
    }
}
