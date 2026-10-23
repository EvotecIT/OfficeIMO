namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private sealed class FootnotePanel {
            public FootnotePanel(PdfPanelStyle style, double x, double width, int start, ContainerBlock container) {
                Style = style; X = x; Width = width; Start = start; Container = container;
            }
            public PdfPanelStyle Style { get; }
            public double X { get; }
            public double Width { get; }
            public int Start { get; }
            public int End { get; set; }
            public ContainerBlock Container { get; }

            public double GetTopPadding(bool continuation, IReadOnlyDictionary<int, TableFragmentPlan>? tables) {
                if (continuation && tables != null && tables.TryGetValue(Start, out var table))
                    return ResolveTableFrameContinuationTopPadding(Container, table.Start.Row) ?? Style.GetFragmentTopPadding(true);
                return Style.GetFragmentTopPadding(continuation);
            }
        }

        private sealed class FootnotePanelPaint {
            public FootnotePanelPaint(FootnotePanel panel, int insertion, double top, bool continuation) {
                Panel = panel; Insertion = insertion; Top = top; IsContinuation = continuation;
            }
            public FootnotePanel Panel { get; }
            public int Insertion { get; }
            public double Top { get; }
            public bool IsContinuation { get; }
            public double? TableContentBottom { get; set; }
        }

        private bool AppendFootnotePanel(ContainerBlock container, double frameX, double frameWidth,
            FootnotePanel[] ancestors, List<FootnoteUnit> units) {
            PdfPanelStyle style = (container.UseDefaultPanelStyle ? currentOpts.DefaultPanelStyleSnapshot ?? container.Style : container.Style).Clone();
            var frame = ResolveContainerFrame(container, style, frameX, frameWidth);
            int start = units.Count;
            var panel = new FootnotePanel(style, frame.X, frame.Width, start, container);
            var panels = new FootnotePanel[ancestors.Length + 1];
            Array.Copy(ancestors, panels, ancestors.Length);
            panels[ancestors.Length] = panel;
            IReadOnlyList<IPdfBlock> children = ResolveContainerContentBlocks(container, style,
                currentOpts.MarginLeft + frame.X + style.PaddingX, frame.ContentWidth, currentOpts.DefaultFontSize);
            AppendFootnoteUnits(children, frame.X + style.PaddingX, frame.ContentWidth, panels, units);
            // Empty static panels still have a measurable decoration/padding box.
            if (start == units.Count) units.Add(new FootnoteUnit { Panels = panels });
            panel.End = units.Count;
            units[start].PanelBefore += style.SpacingBefore + style.GetFragmentTopPadding(isContinuation: false);
            units[panel.End - 1].PanelAfter += style.GetFragmentBottomPadding(continues: false) + style.SpacingAfter;
            if (style.KeepTogether) AddFootnoteKeepRange(units, start, panel.End, relaxIfOversized: false);
            return style.KeepWithNext;
        }

        // The same resolved ancestry defines both measured fragment edges and
        // painted fragment boxes; child callbacks and captured styles stay inert.
        private void PaintFootnoteUnits(IReadOnlyList<FootnoteUnit> units, int start, int count, ref double top,
            double measuredHeight = double.PositiveInfinity, IReadOnlyDictionary<int, TableFragmentPlan>? tables = null,
            bool continuedTable = false) {
            if (count == 0) return;
            int end = start + count;
            bool partialTable = tables != null && tables.TryGetValue(end - 1, out var lastTable) && lastTable.Continues;
            double clippedSpacing = Math.Max(0D, MeasureFootnoteFragmentHeight(units, start, count,
                double.PositiveInfinity, tables, continuedTable) - measuredHeight);
            var active = new List<FootnotePanelPaint>();
            double inset = 0D;
            for (int index = start; index < end;) {
                FootnoteUnit unit = units[index];
                for (int depth = active.Count; depth < unit.Panels.Length; depth++) {
                    FootnotePanel panel = unit.Panels[depth];
                    bool continuation = panel.Start < start || continuedTable && panel.Start == start;
                    if (!continuation) top -= panel.Style.SpacingBefore;
                    active.Add(new FootnotePanelPaint(panel, sb.Length, top, continuation));
                    top -= panel.GetTopPadding(continuation, tables);
                }
                int next = index + 1;
                if (unit.Block is RichParagraphBlock && unit.Line != null) {
                    while (next < end && ReferenceEquals(units[next].Block, unit.Block) &&
                        ReferenceEquals(units[next].Panels, unit.Panels) && units[next].Line != null &&
                        units[next].ParagraphLine == unit.ParagraphLine + next - index) next++;
                    PaintFootnoteParagraphFragment(units, index, next - index, ref top);
                } else if (unit.Table != null && tables != null && tables.TryGetValue(index, out var table)) {
                    PaintTableFragment(table, unit.FrameX, unit.FrameWidth, ref top);
                    foreach (FootnotePanelPaint paint in active) {
                        if (ReferenceEquals(paint.Panel.Container.FrameTable, unit.Block))
                            paint.TableContentBottom = top + table.After + table.Slices[table.Slices.Count - 1].GapAfter;
                    }
                } else PaintFootnoteUnit(unit, ref top);
                while (active.Count > 0 && (active[active.Count - 1].Panel.End == next || next == end)) {
                    FootnotePanelPaint paint = active[active.Count - 1];
                    active.RemoveAt(active.Count - 1);
                    FootnotePanel panel = paint.Panel;
                    bool continues = panel.End > end || partialTable && panel.End == end;
                    double bottom = panel.Style.GetFragmentBottomPadding(continues);
                    top -= bottom;
                    PaintFootnotePanel(paint, top, continues);
                    if (next == end) inset += Math.Max(0D, panel.Style.FragmentBottomInset - bottom);
                    if (!continues) {
                        double spacing = panel.Style.SpacingAfter;
                        if (next == end && panel.Style.ClipSpacingAfterAtFrameEnd) {
                            double clipped = Math.Min(spacing, clippedSpacing);
                            spacing -= clipped;
                            clippedSpacing -= clipped;
                        }
                        top -= spacing;
                    }
                }
                index = next;
            }
            top -= inset;
        }

        private void PaintFootnotePanel(FootnotePanelPaint paint, double bottom, bool continues) {
            bottom = ResolveContainerTableFrameBottom(paint.Panel.Container, paint.TableContentBottom, bottom, continues);
            double height = paint.Top - bottom;
            if (height <= .001D) return;
            FootnotePanel panel = paint.Panel;
            bool first = panel.Style.RepeatFragmentDecoration || !paint.IsContinuation;
            bool final = panel.Style.RepeatFragmentDecoration || !continues;
            var decoration = new StringBuilder();
            if (panel.Style.Background is { } background)
                DrawRoundedRowFill(decoration, background, currentOpts.MarginLeft + panel.X, bottom, panel.Width,
                    height, panel.Style.CornerRadius, first, first, final, final, emitGeneratedStructure);
            DrawPanelBorder(decoration, panel.Style, currentOpts.MarginLeft + panel.X, bottom, panel.Width, height,
                emitGeneratedStructure, first, final);
            DrawContainerTableFrame(decoration, panel.Container, currentOpts.MarginLeft + panel.X, panel.Width,
                bottom, height, first, final);
            if (decoration.Length > 0) sb.Insert(paint.Insertion, decoration.ToString());
        }
    }
}
