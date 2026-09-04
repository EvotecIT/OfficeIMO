namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private double[] ResolveRowColumnWidths(RowBlock row, double availableWidth) {
            int count = row.Columns.Count;
            var result = new double[count];
            if (count == 0 || availableWidth <= 0D) {
                return result;
            }

            bool percentagesOnly = true;
            double percentTotal = 0D;
            for (int index = 0; index < count; index++) {
                PdfColumnWidth width = row.Columns[index].Width;
                percentagesOnly &= width.Unit == PdfColumnWidthUnit.Percent;
                if (width.Unit == PdfColumnWidthUnit.Percent) {
                    percentTotal += width.Value;
                }
            }

            if (percentTotal > 100.0001D) {
                throw new ArgumentException("Percentage row columns cannot exceed 100% of the available width.");
            }

            if (percentagesOnly) {
                double scale = percentTotal <= 0D ? 0D : 100D / percentTotal;
                for (int index = 0; index < count; index++) {
                    result[index] = availableWidth * (row.Columns[index].Width.Value * scale / 100D);
                }

                return result;
            }

            double committed = 0D;
            double relativeWeight = 0D;
            var automaticIndexes = new List<int>();
            for (int index = 0; index < count; index++) {
                PdfColumnWidth width = row.Columns[index].Width;
                switch (width.Unit) {
                    case PdfColumnWidthUnit.Points:
                        result[index] = width.Value;
                        committed += result[index];
                        break;
                    case PdfColumnWidthUnit.Percent:
                        result[index] = availableWidth * width.Value / 100D;
                        committed += result[index];
                        break;
                    case PdfColumnWidthUnit.Auto:
                        double preferred = MeasureRowColumnPreferredWidth(row.Columns[index], availableWidth);
                        result[index] = Math.Max(width.Minimum, Math.Min(width.Maximum ?? availableWidth, preferred));
                        committed += result[index];
                        automaticIndexes.Add(index);
                        break;
                    case PdfColumnWidthUnit.Relative:
                        relativeWeight += width.Value;
                        break;
                    default:
                        throw new ArgumentOutOfRangeException(nameof(row), width.Unit, "Unsupported PDF row column width unit.");
                }
            }

            double overflow = committed - availableWidth;
            if (overflow > 0.001D && automaticIndexes.Count > 0) {
                for (int autoIndex = automaticIndexes.Count - 1; autoIndex >= 0 && overflow > 0.001D; autoIndex--) {
                    int index = automaticIndexes[autoIndex];
                    double minimum = row.Columns[index].Width.Minimum;
                    double reduction = Math.Min(overflow, result[index] - minimum);
                    result[index] -= reduction;
                    committed -= reduction;
                    overflow -= reduction;
                }
            }

            if (committed > availableWidth + 0.001D) {
                throw new ArgumentException("Fixed, automatic, and percentage row columns exceed the available row width.");
            }

            double remaining = Math.Max(0D, availableWidth - committed);
            if (relativeWeight > 0D && remaining <= 0.001D) {
                throw new ArgumentException("Relative row columns require positive width after fixed, automatic, and percentage columns are resolved.");
            }

            if (relativeWeight > 0D) {
                for (int index = 0; index < count; index++) {
                    PdfColumnWidth width = row.Columns[index].Width;
                    if (width.Unit == PdfColumnWidthUnit.Relative) {
                        result[index] = remaining * width.Value / relativeWeight;
                    }
                }
            }

            return result;
        }

        private double MeasureRowColumnPreferredWidth(RowColumn column, double availableWidth) {
            double preferred = 0D;
            foreach (IPdfBlock block in column.Blocks) {
                preferred = Math.Max(preferred, MeasureBlockPreferredWidth(block, availableWidth));
            }

            return Math.Min(availableWidth, Math.Max(1D, preferred));
        }

        private double MeasureBlockPreferredWidth(IPdfBlock block, double availableWidth) {
            switch (block) {
                case HeadingBlock heading:
                    PdfHeadingStyle headingStyle = heading.Style ?? currentOpts.DefaultHeadingStylesSnapshot?.GetSnapshot(heading.Level) ?? new PdfHeadingStyle();
                    double headingSize = headingStyle.GetFontSize(heading.Level);
                    return EstimateSimpleTextWidthForOptions(heading.Text, headingStyle.Font ?? currentOpts.DefaultFont, headingSize, currentOpts);
                case RichParagraphBlock paragraph:
                    return MeasureRunsPreferredWidth(paragraph.Runs);
                case PanelParagraphBlock panel:
                    PdfPanelStyle panelStyle = ResolvePanelStyle(panel, currentOpts);
                    double panelWidth = MeasureRunsPreferredWidth(panel.Runs) + panelStyle.PaddingX * 2D;
                    return panelStyle.MaxWidth.HasValue
                        ? Math.Min(panelStyle.MaxWidth.Value, panelWidth)
                        : panelWidth;
                case BulletListBlock bullets:
                    return bullets.RichItems.Count == 0 ? 1D : bullets.RichItems.Max(item => MeasureRunsPreferredWidth(item.Runs)) + currentOpts.DefaultFontSize * 1.5D;
                case NumberedListBlock numbered:
                    return numbered.RichItems.Count == 0 ? 1D : numbered.RichItems.Max(item => MeasureRunsPreferredWidth(item.Runs)) + currentOpts.DefaultFontSize * 2D;
                case ImageBlock image:
                    return image.Width;
                case ShapeBlock shape:
                    return shape.Shape.Width;
                case DrawingBlock drawing:
                    return drawing.Drawing.Width;
                case TextFieldBlock textField:
                    return textField.Width;
                case ChoiceFieldBlock choiceField:
                    return choiceField.Width;
                case CheckBoxBlock checkBox:
                    return checkBox.Size;
                case RadioButtonGroupBlock radioButtons:
                    return radioButtons.Options.Max(option => EstimateSimpleTextWidthForOptions(option, currentOpts.DefaultFont, currentOpts.DefaultFontSize, currentOpts)) + radioButtons.Size + radioButtons.Gap;
                case TextAnnotationBlock annotation:
                    return annotation.Width;
                case FreeTextAnnotationBlock annotation:
                    return annotation.Width;
                case HighlightAnnotationBlock annotation:
                    return annotation.Width;
                case ContainerBlock container:
                    double contentWidth = container.Blocks.Count == 0 ? 1D : container.Blocks.Max(item => MeasureBlockPreferredWidth(item, availableWidth));
                    return contentWidth + container.Style.PaddingX * 2D;
                case FlowBlock flow when flow.StaticBlocks != null:
                    return flow.StaticBlocks.Count == 0 ? 1D : flow.StaticBlocks.Max(item => MeasureBlockPreferredWidth(item, availableWidth));
                case RowBlock:
                case TableBlock:
                case DeferredTableBlock:
                    return availableWidth;
                default:
                    return Math.Min(availableWidth, currentOpts.DefaultFontSize * 4D);
            }
        }

        private double MeasureRunsPreferredWidth(IReadOnlyList<PdfTextRun> runs) {
            double width = 0D;
            foreach (PdfTextRun run in runs) {
                if (run.InlineElement != null) {
                    width += run.InlineElement.Width;
                    continue;
                }

                double size = run.FontSize ?? currentOpts.DefaultFontSize;
                width += EstimateSimpleTextWidthForOptions(run.Text, run.Font ?? currentOpts.DefaultFont, size, currentOpts);
            }

            return width;
        }
    }
}
