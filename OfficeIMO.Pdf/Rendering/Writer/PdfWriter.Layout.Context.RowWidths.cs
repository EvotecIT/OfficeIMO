namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private double[] ResolveRowColumnWidths(RowBlock row, double availableWidth) {
            int count = row.Columns.Count;
            var result = new double[count];
            if (count == 0 || availableWidth <= 0D) {
                return result;
            }

            double percentTotal = 0D;
            for (int index = 0; index < count; index++) {
                PdfColumnWidth width = row.Columns[index].Width;
                if (width.Unit == PdfColumnWidthUnit.Percent) {
                    percentTotal += width.Value;
                }
            }

            if (percentTotal > 100.0001D) {
                throw new ArgumentException("Percentage row columns cannot exceed 100% of the available width.");
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
                case ContainerBlock container:
                    double inset = ResolveContainerStyle(container).PaddingX * 2D;
                    double content = container.Blocks.Count == 0 ? 0D : container.Blocks.Max(child => MeasureBlockPreferredWidth(child, Math.Max(0D, availableWidth - inset)));
                    return Math.Min(ResolveContainerStyle(container).MaxWidth ?? availableWidth, content + inset);
                case SemanticBlock semantic:
                    return semantic.Blocks.Count == 0 ? 0D : semantic.Blocks.Max(child => MeasureBlockPreferredWidth(child, availableWidth));
                case FlowBlock flow when flow.StaticBlocks != null:
                    return flow.StaticBlocks.Count == 0 ? 0D : flow.StaticBlocks.Max(child => MeasureBlockPreferredWidth(child, availableWidth));
                case HeadingBlock heading:
                    PdfHeadingStyle? headingStyle = ResolveHeadingStyle(heading, currentOpts);
                    double headingSize = GetHeadingFontSize(heading, headingStyle);
                    return MeasureRunsPreferredWidth(
                        CreateHeadingTextRuns(heading, headingStyle, heading.Color ?? headingStyle?.Color),
                        headingSize);
                case RichParagraphBlock paragraph:
                    return MeasureRunsPreferredWidth(paragraph.Runs);
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
                case RowBlock:
                case TableBlock:
                case DeferredTableBlock:
                    return availableWidth;
                default:
                    return Math.Min(availableWidth, currentOpts.DefaultFontSize * 4D);
            }
        }

        private double MeasureRunsPreferredWidth(IReadOnlyList<PdfTextRun> runs, double? fontSize = null) {
            double size = fontSize ?? currentOpts.DefaultFontSize;
            var layout = WrapRichRunsCore(
                runs, double.MaxValue, size, ChooseNormal(currentOpts.DefaultFont), size * 1.4D,
                null, DefaultParagraphTabStopWidth, currentOpts);
            return layout.Lines.Count == 0 ? 0D : layout.Lines.Max(line => MeasureRichLineWidth(line, currentOpts));
        }
    }
}
