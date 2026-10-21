using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Drawing;
using System.Globalization;

namespace OfficeIMO.Word;

public sealed partial class WordPaperStyle {
    /// <summary>Fits relative column widths within the current table/container width while reserving room for each header's longest word and horizontal padding.</summary>
    /// <param name="table">A populated table with one unmerged header cell per column. Apply header typography before sizing.</param>
    /// <param name="relativeWidths">Non-negative column weights; omitted weights give equal desired widths. The values need not sum to 100.</param>
    /// <remarks>
    /// Words break at whitespace, not hyphens, so compound labels remain intact. Measurements use
    /// OfficeIMO.Core estimates, bounded below by available managed TrueType metrics. The resulting
    /// fixed grid and cell widths are in twips. Reapply after changing header text, typography or page geometry.
    /// </remarks>
    /// <exception cref="InvalidOperationException">The header minima exceed the available width, or the table has merged or irregular rows.</exception>
    public void FitTableColumnsToHeaders(WordTable table, params int[] relativeWidths) {
        if (table == null) throw new ArgumentNullException(nameof(table));
        if (relativeWidths == null) throw new ArgumentNullException(nameof(relativeWidths));
        List<WordTableRow> rows = table.Rows;
        int count = rows.Count == 0 ? 0 : rows[0].Cells.Count;
        if (count == 0) throw new InvalidOperationException("Header-aware sizing requires a populated header row.");
        if (rows.Any(row => row.Cells.Count != count || row.Cells.Any(cell =>
                cell._tableCellProperties?.GridSpan != null || cell._tableCellProperties?.VerticalMerge != null
                || cell._tableCellProperties?.HorizontalMerge != null))) {
            throw new InvalidOperationException("Header-aware sizing requires rectangular, unmerged table rows.");
        }
        int[] weights = relativeWidths.Length == 0 ? Enumerable.Repeat(1, count).ToArray() : relativeWidths;
        if (weights.Length != count) throw new ArgumentException($"Expected {count} column weights.", nameof(relativeWidths));
        if (weights.Any(weight => weight < 0) || weights.All(weight => weight == 0)) {
            throw new ArgumentException("Column weights must be non-negative with at least one positive weight.", nameof(relativeWidths));
        }

        int available = Math.Min(table.EstimateTableWidthInDxa(), table.EstimateAvailableContainerWidthInDxa());
        var measurer = OfficeTextMeasurer.Create();
        int[] minima = rows[0].Cells.Select(cell => MeasureHeaderMinimum(cell, table, measurer)).ToArray();
        long required = minima.Sum(width => (long)width);
        if (required > available) {
            throw new InvalidOperationException(string.Format(CultureInfo.InvariantCulture,
                "Header words and horizontal padding require {0:0.##} pt, but the table has only {1:0.##} pt available. Widen the page/table, reduce header typography or padding, or shorten the headers.",
                required / 20D, available / 20D));
        }
        int[] widths = AllocateHeaderWidths(available, minima, weights);
        // Keep the serialized preferred width, fixed layout, grid and every row in agreement.
        table.WidthType = WordTableWidthUnit.Dxa;
        table.Width = available;
        table.ColumnWidthType = WordTableWidthUnit.Dxa;
        table.ColumnWidth = widths.ToList();
        table._tableProperties!.TableLayout = new TableLayout { Type = TableLayoutValues.Fixed };
        table.RefreshGrid();
    }

    private int MeasureHeaderMinimum(WordTableCell cell, WordTable table, OfficeTextMeasurer measurer) {
        double longest = 0D;
        foreach (Paragraph paragraph in cell._tableCell.Elements<Paragraph>()) {
            double wordWidth = 0D;
            foreach (Run run in paragraph.Descendants<Run>()) {
                var scoped = new WordParagraph(table.Document, paragraph, run);
                var inherited = WordRunFormatting.EnumerateRunStyleProperties(scoped).ToArray();
                string family = scoped.FontFamily ?? inherited.Select(style => style.RunFonts?.Ascii?.Value)
                    .FirstOrDefault(value => !string.IsNullOrEmpty(value)) ?? FontFamily;
                double size = scoped.FontSizePoints ?? inherited.Select(style => style.FontSize?.Val?.Value)
                    .Select(value => double.TryParse(value, NumberStyles.Number, CultureInfo.InvariantCulture, out double halfPoints)
                        ? (double?)(halfPoints / 2D) : null).FirstOrDefault(value => value.HasValue) ?? 9D;
                bool bold = (run.RunProperties?.Bold == null ? (bool?)null : scoped.Bold) ?? inherited.Select(style => style.Bold == null ? (bool?)null : style.Bold.Val?.Value ?? true)
                    .FirstOrDefault(value => value.HasValue) ?? false;
                bool italic = (run.RunProperties?.Italic == null ? (bool?)null : scoped.Italic) ?? inherited.Select(style => style.Italic == null ? (bool?)null : style.Italic.Val?.Value ?? true)
                    .FirstOrDefault(value => value.HasValue) ?? false;
                var font = new OfficeFontInfo(family, size, (bold ? OfficeFontStyle.Bold : 0) | (italic ? OfficeFontStyle.Italic : 0));
                OfficeTextMeasurementStyle measurement = measurer.CreateStyle(font, 72D);
                OfficeTrueTypeFont? metrics = OfficeTrueTypeFont.TryLoadFontFamilyForText(family, font.Face, run.InnerText, out _);
                foreach (OpenXmlElement content in run.ChildElements) {
                    if (content is Break || content is TabChar) {
                        longest = Math.Max(longest, wordWidth);
                        wordWidth = 0D;
                    } else if (content is Text text) {
                        var segment = new System.Text.StringBuilder();
                        foreach (char character in text.Text) {
                            if (char.IsWhiteSpace(character) && character != '\u00a0' && character != '\u202f') {
                                wordWidth += MeasureSegment(segment.ToString());
                                segment.Clear();
                                longest = Math.Max(longest, wordWidth);
                                wordWidth = 0D;
                            } else {
                                segment.Append(character);
                            }
                        }
                        wordWidth += MeasureSegment(segment.ToString());
                    }
                }
                double MeasureSegment(string text) => Math.Max(measurer.MeasureWidth(text, measurement),
                    metrics?.Measure(text, size) ?? 0D);
            }
            longest = Math.Max(longest, wordWidth);
        }
        int left = cell.MarginLeftWidth ?? table.StyleDetails?.MarginDefaultLeftWidth ?? TableCellHorizontalPadding;
        int right = cell.MarginRightWidth ?? table.StyleDetails?.MarginDefaultRightWidth ?? TableCellHorizontalPadding;
        return Math.Max(1, checked((int)Math.Ceiling(longest * 20D) + Math.Max(0, left) + Math.Max(0, right)));
    }

    // Clamp undersized columns, then redistribute the remaining width in the requested ratio.
    private static int[] AllocateHeaderWidths(int available, int[] minima, int[] weights) {
        var widths = new int[minima.Length];
        var pending = new HashSet<int>(Enumerable.Range(0, minima.Length));
        int remaining = available;
        while (pending.Count > 0) {
            double totalWeight = pending.Sum(index => (double)weights[index]);
            int[] clamped = pending.Where(index => (totalWeight == 0D ? remaining / (double)pending.Count
                : remaining * (double)weights[index] / totalWeight) < minima[index]).ToArray();
            if (clamped.Length == 0) {
                int[] indices = pending.OrderBy(index => index).ToArray();
                int assigned = 0;
                foreach (int index in indices) {
                    widths[index] = (int)Math.Floor(totalWeight == 0D ? remaining / (double)indices.Length
                        : remaining * (double)weights[index] / totalWeight);
                    assigned += widths[index];
                }
                for (int index = 0; index < remaining - assigned; index++) widths[indices[index % indices.Length]]++;
                break;
            }
            foreach (int index in clamped) {
                widths[index] = minima[index];
                remaining -= minima[index];
                pending.Remove(index);
            }
        }
        return widths;
    }
}
