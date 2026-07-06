using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    internal static partial class WordDocumentImageRenderer {
        private static IReadOnlyList<WordImageColumnFrame> CreateColumnFrames(WordSection? section, double contentLeft, double contentWidth) {
            if (section == null) {
                return Array.Empty<WordImageColumnFrame>();
            }

            int columnCount = GetSectionColumnCount(section);
            if (columnCount <= 1) {
                return Array.Empty<WordImageColumnFrame>();
            }

            double gap = GetSectionColumnGap(section);
            double availableWidth = Math.Max(1D, contentWidth - (gap * Math.Max(0, columnCount - 1)));
            IReadOnlyList<double> widthPercents = GetSectionColumnWidthPercents(section, columnCount);
            var frames = new List<WordImageColumnFrame>(widthPercents.Count);
            double currentLeft = contentLeft;
            for (int index = 0; index < widthPercents.Count; index++) {
                double columnWidth = index == widthPercents.Count - 1
                    ? Math.Max(1D, contentLeft + contentWidth - currentLeft)
                    : Math.Max(1D, availableWidth * widthPercents[index] / 100D);
                frames.Add(new WordImageColumnFrame(currentLeft, columnWidth));
                currentLeft += columnWidth + gap;
            }

            return frames.Count > 1 ? frames : Array.Empty<WordImageColumnFrame>();
        }

        private static int GetSectionColumnCount(WordSection section) {
            int? explicitColumnCount = section._sectionProperties
                .GetFirstChild<Columns>()?
                .Elements<Column>()
                .Count();
            int count = section.ColumnCount ?? explicitColumnCount ?? 1;
            return Math.Min(Math.Max(1, count), 8);
        }

        private static IReadOnlyList<double> GetSectionColumnWidthPercents(WordSection section, int columnCount) {
            List<int>? explicitWidths = GetExplicitSectionColumnWidths(section, columnCount);
            if (explicitWidths == null || explicitWidths.Count == 0) {
                return CreateEqualColumnWidthPercents(columnCount);
            }

            int total = explicitWidths.Sum();
            if (total <= 0) {
                return CreateEqualColumnWidthPercents(columnCount);
            }

            var widths = new List<double>(explicitWidths.Count);
            double accumulated = 0D;
            for (int index = 0; index < explicitWidths.Count; index++) {
                double percent = index == explicitWidths.Count - 1
                    ? 100D - accumulated
                    : explicitWidths[index] * 100D / total;
                widths.Add(percent);
                accumulated += percent;
            }

            return widths;
        }

        private static List<double> CreateEqualColumnWidthPercents(int columnCount) {
            var widths = new List<double>(columnCount);
            for (int index = 0; index < columnCount; index++) {
                widths.Add(100D / columnCount);
            }

            return widths;
        }

        private static List<int>? GetExplicitSectionColumnWidths(WordSection section, int columnCount) {
            Columns? columns = section._sectionProperties.GetFirstChild<Columns>();
            if (columns == null) {
                return null;
            }

            var widths = new List<int>(columnCount);
            foreach (Column column in columns.Elements<Column>().Take(columnCount)) {
                if (!TryParseTwips(column.Width?.Value, out int width) || width <= 0) {
                    return null;
                }

                widths.Add(width);
            }

            return widths.Count == columnCount ? widths : null;
        }

        private static double GetSectionColumnGap(WordSection section) {
            double? gap = section.ColumnsSpace.HasValue ? section.ColumnsSpace.Value / TwipsPerPoint : null;
            if (!gap.HasValue) {
                Column? firstColumn = section._sectionProperties.GetFirstChild<Columns>()?.Elements<Column>().FirstOrDefault();
                if (TryParseTwips(firstColumn?.Space?.Value, out int columnGap)) {
                    gap = columnGap / TwipsPerPoint;
                }
            }

            return !gap.HasValue || gap.Value < 0D ? 36D : gap.Value;
        }

        private static bool TryParseTwips(string? value, out int twips) =>
            int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out twips);

        private readonly struct WordImageColumnFrame {
            internal WordImageColumnFrame(double left, double width) {
                Left = left;
                Width = width;
            }

            internal double Left { get; }

            internal double Width { get; }
        }
    }
}
