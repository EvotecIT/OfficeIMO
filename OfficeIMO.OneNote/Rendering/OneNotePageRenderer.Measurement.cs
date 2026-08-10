using OfficeIMO.Drawing;

namespace OfficeIMO.OneNote;

public static partial class OneNotePageRenderer {
    private sealed partial class RenderContext {
        private int _listMeasurementDepth;

        internal void MeasureWithTemporaryListNumbering(Action measure) {
            if (measure == null) throw new ArgumentNullException(nameof(measure));
            MeasureWithTemporaryListNumbering(() => {
                measure();
                return true;
            });
        }

        internal double MeasureElementHeight(OneNoteElement element, double width) =>
            MeasureWithTemporaryListNumbering(() => MeasureElementHeightCore(element, width));

        private double MeasureElementHeightCore(OneNoteElement element, double width) {
            if (element is OneNoteImage image) {
                if (!_options.IncludeImages) return 0D;
                if (TryIdentifyImage(image, out OfficeImageInfo? info)) {
                    double renderWidth = ResolveImageWidth(image, info, width);
                    return Math.Max(1D, ResolveImageHeight(image, info, renderWidth));
                }
                return ResolveImageHeight(image, null, width);
            }
            if (element.Layout?.Height.HasValue == true) {
                AdvanceListNumberingForElement(element);
                return Math.Max(1D, element.Layout.Height.Value * PointsPerHalfInch);
            }
            if (element is OneNoteParagraph paragraph) return MeasureParagraphHeight(paragraph, width);
            if (element is OneNoteOutline outline) {
                ResetListNumbering();
                return MeasureElementsBounds(outline.Children, width).Bottom;
            }
            if (element is OneNoteTable table) return MeasureTableHeight(table, width);
            if (element is OneNoteInk ink) {
                OfficeInkBounds bounds = ink.Ink.GetBounds();
                if (bounds.IsEmpty) return DefaultParagraphHeight;
                double sourceWidth = Math.Max(0.000001D, (bounds.X + bounds.Width) * PointsPerHalfInch);
                double fit = Math.Min(1D, Math.Max(1D, width) / sourceWidth);
                return Math.Max(DefaultParagraphHeight, (bounds.Y + bounds.Height) * PointsPerHalfInch * fit);
            }
            if (element is OneNoteMath math) {
                OfficeMathLayoutMetrics metrics = OfficeMathRenderer.Measure(math.GetExpression(), _options.Math);
                return Math.Max(DefaultParagraphHeight, metrics.Height);
            }
            if (element is OneNoteBinaryElement) return 34D;
            return 32D;
        }

        internal double MeasureElementsHeight(IEnumerable<OneNoteElement> elements, double width) =>
            MeasureElementsBounds(elements, width).Bottom;

        internal (double Right, double Bottom) MeasureElementsBounds(IEnumerable<OneNoteElement> elements, double width) =>
            MeasureWithTemporaryListNumbering(() => MeasureElementsBoundsCore(elements, width));

        private (double Right, double Bottom) MeasureElementsBoundsCore(IEnumerable<OneNoteElement> elements, double width) {
            double right = 0D;
            double bottom = 0D;
            double cursor = 0D;
            double pendingSpace = 0D;
            foreach (OneNoteElement element in elements) {
                bool participatesInFlow = element.Layout?.Y.HasValue != true;
                double elementX = element.Layout?.X.HasValue == true ? element.Layout.X.Value * PointsPerHalfInch : 0D;
                double elementY = element.Layout?.Y.HasValue == true
                    ? element.Layout.Y.Value * PointsPerHalfInch
                    : cursor + Math.Max(pendingSpace, ParagraphSpaceBefore(element));
                double remainingWidth = Math.Max(1D, width - Math.Max(0D, elementX));
                double elementWidth = ResolveEstimatedWidth(element, remainingWidth, _options);
                double elementHeight = MeasureElementHeight(element, elementWidth);
                double extentWidth = MeasureElementWidthExtent(element, elementWidth);
                right = Math.Max(right, elementX + extentWidth);
                bottom = Math.Max(bottom, elementY + elementHeight);
                if (participatesInFlow) {
                    cursor = Math.Max(cursor, elementY + elementHeight);
                    pendingSpace = element is OneNoteParagraph ? ParagraphSpaceAfter(element) : 6D;
                }
            }
            return (Math.Max(1D, right), Math.Max(DefaultParagraphHeight, Math.Max(bottom, cursor + pendingSpace)));
        }

        internal double MeasureElementWidthExtent(OneNoteElement element, double width) {
            if (element is OneNoteOutline outline) return Math.Max(width,
                MeasureWithoutChangingListNumbering(() => MeasureElementsBounds(outline.Children, width).Right));
            if (element is OneNoteParagraph paragraph) return Math.Max(width,
                MeasureWithoutChangingListNumbering(() => MeasureElementsBounds(paragraph.Children, width).Right));
            return width;
        }

        private double MeasureParagraphHeight(OneNoteParagraph paragraph, double width) {
            string prefix = CreateParagraphPrefix(paragraph, advanceListState: true);
            double textHeight;
            if (paragraph.Runs.Count == 0 && prefix.Length == 0) {
                textHeight = DefaultParagraphHeight;
            } else if (paragraph.Runs.Any(run => run.MathExpression != null)) {
                IReadOnlyList<InlineMathLine> lines = CreateInlineMathLines(paragraph, prefix, width);
                double exactLineHeight = ParagraphDistance(paragraph.Style.ExactLineSpacing);
                double height = lines.Sum(line => Math.Max(line.Height + 3D, exactLineHeight)) - 3D;
                textHeight = Math.Max(DefaultParagraphHeight, height);
            } else {
                IReadOnlyList<OfficeRichTextRun> runs = CreateParagraphRichTextRuns(paragraph, prefix);
                double fontSize = paragraph.Runs.Count == 0
                    ? _options.DefaultFont.Size
                    : paragraph.Runs.Max(run => run.Style.FontSize ?? _options.DefaultFont.Size);
                double lineHeight = ResolveParagraphLineHeight(paragraph, fontSize);
                textHeight = MeasureRichTextHeight(runs, width, lineHeight, CreateParagraphIndent(paragraph));
            }
            double cursor = textHeight;
            double bottom = textHeight;
            double pendingSpace = 0D;
            foreach (OneNoteElement child in paragraph.Children) {
                bool participatesInFlow = child.Layout?.Y.HasValue != true;
                double childX = child.Layout?.X.HasValue == true ? child.Layout.X.Value * PointsPerHalfInch : 0D;
                double childY = child.Layout?.Y.HasValue == true
                    ? child.Layout.Y.Value * PointsPerHalfInch
                    : cursor + Math.Max(pendingSpace, ParagraphSpaceBefore(child));
                double childWidth = ResolveEstimatedWidth(
                    child,
                    Math.Max(1D, width - Math.Max(0D, childX)),
                    _options);
                double childHeight = MeasureElementHeight(child, childWidth);
                bottom = Math.Max(bottom, childY + childHeight);
                if (participatesInFlow) {
                    cursor = Math.Max(cursor, childY + childHeight);
                    pendingSpace = child is OneNoteParagraph ? ParagraphSpaceAfter(child) : 5D;
                }
            }
            return Math.Max(DefaultParagraphHeight, Math.Max(bottom, cursor + pendingSpace));
        }

        private double MeasureTableHeight(OneNoteTable table, double width) {
            int columns = Math.Max(table.ColumnWidths.Count, table.Rows.Count == 0 ? 0 : table.Rows.Max(row => row.Cells.Count));
            if (columns == 0) return DefaultParagraphHeight;
            double[] columnWidths = ResolveTableColumns(table, columns, width);
            double height = 0D;
            foreach (OneNoteTableRow row in table.Rows) {
                double rowHeight = 32D;
                for (int column = 0; column < row.Cells.Count && column < columnWidths.Length; column++) {
                    rowHeight = Math.Max(rowHeight,
                        MeasureElementsHeight(row.Cells[column].Content, Math.Max(1D, columnWidths[column] - 8D)) + 8D);
                }
                height += rowHeight;
            }
            return Math.Max(DefaultParagraphHeight, height);
        }

        private T MeasureWithTemporaryListNumbering<T>(Func<T> measure) {
            bool ownsSnapshot = _listMeasurementDepth == 0;
            Dictionary<string, int>? snapshot = ownsSnapshot ? SnapshotListNumbering() : null;
            _listMeasurementDepth++;
            try {
                return measure();
            } finally {
                _listMeasurementDepth--;
                if (ownsSnapshot) RestoreListNumbering(snapshot!);
            }
        }

        private T MeasureWithoutChangingListNumbering<T>(Func<T> measure) {
            Dictionary<string, int> snapshot = SnapshotListNumbering();
            try {
                return measure();
            } finally {
                RestoreListNumbering(snapshot);
            }
        }
    }
}
