using System;
using System.Linq;
using DocumentFormat.OpenXml.Drawing;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Represents a cell within a PowerPoint table.
    /// </summary>
    public class PowerPointTableCell {
        private const int EmusPerPoint = 12700;

        internal PowerPointTableCell(TableCell cell) {
            Cell = cell;
        }

        internal TableCell Cell { get; }

        /// <summary>
        ///     Gets or sets the text contained in the cell.
        /// </summary>
        public string Text {
            get => Cell.TextBody?.InnerText ?? string.Empty;


            set {
                Cell.TextBody ??= new A.TextBody(new A.BodyProperties(), new A.ListStyle());
                A.Paragraph paragraph = Cell.TextBody.GetFirstChild<A.Paragraph>() ?? new A.Paragraph();
                Cell.TextBody.RemoveAllChildren<A.Paragraph>();
                paragraph.RemoveAllChildren<A.Run>();
                paragraph.Append(new A.Run(new A.Text(value ?? string.Empty)));
                Cell.TextBody.Append(paragraph);
            }
        }


        /// <summary>
        ///     Gets or sets the merge information for this cell.
        ///     Tuple is in format (rows, columns).
        /// </summary>
        public (int rows, int columns) Merge {
            get {
                int rows = Cell.RowSpan?.Value ?? 1;
                int cols = Cell.GridSpan?.Value ?? 1;
                return (rows, cols);
            }
            set {
                if (value.rows <= 1) {
                    Cell.RowSpan = null;
                } else {
                    Cell.RowSpan = value.rows;
                }

                if (value.columns <= 1) {
                    Cell.GridSpan = null;
                } else {
                    Cell.GridSpan = value.columns;
                }
            }
        }

        /// <summary>
        ///     Gets or sets the horizontal alignment of the cell text.
        /// </summary>
        public A.TextAlignmentTypeValues? HorizontalAlignment {
            get {
                var pPr = Cell.TextBody?.Elements<Paragraph>().FirstOrDefault()?.ParagraphProperties;
                return pPr?.Alignment?.Value;
            }
            set {
                Cell.TextBody ??= new A.TextBody(new A.BodyProperties(), new A.ListStyle(), new A.Paragraph());
                var paragraph = Cell.TextBody.Elements<A.Paragraph>().First();
                paragraph.ParagraphProperties ??= new A.ParagraphProperties();
                paragraph.ParagraphProperties.Alignment = value;
            }
        }


        /// <summary>
        ///     Gets or sets the fill color of the cell in hex format (e.g. "FF0000").
        /// </summary>
        public string? FillColor {
            get {
                SolidFill? solid = Cell.TableCellProperties?.GetFirstChild<SolidFill>();
                return solid?.RgbColorModelHex?.Val;
            }


            set {
                Cell.TableCellProperties ??= new TableCellProperties();
                Cell.TableCellProperties.RemoveAllChildren<SolidFill>();
                if (value != null) {
                    Cell.TableCellProperties.Append(new SolidFill(new RgbColorModelHex { Val = value }));
                }
            }
        }

        /// <summary>
        ///     Gets or sets the border color (all sides) in hex format.        
        /// </summary>
        public string? BorderColor {
            get {
                var line = Cell.TableCellProperties?.LeftBorderLineProperties;
                return GetLineColor(line);
            }
            set {
                if (value == null) {
                    ClearBorders(TableCellBorders.All);
                    return;
                }

                SetBorders(TableCellBorders.All, value);
            }
        }

        /// <summary>
        ///     Gets a value indicating whether this cell is a merged continuation cell.
        /// </summary>
        public bool IsMergedCell =>
            Cell.HorizontalMerge?.Value == true || Cell.VerticalMerge?.Value == true;

        /// <summary>
        ///     Gets a value indicating whether this cell is the anchor for a merged range.
        /// </summary>
        public bool IsMergeAnchor =>
            (Cell.RowSpan?.Value ?? 1) > 1 || (Cell.GridSpan?.Value ?? 1) > 1;

        /// <summary>
        ///     Replaces text within the cell while preserving run formatting.
        /// </summary>
        public int ReplaceText(string oldValue, string newValue) {
            if (oldValue == null) {
                throw new ArgumentNullException(nameof(oldValue));
            }
            if (oldValue.Length == 0) {
                throw new ArgumentException("Old value cannot be empty.", nameof(oldValue));
            }

            string replacement = newValue ?? string.Empty;
            int count = 0;

            if (Cell.TextBody == null) {
                return 0;
            }

            foreach (A.Paragraph paragraph in Cell.TextBody.Elements<A.Paragraph>()) {
                foreach (A.Run run in paragraph.Elements<A.Run>()) {
                    foreach (A.Text text in run.Elements<A.Text>()) {
                        string current = text.Text ?? string.Empty;
                        int occurrences = CountOccurrences(current, oldValue);
                        if (occurrences == 0) {
                            continue;
                        }

                        text.Text = current.Replace(oldValue, replacement);
                        count += occurrences;
                    }
                }
            }

            return count;
        }
        /// <summary>
        ///     Gets or sets whether the cell text is bold.
        /// </summary>
        public bool Bold {
            get => GetRun()?.RunProperties?.Bold?.Value == true;
            set {
                var props = EnsureRunProperties();
                props.Bold = value ? true : null;
            }
        }

        /// <summary>
        ///     Gets or sets whether the cell text is italic.
        /// </summary>
        public bool Italic {
            get => GetRun()?.RunProperties?.Italic?.Value == true;
            set {
                var props = EnsureRunProperties();
                props.Italic = value ? true : null;
            }
        }

        /// <summary>
        ///     Gets or sets the font size in points.
        /// </summary>
        public int? FontSize {
            get {
                int? size = GetRun()?.RunProperties?.FontSize?.Value;
                return size != null ? size / 100 : null;
            }
            set {
                var props = EnsureRunProperties();
                props.FontSize = value != null ? value * 100 : null;
            }
        }

        /// <summary>
        ///     Gets or sets the font name.
        /// </summary>
        public string? FontName {
            get => GetRun()?.RunProperties?.GetFirstChild<A.LatinFont>()?.Typeface;
            set {
                var props = EnsureRunProperties();
                props.RemoveAllChildren<A.LatinFont>();
                if (value != null) {
                    props.Append(new A.LatinFont { Typeface = value });
                }
            }
        }

        /// <summary>
        ///     Gets or sets the text color in hex format (e.g. "FF0000").
        /// </summary>
        public string? Color {
            get => GetRun()?.RunProperties?.GetFirstChild<A.SolidFill>()?.RgbColorModelHex?.Val;
            set {
                var props = EnsureRunProperties();
                var latin = props.GetFirstChild<A.LatinFont>();
                var ea = props.GetFirstChild<A.EastAsianFont>();
                var cs = props.GetFirstChild<A.ComplexScriptFont>();

                props.RemoveAllChildren<A.SolidFill>();
                props.RemoveAllChildren<A.LatinFont>();
                props.RemoveAllChildren<A.EastAsianFont>();
                props.RemoveAllChildren<A.ComplexScriptFont>();

                if (value != null) {
                    props.Append(new A.SolidFill(new A.RgbColorModelHex { Val = value }));
                }

                if (latin != null) props.Append((A.LatinFont)latin.CloneNode(true));
                if (ea != null) props.Append((A.EastAsianFont)ea.CloneNode(true));
                if (cs != null) props.Append((A.ComplexScriptFont)cs.CloneNode(true));
            }
        }

        /// <summary>
        ///     Gets or sets the left padding in points.
        /// </summary>
        public double? PaddingLeftPoints {
            get => FromEmus(Cell.TableCellProperties?.LeftMargin?.Value);
            set {
                TableCellProperties props = EnsureProperties();
                props.LeftMargin = ToEmus(value);
            }
        }

        /// <summary>
        ///     Gets or sets the right padding in points.
        /// </summary>
        public double? PaddingRightPoints {
            get => FromEmus(Cell.TableCellProperties?.RightMargin?.Value);
            set {
                TableCellProperties props = EnsureProperties();
                props.RightMargin = ToEmus(value);
            }
        }

        /// <summary>
        ///     Gets or sets the top padding in points.
        /// </summary>
        public double? PaddingTopPoints {
            get => FromEmus(Cell.TableCellProperties?.TopMargin?.Value);
            set {
                TableCellProperties props = EnsureProperties();
                props.TopMargin = ToEmus(value);
            }
        }

        /// <summary>
        ///     Gets or sets the bottom padding in points.
        /// </summary>
        public double? PaddingBottomPoints {
            get => FromEmus(Cell.TableCellProperties?.BottomMargin?.Value);
            set {
                TableCellProperties props = EnsureProperties();
                props.BottomMargin = ToEmus(value);
            }
        }

        // VerticalAlignment is supported through TableCellProperties.Anchor.   

        /// <summary>
        ///     Gets or sets the vertical alignment of the cell text (top/center/bottom).
        /// </summary>
        public A.TextAnchoringTypeValues? VerticalAlignment {
            get => Cell.TableCellProperties?.Anchor?.Value;
            set {
                Cell.TableCellProperties ??= new TableCellProperties();
                Cell.TableCellProperties.Anchor = value;
            }
        }

        /// <summary>
        ///     Applies border styling to the specified sides.
        /// </summary>
        public void SetBorders(TableCellBorders borders, string color, double? widthPoints = null) {
            if (string.IsNullOrWhiteSpace(color)) {
                throw new ArgumentException("Border color cannot be null or empty.", nameof(color));
            }

            TableCellProperties props = EnsureProperties();

            if (borders.HasFlag(TableCellBorders.Left)) {
                props.LeftBorderLineProperties ??= new LeftBorderLineProperties();
                ApplyLine(props.LeftBorderLineProperties, color, widthPoints);
            }
            if (borders.HasFlag(TableCellBorders.Top)) {
                props.TopBorderLineProperties ??= new TopBorderLineProperties();
                ApplyLine(props.TopBorderLineProperties, color, widthPoints);
            }
            if (borders.HasFlag(TableCellBorders.Right)) {
                props.RightBorderLineProperties ??= new RightBorderLineProperties();
                ApplyLine(props.RightBorderLineProperties, color, widthPoints);
            }
            if (borders.HasFlag(TableCellBorders.Bottom)) {
                props.BottomBorderLineProperties ??= new BottomBorderLineProperties();
                ApplyLine(props.BottomBorderLineProperties, color, widthPoints);
            }
        }

        /// <summary>
        ///     Clears borders on the specified sides.
        /// </summary>
        public void ClearBorders(TableCellBorders borders) {
            TableCellProperties props = EnsureProperties();

            if (borders.HasFlag(TableCellBorders.Left)) {
                props.LeftBorderLineProperties = null;
            }
            if (borders.HasFlag(TableCellBorders.Top)) {
                props.TopBorderLineProperties = null;
            }
            if (borders.HasFlag(TableCellBorders.Right)) {
                props.RightBorderLineProperties = null;
            }
            if (borders.HasFlag(TableCellBorders.Bottom)) {
                props.BottomBorderLineProperties = null;
            }
        }

        private static void ApplyLine(LinePropertiesType line, string color, double? widthPoints) {
            line.RemoveAllChildren<SolidFill>();
            line.Append(new SolidFill(new RgbColorModelHex { Val = color }));
            if (widthPoints != null) {
                line.Width = (int)Math.Round(widthPoints.Value * EmusPerPoint);
            }
        }

        private static string? GetLineColor(LinePropertiesType? line) {
            return line?.GetFirstChild<SolidFill>()?.RgbColorModelHex?.Val;
        }

        private static int? ToEmus(double? points) {
            return points != null ? (int)Math.Round(points.Value * EmusPerPoint) : null;
        }

        private static double? FromEmus(int? emus) {
            return emus != null ? emus.Value / (double)EmusPerPoint : null;
        }

        private static int CountOccurrences(string value, string oldValue) {
            int count = 0;
            int index = 0;
            while (true) {
                index = value.IndexOf(oldValue, index, StringComparison.Ordinal);
                if (index < 0) {
                    break;
                }
                count++;
                index += oldValue.Length;
            }
            return count;
        }

        private TableCellProperties EnsureProperties() {
            return Cell.TableCellProperties ??= new TableCellProperties();
        }

        private A.Run? GetRun() {
            return Cell.TextBody?
                .Elements<A.Paragraph>()
                .FirstOrDefault()?
                .Elements<A.Run>()
                .FirstOrDefault();
        }

        private A.Run EnsureRun() {
            Cell.TextBody ??= new A.TextBody(new A.BodyProperties(), new A.ListStyle());
            A.Paragraph paragraph = Cell.TextBody.Elements<A.Paragraph>().FirstOrDefault() ?? new A.Paragraph();
            if (paragraph.Parent == null) {
                Cell.TextBody.Append(paragraph);
            }

            A.Run run = paragraph.Elements<A.Run>().FirstOrDefault() ?? new A.Run(new A.Text(string.Empty));
            if (run.Parent == null) {
                paragraph.Append(run);
            }

            return run;
        }

        private A.RunProperties EnsureRunProperties() {
            A.Run run = EnsureRun();
            return run.RunProperties ??= new A.RunProperties();
        }
    }
}
