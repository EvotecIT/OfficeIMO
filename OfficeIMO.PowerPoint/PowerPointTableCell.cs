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

        private TableCellProperties EnsureProperties() {
            return Cell.TableCellProperties ??= new TableCellProperties();
        }
    }
}
