using DocumentFormat.OpenXml.Drawing;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Represents a cell within a PowerPoint table.
    /// </summary>
    public class PowerPointTableCell {
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
                var ln = Cell.TableCellProperties?.GetFirstChild<Outline>();
                return ln?.GetFirstChild<SolidFill>()?.RgbColorModelHex?.Val;
            }
            set {
                Cell.TableCellProperties ??= new TableCellProperties();
                var outline = Cell.TableCellProperties.GetFirstChild<Outline>();
                if (value == null) {
                    outline?.Remove();
                    return;
                }
                if (outline == null) {
                    outline = new Outline();
                    Cell.TableCellProperties.Append(outline);
                }
                outline.RemoveAllChildren<SolidFill>();
                outline.Append(new SolidFill(new RgbColorModelHex { Val = value }));
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
    }
}
