using DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// Represents a cell within a PowerPoint table.
    /// </summary>
    public class PowerPointTableCell {
        internal TableCell Cell { get; }

        internal PowerPointTableCell(TableCell cell) {
            Cell = cell;
        }

        /// <summary>
        /// Gets or sets the text contained in the cell.
        /// </summary>
        public string Text {
            get {


                return Cell.TextBody?.InnerText ?? string.Empty;


            }


            set {


                Cell.TextBody ??= new TextBody(new BodyProperties(), new ListStyle());


                Paragraph paragraph = Cell.TextBody.GetFirstChild<Paragraph>() ?? new Paragraph();


                Cell.TextBody.RemoveAllChildren<Paragraph>();


                paragraph.RemoveAllChildren<Run>();


                paragraph.Append(new Run(new Text(value ?? string.Empty)));


                Cell.TextBody.Append(paragraph);


            }


        }





        /// <summary>


        /// Gets or sets the merge information for this cell.


        /// Tuple is in format (rows, columns).


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


        /// Gets or sets the fill color of the cell in hex format (e.g. "FF0000").


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


    }


}


