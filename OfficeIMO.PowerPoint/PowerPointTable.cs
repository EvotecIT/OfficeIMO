using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Represents a table on a slide.
    /// </summary>
    public partial class PowerPointTable : PowerPointShape {
        private const int EmusPerPoint = 12700;
        private readonly SlidePart? _slidePart;

        internal PowerPointTable(GraphicFrame frame, SlidePart? slidePart = null) : base(frame) {
            _slidePart = slidePart;
        }

        private GraphicFrame Frame => (GraphicFrame)Element;
        internal SlidePart? SlidePart => _slidePart;
        internal A.Table TableElement => Frame.Graphic!.GraphicData!.GetFirstChild<A.Table>()!;

        /// <summary>
        ///     Returns number of rows in the table.
        /// </summary>
        public int Rows => Frame.Graphic!.GraphicData!.GetFirstChild<A.Table>()!.Elements<A.TableRow>().Count();

        /// <summary>
        ///     Returns number of columns in the table.
        /// </summary>
        public int Columns => Frame.Graphic!.GraphicData!.GetFirstChild<A.Table>()!.TableGrid!.Elements<A.GridColumn>()
            .Count();

        /// <summary>
        ///     Row wrappers for the table.
        /// </summary>
        public IReadOnlyList<PowerPointTableRow> RowItems =>
            TableElement.Elements<A.TableRow>().Select(r => new PowerPointTableRow(this, r)).ToList();

        /// <summary>
        ///     Column wrappers for the table.
        /// </summary>
        public IReadOnlyList<PowerPointTableColumn> ColumnItems =>
            TableElement.TableGrid?.Elements<A.GridColumn>()
                .Select(c => new PowerPointTableColumn(this, c))
                .ToList() ?? new List<PowerPointTableColumn>();

        /// <summary>
        ///     Enables or disables header row styling (firstRow attribute) on the table.
        /// </summary>
        public bool HeaderRow {
            get => FirstRow;
            set => FirstRow = value;
        }

        /// <summary>
        ///     Enables or disables first row styling on the table.
        /// </summary>
        public bool FirstRow {
            get => TableElement.TableProperties?.FirstRow?.Value == true;
            set {
                TableElement.TableProperties ??= new A.TableProperties();
                TableElement.TableProperties.FirstRow = value;
            }
        }

        /// <summary>
        ///     Enables or disables last row styling on the table.
        /// </summary>
        public bool LastRow {
            get => TableElement.TableProperties?.LastRow?.Value == true;
            set {
                TableElement.TableProperties ??= new A.TableProperties();
                TableElement.TableProperties.LastRow = value;
            }
        }

        /// <summary>
        ///     Enables or disables first column styling on the table.
        /// </summary>
        public bool FirstColumn {
            get => TableElement.TableProperties?.FirstColumn?.Value == true;
            set {
                TableElement.TableProperties ??= new A.TableProperties();
                TableElement.TableProperties.FirstColumn = value;
            }
        }

        /// <summary>
        ///     Enables or disables last column styling on the table.
        /// </summary>
        public bool LastColumn {
            get => TableElement.TableProperties?.LastColumn?.Value == true;
            set {
                TableElement.TableProperties ??= new A.TableProperties();
                TableElement.TableProperties.LastColumn = value;
            }
        }

        /// <summary>
        ///     Enables or disables banded rows styling (bandRow attribute) on the table.
        /// </summary>
        public bool BandedRows {
            get => TableElement.TableProperties?.BandRow?.Value == true;
            set {
                TableElement.TableProperties ??= new A.TableProperties();
                TableElement.TableProperties.BandRow = value;
            }
        }

        /// <summary>
        ///     Enables or disables banded columns styling (bandCol attribute) on the table.
        /// </summary>
        public bool BandedColumns {
            get => TableElement.TableProperties?.BandColumn?.Value == true;
            set {
                TableElement.TableProperties ??= new A.TableProperties();
                TableElement.TableProperties.BandColumn = value;
            }
        }

        /// <summary>
        ///     Gets or sets the table style ID GUID.
        /// </summary>
        public string? StyleId {
            get => TableElement.TableProperties?.GetFirstChild<A.TableStyleId>()?.Text;
            set {
                TableElement.TableProperties ??= new A.TableProperties();
                TableElement.TableProperties.RemoveAllChildren<A.TableStyleId>();
                if (!string.IsNullOrWhiteSpace(value)) {
                    TableElement.TableProperties.Append(new A.TableStyleId { Text = value! });
                }
            }
        }

    }
}
