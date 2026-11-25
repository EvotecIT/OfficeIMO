using OfficeIMO.Core.Internal.OpenXml;
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
        private readonly SlidePart? _initialSlidePart;
        private SlidePart? _slidePart => Element.Ancestors<DocumentFormat.OpenXml.Presentation.Slide>().FirstOrDefault()?.SlidePart ?? _initialSlidePart;

        internal PowerPointTable(GraphicFrame frame, OpenXmlPartContainer? ownerPart = null) : base(frame, ownerPart) {
            _initialSlidePart = ownerPart as SlidePart;
        }

        internal PowerPointTable(PowerPointSlide slide, OfficeXmlTree tree, int node) : base(slide, tree, node) {
        }

        private GraphicFrame Frame => (GraphicFrame)Element;
        internal SlidePart? SlidePart => _slidePart ?? OwnerSlide?.SlidePart;
        internal PowerPointTableCell WrapCell(A.TableCell cell) => new(cell, SlidePart, RelationshipOwnerPart ?? SlidePart);
        internal A.Table TableElement => Frame.Graphic!.GraphicData!.GetFirstChild<A.Table>()!;

        /// <summary>
        ///     Returns number of rows in the table.
        /// </summary>
        public int RowCount => TryGetPackageTable(out OfficeXmlTree tree, out int table)
            ? CountChildren(tree, table, "tr")
            : Frame.Graphic!.GraphicData!.GetFirstChild<A.Table>()!.Elements<A.TableRow>().Count();

        /// <summary>
        ///     Returns number of columns in the table.
        /// </summary>
        public int ColumnCount => TryGetPackageTable(out OfficeXmlTree tree, out int table) && tree.FindChild(table, OfficeXmlNamespaces.DrawingML, "tblGrid") is var grid
            && grid != OfficeXmlTree.NoNode
            ? CountChildren(tree, grid, "gridCol")
            : Frame.Graphic!.GraphicData!.GetFirstChild<A.Table>()!.TableGrid!.Elements<A.GridColumn>().Count();

        // The a:tbl of the frame on the package tree, as the SDK finds it; false once the table is bound or has none.
        internal bool TryGetPackageTable(out OfficeXmlTree tree, out int table) {
            table = OfficeXmlTree.NoNode;
            if (!TryGetPackageElement(out tree, out int frame)) return false;
            int graphic = tree.FindChild(frame, OfficeXmlNamespaces.DrawingML, "graphic");
            int data = graphic == OfficeXmlTree.NoNode ? OfficeXmlTree.NoNode : tree.FindChild(graphic, OfficeXmlNamespaces.DrawingML, "graphicData");
            table = data == OfficeXmlTree.NoNode ? OfficeXmlTree.NoNode : tree.FindChild(data, OfficeXmlNamespaces.DrawingML, "tbl");
            return table != OfficeXmlTree.NoNode;
        }

        private static int CountChildren(OfficeXmlTree tree, int parent, string localName) {
            int count = 0;
            for (int child = tree.FindChild(parent, OfficeXmlNamespaces.DrawingML, localName); child != OfficeXmlTree.NoNode;
                 child = tree.FindNextSibling(child, OfficeXmlNamespaces.DrawingML, localName)) {
                count++;
            }
            return count;
        }

        private static int NthChild(OfficeXmlTree tree, int parent, string localName, int index) {
            if (index < 0) return OfficeXmlTree.NoNode;
            int child = tree.FindChild(parent, OfficeXmlNamespaces.DrawingML, localName);
            for (; child != OfficeXmlTree.NoNode && index > 0; index--) child = tree.FindNextSibling(child, OfficeXmlNamespaces.DrawingML, localName);
            return child;
        }

        /// <summary>
        ///     Row wrappers for the table.
        /// </summary>
        public IReadOnlyList<PowerPointTableRow> Rows =>
            TableElement.Elements<A.TableRow>().Select(r => new PowerPointTableRow(this, r)).ToList();

        /// <summary>
        ///     Column wrappers for the table.
        /// </summary>
        public IReadOnlyList<PowerPointTableColumn> Columns =>
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
