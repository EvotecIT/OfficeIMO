using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.PowerPoint.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Represents a table on a slide.
    /// </summary>
    public partial class PowerPointTable : PowerPointShape {
        private const int D = OfficeXmlNamespaces.DrawingML;
        private const int EmusPerPoint = 12700;

        internal PowerPointTable(PowerPointSlide slide, OfficeXmlTree tree, int node) : base(slide, tree, node) {
        }

        internal PowerPointTable(OpenXml.PowerPointPartRef part, OfficeXmlTree tree, int node) : base(part, tree, node) {
        }

        /// <summary>The table's <c>a:tbl</c> in its graphic frame, held where the frame is.</summary>
        internal OfficeXmlNode TableNode => XmlElement.Child(D, "graphic").Child(D, "graphicData").Child(D, "tbl");

        private OfficeXmlNode GridNode => TableNode.Child(D, "tblGrid");

        private OfficeXmlNode PropertiesNode => TableNode.Child(D, "tblPr");

        private void MarkEdited() => OwnerPartRef.MarkEdited();

        /// <summary>A cell wrapper for a cell of this table.</summary>
        internal PowerPointTableCell CellFor(OfficeXmlNode cell) => new(this, cell);

        /// <summary>
        ///     Returns number of rows in the table.
        /// </summary>
        public int RowCount => CountChildren(TableNode, "tr");

        /// <summary>
        ///     Returns number of columns in the table.
        /// </summary>
        public int ColumnCount => CountChildren(GridNode, "gridCol");

        private static int CountChildren(OfficeXmlNode parent, string localName) {
            int count = 0;
            foreach (OfficeXmlNode _ in parent.Children(D, localName)) count++;
            return count;
        }

        // The child at an index among those with a name, or none.
        private static OfficeXmlNode NthChild(OfficeXmlNode parent, string localName, int index) {
            if (index < 0) return default;
            foreach (OfficeXmlNode child in parent.Children(D, localName)) {
                if (index-- == 0) return child;
            }
            return default;
        }

        // The table's rows, and each row's cells, as a grid.
        private List<List<OfficeXmlNode>> CellNodes() {
            var rows = new List<List<OfficeXmlNode>>();
            foreach (OfficeXmlNode row in TableNode.Children(D, "tr")) rows.Add(row.Children(D, "tc").ToList());
            return rows;
        }

        /// <summary>
        ///     Row wrappers for the table.
        /// </summary>
        public IReadOnlyList<PowerPointTableRow> Rows {
            get {
                var rows = new List<PowerPointTableRow>();
                foreach (OfficeXmlNode row in TableNode.Children(D, "tr")) rows.Add(new PowerPointTableRow(this, row));
                return rows;
            }
        }

        /// <summary>
        ///     Column wrappers for the table.
        /// </summary>
        public IReadOnlyList<PowerPointTableColumn> Columns {
            get {
                var columns = new List<PowerPointTableColumn>();
                foreach (OfficeXmlNode column in GridNode.Children(D, "gridCol")) columns.Add(new PowerPointTableColumn(this, column));
                return columns;
            }
        }

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
            get => ReadFlag("firstRow");
            set => SetFlag("firstRow", value);
        }

        /// <summary>
        ///     Enables or disables last row styling on the table.
        /// </summary>
        public bool LastRow {
            get => ReadFlag("lastRow");
            set => SetFlag("lastRow", value);
        }

        /// <summary>
        ///     Enables or disables first column styling on the table.
        /// </summary>
        public bool FirstColumn {
            get => ReadFlag("firstCol");
            set => SetFlag("firstCol", value);
        }

        /// <summary>
        ///     Enables or disables last column styling on the table.
        /// </summary>
        public bool LastColumn {
            get => ReadFlag("lastCol");
            set => SetFlag("lastCol", value);
        }

        /// <summary>
        ///     Enables or disables banded rows styling (bandRow attribute) on the table.
        /// </summary>
        public bool BandedRows {
            get => ReadFlag("bandRow");
            set => SetFlag("bandRow", value);
        }

        /// <summary>
        ///     Enables or disables banded columns styling (bandCol attribute) on the table.
        /// </summary>
        public bool BandedColumns {
            get => ReadFlag("bandCol");
            set => SetFlag("bandCol", value);
        }

        private bool ReadFlag(string attribute) => PropertiesNode.BooleanAttribute(attribute) == true;

        private void SetFlag(string attribute, bool value) {
            EnsureProperties().SetAttribute(attribute, value ? "1" : "0");
            MarkEdited();
        }

        private OfficeXmlNode EnsureProperties() => TableNode.GetOrAddChild(D, "tblPr", ATableView.Order);

        /// <summary>
        ///     Gets or sets the table style ID GUID.
        /// </summary>
        public string? StyleId {
            get => PropertiesNode.Child(D, "tableStyleId") is { IsNone: false } style ? style.InnerText : null;
            set {
                OfficeXmlNode properties = EnsureProperties();
                properties.RemoveChildren(D, "tableStyleId");
                if (!string.IsNullOrWhiteSpace(value)) {
                    properties.Append(properties.CreateElement(D, "tableStyleId")).SetInnerText(value!);
                }
                MarkEdited();
            }
        }

    }
}
