using System.Linq;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// Represents a table on a slide.
    /// </summary>
    public class PPTable : PPShape {
        internal PPTable(GraphicFrame frame) : base(frame) {
        }

        private GraphicFrame Frame => (GraphicFrame)Element;

        /// <summary>
        /// Returns number of rows in the table.
        /// </summary>
        public int Rows => Frame.Graphic!.GraphicData!.GetFirstChild<A.Table>()!.Elements<A.TableRow>().Count();

        /// <summary>
        /// Returns number of columns in the table.
        /// </summary>
        public int Columns => Frame.Graphic!.GraphicData!.GetFirstChild<A.Table>()!.TableGrid!.Elements<A.GridColumn>().Count();
    }
}

