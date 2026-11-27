using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.OpenXml.Internal;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// The SDK side of a table: its constructor over an SDK graphic frame, and the SDK table, cells and slide part that
    /// the image renderer, the layout engine and the legacy PPT writer still read.
    /// </summary>
    public partial class PowerPointTable {
        internal PowerPointTable(GraphicFrame frame, OpenXmlPartContainer? ownerPart = null) : base(frame, ownerPart) {
        }

        /// <summary>The SDK table; reaching it for a table held in a package tree opens the SDK.</summary>
        internal A.Table TableElement => ((GraphicFrame)Element).Graphic!.GraphicData!.GetFirstChild<A.Table>()!;

        /// <summary>The SDK slide part the table is on.</summary>
        internal SlidePart? SlidePart => SlidePartRef.Current.SdkPart as SlidePart ?? OwnerSlide?.SlidePart;

        internal PowerPointTableCell WrapCell(A.TableCell cell) => CellFor(OfficeXmlNode.FromSdk(cell));
    }

    public partial class PowerPointTableCell {
        /// <summary>The SDK cell; reaching it for a cell held in a package tree opens the SDK.</summary>
        internal A.TableCell Cell => (A.TableCell)_cell.RequireSdkElement();

        /// <summary>The SDK slide part the cell is on.</summary>
        internal SlidePart? SlidePart => _table.SlidePart;

        /// <summary>The SDK part that holds the cell.</summary>
        internal OpenXmlPartContainer? OwnerPart => _table.RelationshipOwnerPart ?? _table.SlidePart;
    }
}
