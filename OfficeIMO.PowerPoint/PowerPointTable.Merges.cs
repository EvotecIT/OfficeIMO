using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointTable {
        // Every grid position holds one a:tc: an anchor carries rowSpan/gridSpan and the cells it covers carry
        // vMerge/hMerge. PowerPoint repairs tables whose spans and covered cells disagree, so structural edits
        // keep both sides in step.

        private static bool IsCoveredHorizontally(OfficeXmlNode cell) => cell.BooleanAttribute("hMerge") == true;
        private static bool IsCoveredVertically(OfficeXmlNode cell) => cell.BooleanAttribute("vMerge") == true;

        private static (int Row, int Column) FindMergeAnchor(List<List<OfficeXmlNode>> cells, int row, int column) {
            while (column > 0 && IsCoveredHorizontally(cells[row][column])) column--;
            while (row > 0 && IsCoveredVertically(cells[row][column])) row--;
            return (row, column);
        }

        private void DissolveMergesTouching(int topRow, int leftColumn, int bottomRow, int rightColumn) {
            List<List<OfficeXmlNode>> cells = CellNodes();
            var anchors = new HashSet<(int Row, int Column)>();
            for (int r = topRow; r <= bottomRow; r++) {
                for (int c = leftColumn; c <= rightColumn; c++) anchors.Add(FindMergeAnchor(cells, r, c));
            }

            foreach ((int anchorRow, int anchorColumn) in anchors) {
                OfficeXmlNode anchor = cells[anchorRow][anchorColumn];
                int lastRow = Math.Min(cells.Count, anchorRow + (anchor.Int32Attribute("rowSpan") ?? 1));
                int lastColumn = Math.Min(cells[anchorRow].Count, anchorColumn + (anchor.Int32Attribute("gridSpan") ?? 1));
                for (int r = anchorRow; r < lastRow; r++) {
                    for (int c = anchorColumn; c < lastColumn; c++) SetMerge(cells[r][c], null, null, false, false);
                }
            }
        }

    }
}
