namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed class PreparedFlowTableRows {
        public PreparedFlowTableRows(int rows) {
            Lines = new TableCellTextLayout[rows][];
            LineCounts = new int[rows]; Heights = new double[rows]; Leadings = new double[rows];
            IntrinsicHeights = new double[rows];
            Sizes = new double[rows]; Bold = new bool[rows];
            RunFontSizeScales = new double[rows];
        }
        public TableCellTextLayout[][] Lines { get; }
        public int[] LineCounts { get; }
        public double[] Heights { get; }
        public double[] IntrinsicHeights { get; }
        public double[] Leadings { get; }
        public double[] Sizes { get; }
        public double[] RunFontSizeScales { get; }
        public bool[] Bold { get; }
    }
}
