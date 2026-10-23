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
        /// <summary>Shares immutable wrapped content while isolating the physical row heights used by a fragment preview.</summary>
        public PreparedFlowTableRows WithIndependentHeights() {
            var copy = new PreparedFlowTableRows(Lines.Length);
            Array.Copy(Lines, copy.Lines, Lines.Length);
            Array.Copy(LineCounts, copy.LineCounts, LineCounts.Length);
            Array.Copy(Heights, copy.Heights, Heights.Length);
            Array.Copy(IntrinsicHeights, copy.IntrinsicHeights, IntrinsicHeights.Length);
            Array.Copy(Leadings, copy.Leadings, Leadings.Length);
            Array.Copy(Sizes, copy.Sizes, Sizes.Length);
            Array.Copy(RunFontSizeScales, copy.RunFontSizeScales, RunFontSizeScales.Length);
            Array.Copy(Bold, copy.Bold, Bold.Length);
            return copy;
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
