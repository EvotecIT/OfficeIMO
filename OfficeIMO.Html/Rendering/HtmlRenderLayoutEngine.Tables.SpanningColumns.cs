namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    /// <summary>Distributes spanning content through the same automatic, length and percentage tiers as ordinary columns.</summary>
    private void ApplyTableSpanningContentWidths(double[] minimums, double[] preferred,
        IReadOnlyList<(int Column, int Span, double Minimum, double Preferred)> cells, double spacing,
        IReadOnlyList<bool> authored, IReadOnlyList<double> percentages) {
        foreach (var group in cells.GroupBy(cell => cell.Span).OrderBy(group => group.Key)) {
            CheckCancellation();
            double[] priorMinimums = minimums.ToArray();
            double[] priorPreferred = preferred.ToArray();
            foreach (var cell in group) {
                CheckCancellation();
                ChargeLayoutOperations(cell.Span, "table spanning column contributions");
                double innerSpacing = spacing * (cell.Span - 1);
                int[] columns = Enumerable.Range(cell.Column, cell.Span).ToArray();
                double[] lower = columns.Select(index => priorMinimums[index]).ToArray();
                double[] upper = columns.Select(index => priorPreferred[index]).ToArray();
                bool[] constrained = columns.Select(index => authored[index]).ToArray();
                double[] shares = columns.Select(index => percentages[index]).ToArray();
                double[] minimum = AllocateAutoColumnSizingGuesses(lower, upper, constrained, shares,
                    Math.Max(0D, cell.Minimum - innerSpacing));
                double[] maximum = AllocateAutoColumnSizingGuesses(minimum, upper, constrained, shares,
                    Math.Max(0D, cell.Preferred - innerSpacing));
                for (int index = 0; index < columns.Length; index++) {
                    int column = columns[index];
                    minimums[column] = Math.Max(minimums[column], minimum[index]);
                    preferred[column] = Math.Max(preferred[column], maximum[index]);
                }
            }
        }
    }

    /// <summary>Fits available width between content, percentage, specified and preferred width guesses.</summary>
    private static double[] AllocateAutoColumnSizingGuesses(double[] minimums, double[] preferred,
        bool[] authored, double[] percentages, double requestedWidth) {
        double width = Math.Max(requestedWidth, minimums.Sum());
        double[] prior = minimums;
        for (int stage = 0; stage < 3; stage++) {
            var next = new double[minimums.Length];
            for (int index = 0; index < next.Length; index++) {
                next[index] = percentages[index] > 0D ? Math.Max(minimums[index], percentages[index] * width)
                    : stage == 2 || stage == 1 && authored[index] ? Math.Max(minimums[index], preferred[index])
                    : minimums[index];
            }
            double priorTotal = prior.Sum();
            double nextTotal = next.Sum();
            if (nextTotal >= width) {
                double progress = nextTotal > priorTotal ? (width - priorTotal) / (nextTotal - priorTotal) : 0D;
                for (int index = 0; index < next.Length; index++) next[index] = prior[index] + (next[index] - prior[index]) * progress;
                return next;
            }
            prior = next;
        }
        return AllocateAutoColumnWidths(minimums, prior, authored, percentages, minimums, width).ToArray();
    }
}
