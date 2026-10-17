namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    /// <summary>
    /// Distributes a spanning percentage's remaining request over tracks without
    /// an earlier percentage, weighted by their non-spanning preferred content.
    /// </summary>
    private void ApplyTableSpanningPercentageWidths(double[] percentages, IReadOnlyList<double> preferred,
        IReadOnlyList<(int Column, int Span, double Percentage)> cells) {
        foreach (var group in cells.GroupBy(cell => cell.Span).OrderBy(group => group.Key)) {
            CheckCancellation();
            ChargeLayoutOperations(percentages.Length, "table spanning percentage column contributions");
            double[] prior = percentages.ToArray();
            foreach (var cell in group) {
                CheckCancellation();
                ChargeLayoutOperations(cell.Span, "table spanning percentage column contributions");
                double existing = 0D;
                double weight = 0D;
                int eligible = 0;
                for (int offset = 0; offset < cell.Span; offset++) {
                    int column = cell.Column + offset;
                    existing += prior[column];
                    if (prior[column] > 0D) continue;
                    weight += preferred[column];
                    eligible++;
                }
                double remaining = Math.Max(0D, cell.Percentage - existing);
                if (eligible == 0 || remaining == 0D) continue;
                for (int offset = 0; offset < cell.Span; offset++) {
                    int column = cell.Column + offset;
                    if (prior[column] > 0D) continue;
                    double share = weight > 0D ? preferred[column] / weight : 1D / eligible;
                    percentages[column] = Math.Max(percentages[column], remaining * share);
                }
            }
        }

        // Earlier source columns retain their requests when the total exceeds 100%.
        double available = 1D;
        for (int column = 0; column < percentages.Length; column++) {
            percentages[column] = Math.Min(available, percentages[column]);
            available = Math.Max(0D, available - percentages[column]);
        }
    }
}
