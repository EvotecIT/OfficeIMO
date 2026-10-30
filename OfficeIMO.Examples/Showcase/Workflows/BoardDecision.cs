using OfficeIMO.PowerPoint;

namespace OfficeIMO.Examples.Showcase.Workflows {
    /// <summary>Builds a two-slide decision brief with explicit options, trade-offs and an approval record.</summary>
    internal static class BoardDecision {
        internal static void Create(string folder) {
            using var deck = PowerPointDocument.Create(Path.Combine(folder, "example.pptx"));
            var options = deck.AddSlide();
            Text(options, "Choose the next service investment", 1.5, 1, 30, 2, 30, "17365D");
            Text(options, "NORTHWIND / GOVERNANCE REVIEW / SYNTHETIC ESTIMATES", 1.5, 3, 30, 1, 14, "526179");
            Card(options, 1.5, "1 / Keep current", "No new spend\nManual triage continues\nReview in three months", "F1F5F9");
            Card(options, 12, "2 / Improve intake", "Six-week pilot\nEUR 4,800 estimate\nReversible first step", "E7F6ED");
            Card(options, 22.5, "3 / Replace platform", "Larger investment\nProcurement required\nLonger lead time", "EAF1FB");
            options.Notes.Text = "Explain the assumptions behind each estimate. The green panel identifies the proposed option, not an approved commitment.";
            var decision = deck.AddSlide();
            Text(decision, "Decision requested: approve a bounded pilot", 1.5, 1, 30, 2.5, 30, "17365D");
            Text(decision, "Outcome", 1.5, 5, 13, 1, 23, "17365D");
            Text(decision, "Measure time spent assigning requests.\nUse the same baseline and pilot definition.", 1.5, 7, 14, 5, 22, "243247");
            Text(decision, "Approval record", 18, 5, 13, 1, 23, "17365D");
            Text(decision, "Sponsor: to be named\nBudget: illustrative EUR 4,800\nReview: after six weeks\nStop if ownership is unresolved", 18, 7, 13, 7, 22, "243247");
            decision.Notes.Text = "Record approval, conditions and the named sponsor in the actual governance record. This deck is a fictional example.";
            deck.Save();
        }
        private static void Card(PowerPointSlide slide, double x, string title, string body, string fill) {
            slide.AddRectangle(OfficeIMO.OfficeLength.Centimeters(x), OfficeIMO.OfficeLength.Centimeters(5), OfficeIMO.OfficeLength.Centimeters(9.5), OfficeIMO.OfficeLength.Centimeters(10)).Fill(fill).Stroke("CBD5E1", global::OfficeIMO.OfficeLength.Points(1));
            Text(slide, title, x + 0.6, 5.8, 8.3, 2, 22, "17365D");
            Text(slide, body, x + 0.6, 8.2, 8.3, 5.5, 20, "243247");
        }
        private static void Text(PowerPointSlide slide, string value, double x, double y, double width, double height, int size, string color) {
            var box = slide.AddTextBox(value,global::OfficeIMO.OfficeLength.Centimeters(x),global::OfficeIMO.OfficeLength.Centimeters(y),global::OfficeIMO.OfficeLength.Centimeters(width),global::OfficeIMO.OfficeLength.Centimeters(height)); box.FontSize = global::OfficeIMO.OfficeLength.Points(size); box.Color = color;
        }
    }
}
