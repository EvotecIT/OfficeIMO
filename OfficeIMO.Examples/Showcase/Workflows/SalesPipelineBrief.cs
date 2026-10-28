using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.Examples.Showcase.Workflows {
    /// <summary>Shows pipeline value by stage separately from committed revenue and assigns review actions.</summary>
    internal static class SalesPipelineBrief {
        internal static void Create(string folder) {
            using var deck = PowerPointDocument.Create(Path.Combine(folder, "example.pptx"));
            var pipeline = deck.AddSlide();
            Text(pipeline, "Sales pipeline: where to focus next", 1.5, 1, 30, 2, 30, "17365D");
            Text(pipeline, "Synthetic snapshot / opportunity value in EUR thousands / not a revenue forecast", 1.5, 3, 30, 1.5, 17, "526179");
            var data = new OfficeChartData(new[] { "Qualified", "Discovery", "Proposal", "Negotiation" },
                new[] { new OfficeSeries("Opportunity value", new double[] { 180, 145, 96, 52 }) });
            pipeline.AddChart(OfficeChartType.ColumnClustered,data,global::OfficeIMO.OfficeLength.Centimeters(1.5),global::OfficeIMO.OfficeLength.Centimeters(5),global::OfficeIMO.OfficeLength.Centimeters(30),global::OfficeIMO.OfficeLength.Centimeters(11))
                .SetTitle("Open value by current stage").SetDataLabels(showValue: true).SetValueAxisTitle("EUR thousands");
            pipeline.Notes.Text = "Each opportunity appears in one current stage. Do not add stage percentages or describe this as contracted revenue. Values are synthetic.";
            var actions = deck.AddSlide();
            Text(actions, "Move the right conversations forward", 1.5, 1.5, 30, 2, 30, "17365D");
            string[] titles = { "Qualify", "Unblock", "Confirm" };
            string[] bodies = { "Validate need and timing.\nOwner: account team", "Resolve proposal questions.\nOwner: solution lead", "Record the next decision date.\nOwner: opportunity owner" };
            for (int i = 0; i < 3; i++) {
                double y = 5 + i * 3.6;
                Text(actions, titles[i], 1.5, y, 7, 1.5, 24, "2563EB");
                Text(actions, bodies[i], 10, y, 21, 2.7, 22, "243247");
            }
            actions.Notes.Text = "Use the real CRM snapshot date and keep confidential customer details out of public examples.";
            deck.Save();
        }
        private static void Text(PowerPointSlide slide, string value, double x, double y, double width, double height, int size, string color) {
            var box = slide.AddTextBox(value,global::OfficeIMO.OfficeLength.Centimeters(x),global::OfficeIMO.OfficeLength.Centimeters(y),global::OfficeIMO.OfficeLength.Centimeters(width),global::OfficeIMO.OfficeLength.Centimeters(height)); box.FontSize = size; box.Color = color;
        }
    }
}
