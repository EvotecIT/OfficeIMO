using System.IO;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.Examples.Showcase.Workflows;

/// <summary>Creates an editable roadmap deck with phase cards and a decision table.</summary>
internal static class ProjectRoadmap {
    internal static void Create(string folder) {
        using PowerPointPresentation deck = PowerPointPresentation.Create(Path.Combine(folder, "example.pptx"));
        PowerPointSlide roadmap = deck.AddSlide();
        Heading(roadmap, "From pilot to everyday service", "PROJECT ROADMAP / NEXT 90 DAYS");
        string[] phases = { "01 / Prepare", "02 / Prove", "03 / Expand" };
        string[] details = { "Agree the scope\nName the owners\nPrepare support", "Run the pilot\nExercise recovery\nCollect feedback", "Invite more teams\nReview the measures\nHand over ownership" };
        string[] colors = { "EAF1FB", "EEE8F7", "E7F6ED" };
        for (int index = 0; index < phases.Length; index++) {
            double x = 1.5 + index * 10.3;
            roadmap.AddRectangleCm(x, 5.5, 9.3, 9, phases[index]).Fill(colors[index]).Stroke("CBD5E1", 1);
            Text(roadmap, phases[index], x + 0.5, 6.1, 8.3, 1.4, 25);
            Text(roadmap, details[index], x + 0.5, 8.3, 8.3, 5, 21);
        }

        PowerPointSlide gates = deck.AddSlide();
        Heading(gates, "Advance on evidence", "DECISION CHECKPOINTS");
        PowerPointTable table = gates.AddTableCm(4, 3, 1.5, 5.3, 30, 10);
        string[,] values = {
            { "Checkpoint", "Evidence", "Decision owner" },
            { "Pilot ready", "Support and recovery walkthrough", "Operations" },
            { "Pilot complete", "Observed request outcomes", "Product" },
            { "Service accepted", "Named owner and runbook", "Service owner" }
        };
        for (int row = 0; row < 4; row++) {
            for (int column = 0; column < 3; column++) {
                var cell = table.GetCell(row, column);
                cell.Text = values[row, column];
                cell.FontSize = 19;
                cell.Bold = row == 0;
                cell.FillColor = row == 0 ? "17365D" : "F1F5F9";
                cell.Color = row == 0 ? "FFFFFF" : "17365D";
                cell.PaddingLeftPoints = 12;
                cell.VerticalAlignment = PowerPointTextVerticalAlignment.Center;
            }
        }
        deck.Save();

        void Heading(PowerPointSlide slide, string title, string subtitle) {
            Text(slide, title, 1.5, 1.3, 30, 1.8, 32);
            Text(slide, subtitle, 1.5, 3.4, 30, 1, 15);
        }
        void Text(PowerPointSlide slide, string value, double x, double y, double width, double height, int size) {
            var box = slide.AddTextBoxCm(value, x, y, width, height);
            box.FontSize = size;
            box.Color = "17365D";
        }
    }
}
