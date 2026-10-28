using System.IO;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.Examples.Showcase.Workflows;

/// <summary>Creates a short question-and-answer deck from structured training content.</summary>
internal static class TrainingQuiz {
    internal static void Create(string folder) {
        using PowerPointDocument deck = PowerPointDocument.Create(Path.Combine(folder, "example.pptx"));
        AddSlide("What should happen first?",
            "A user reports that a shared service is unavailable.",
            new[] { "A / Restart every component", "B / Confirm impact and name the incident owner", "C / Wait for more reports" },
            "DISCUSS / 60 SECONDS", "EAF1FB");
        AddSlide("Confirm impact. Establish ownership.",
            "The incident owner coordinates the response and keeps communication consistent.",
            new[] { "Identify who and what is affected.", "Record the time and the observed symptoms.", "Choose the next check before making changes." },
            "ANSWER / B", "E7F6ED");
        AddSlide("Put the lesson into practice",
            "Use a recent example from your own service.",
            new[] { "Who would take ownership?", "Where would you record impact?", "How would the team communicate updates?" },
            "TEAM EXERCISE / 5 MINUTES", "F7EEDC");
        deck.Save();

        void AddSlide(string title, string description, string[] points, string label, string fill) {
            PowerPointSlide slide = deck.AddSlide();
            slide.AddRectangle(OfficeIMO.OfficeLength.Centimeters(0), OfficeIMO.OfficeLength.Centimeters(0), OfficeIMO.OfficeLength.Centimeters(33.866), OfficeIMO.OfficeLength.Centimeters(19.05), "Background").Fill(fill).Stroke(fill, 0);
            var eyebrow = slide.AddTextBox(label,global::OfficeIMO.OfficeLength.Centimeters(1.6),global::OfficeIMO.OfficeLength.Centimeters(1.2),global::OfficeIMO.OfficeLength.Centimeters(29),global::OfficeIMO.OfficeLength.Centimeters(1));
            eyebrow.FontSize = global::OfficeIMO.OfficeLength.Points(15);
            eyebrow.Color = "526179";
            var heading = slide.AddTextBox(title,global::OfficeIMO.OfficeLength.Centimeters(1.6),global::OfficeIMO.OfficeLength.Centimeters(3),global::OfficeIMO.OfficeLength.Centimeters(29),global::OfficeIMO.OfficeLength.Centimeters(2.3));
            heading.FontSize = global::OfficeIMO.OfficeLength.Points(32);
            heading.Color = "17365D";
            var introduction = slide.AddTextBox(description,global::OfficeIMO.OfficeLength.Centimeters(1.6),global::OfficeIMO.OfficeLength.Centimeters(5.8),global::OfficeIMO.OfficeLength.Centimeters(29),global::OfficeIMO.OfficeLength.Centimeters(2));
            introduction.FontSize = global::OfficeIMO.OfficeLength.Points(21);
            introduction.Color = "526179";
            for (int index = 0; index < points.Length; index++) {
                var point = slide.AddTextBox(points[index],global::OfficeIMO.OfficeLength.Centimeters(2.2),global::OfficeIMO.OfficeLength.Centimeters(8.6 + index * 2.2),global::OfficeIMO.OfficeLength.Centimeters(28),global::OfficeIMO.OfficeLength.Centimeters(1.7));
                point.FontSize = global::OfficeIMO.OfficeLength.Points(23);
                point.Color = "17365D";
            }
        }
    }
}
