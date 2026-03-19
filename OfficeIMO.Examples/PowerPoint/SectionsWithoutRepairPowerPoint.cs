using System;
using System.IO;
using System.Linq;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.Examples.PowerPoint {
    /// <summary>
    /// Demonstrates section creation without triggering a PowerPoint repair prompt.
    /// </summary>
    public static class SectionsWithoutRepairPowerPoint {
        public static void Example_PowerPointSectionsWithoutRepair(string folderPath, bool openPowerPoint) {
            Console.WriteLine("[*] PowerPoint - Sections without repair");
            string filePath = Path.Combine(folderPath, "Sections Without Repair.pptx");

            using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
            const double marginCm = 1.5;
            PowerPointLayoutBox content = presentation.SlideSize.GetContentBoxCm(marginCm);

            PowerPointSlide overview = presentation.AddSlide();
            AddTitleAndBody(
                overview,
                content,
                "Sectioned Presentation",
                "This example creates real PowerPoint sections and opens cleanly without a repair prompt.");
            overview.Notes.Text = "Use this deck to verify that sections now persist correctly.";

            PowerPointSlide agenda = presentation.AddSlide();
            AddTitleAndBody(
                agenda,
                content,
                "Introduction",
                "The first section contains the overview and agenda slides.");
            PowerPointTextBox bullets = agenda.AddTextBoxCm(string.Empty, content.LeftCm, content.TopCm + 3.0, content.WidthCm, 3.5);
            bullets.AddBullets(new[] {
                "Introduction section",
                "Results section",
                "Closing section",
                "Morph transition slide"
            });
            bullets.ApplyAutoSpacing(lineSpacingMultiplier: 1.15, spaceAfterPoints: 2);

            PowerPointSlide results = presentation.AddSlide();
            AddTitleAndBody(
                results,
                content,
                "Results",
                "This slide belongs to a different section and uses a morph transition.");
            results.Transition = SlideTransition.Morph;
            results.AddTextBoxCm("If PowerPoint opens this deck directly, both sections and transitions are serialized correctly.",
                content.LeftCm, content.TopCm + 3.0, content.WidthCm, 1.4);

            PowerPointSlide closing = presentation.AddSlide();
            AddTitleAndBody(
                closing,
                content,
                "Closing",
                "The final section confirms the issue reported in #1714 is covered by a concrete example.");

            presentation.AddSection("Introduction", startSlideIndex: 0);
            presentation.AddSection("Results", startSlideIndex: 2);
            presentation.AddSection("Closing", startSlideIndex: 3);

            presentation.Save();

            var errors = presentation.ValidateDocument();
            if (errors.Count > 0) {
                string message = string.Join(Environment.NewLine, errors.Select(error =>
                    $"{error.Description} | Part={error.Part?.Uri} | Path={error.Path?.XPath}"));
                throw new InvalidOperationException("Generated example failed Open XML validation:" + Environment.NewLine + message);
            }

            Console.WriteLine($"[+] Created: {filePath}");
            Console.WriteLine("[+] Validation passed with no Open XML errors.");
            Helpers.Open(filePath, openPowerPoint);
        }

        private static void AddTitleAndBody(PowerPointSlide slide, PowerPointLayoutBox content, string title, string body) {
            PowerPointTextBox titleBox = slide.AddTitleCm(title, content.LeftCm, content.TopCm, content.WidthCm, 1.4);
            if (titleBox.Paragraphs.Count > 0) {
                PowerPointTextStyle.Title.WithColor("1F4E79").Apply(titleBox.Paragraphs[0]);
            }

            PowerPointTextBox bodyBox = slide.AddTextBoxCm(body, content.LeftCm, content.TopCm + 1.9, content.WidthCm, 1.2);
            bodyBox.ApplyTextStyle(PowerPointTextStyle.Body);
            bodyBox.SetTextMarginsCm(0.2, 0.1, 0.2, 0.1);
        }
    }
}
