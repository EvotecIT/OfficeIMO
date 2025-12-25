using System;
using System.IO;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.Examples.PowerPoint {
    /// <summary>
    /// Demonstrates slide removal and reordering.
    /// </summary>
    public static class SlidesManagementPowerPoint {
        public static void Example_SlidesManagement(string folderPath, bool openPowerPoint) {
            Console.WriteLine("[*] PowerPoint - Slides management");
            string filePath = Path.Combine(folderPath, "Slides Management.pptx");
            using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
            const double marginCm = 1.5;
            PowerPointLayoutBox content = presentation.SlideSize.GetContentBoxCm(marginCm);

            PowerPointSlide slide1 = presentation.AddSlide();
            ConfigureSlide(slide1, content, "Original Slide 1", "This slide is created first.");

            PowerPointSlide slide2 = presentation.AddSlide();
            ConfigureSlide(slide2, content, "Original Slide 2", "This slide will be hidden in slide show.");

            PowerPointSlide slide3 = presentation.AddSlide();
            ConfigureSlide(slide3, content, "Original Slide 3", "This slide will move to the front.");

            PowerPointSlide duplicate = presentation.DuplicateSlide(0);
            PowerPointTextBox duplicateBadge = duplicate.AddTextBoxCm(
                "Duplicate of Slide 1",
                content.LeftCm, content.TopCm + 3.2, content.WidthCm, 0.8);
            duplicateBadge.ApplyTextStyle(PowerPointTextStyle.Caption.WithColor("1F4E79"));
            duplicateBadge.SetTextMarginsCm(0.2, 0.1, 0.2, 0.1);

            slide2.Hidden = true;

            presentation.MoveSlide(3, 0);

            for (int i = 0; i < presentation.Slides.Count; i++) {
                PowerPointSlide slide = presentation.Slides[i];
                string hiddenLabel = slide.Hidden ? " (hidden)" : string.Empty;
                PowerPointTextBox badge = slide.AddTextBoxCm(
                    $"Final position: {i + 1}{hiddenLabel}",
                    content.LeftCm, content.TopCm + 4.2, content.WidthCm, 0.8);
                badge.ApplyTextStyle(PowerPointTextStyle.Caption.WithColor("666666"));
                badge.SetTextMarginsCm(0.2, 0.1, 0.2, 0.1);
            }

            presentation.Save();
            Helpers.Open(filePath, openPowerPoint);
        }

        private static void ConfigureSlide(PowerPointSlide slide, PowerPointLayoutBox content, string title, string body) {
            PowerPointTextBox titleBox = slide.AddTitleCm(title, content.LeftCm, content.TopCm, content.WidthCm, 1.4);
            if (titleBox.Paragraphs.Count > 0) {
                PowerPointTextStyle.Title.WithColor("1F4E79").Apply(titleBox.Paragraphs[0]);
            }

            PowerPointTextBox description = slide.AddTextBoxCm(body,
                content.LeftCm, content.TopCm + 1.9, content.WidthCm, 1.2);
            description.ApplyTextStyle(PowerPointTextStyle.Body);
            description.SetTextMarginsCm(0.2, 0.1, 0.2, 0.1);
        }
    }
}
