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
            using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
            const double marginCm = 1.5;
            PowerPointLayoutBox content = presentation.SlideSize.GetContentBox(OfficeIMO.OfficeLength.Centimeters(marginCm));

            PowerPointSlide slide1 = presentation.AddSlide();
            ConfigureSlide(slide1, content, "Original Slide 1", "This slide is created first.");

            PowerPointSlide slide2 = presentation.AddSlide();
            ConfigureSlide(slide2, content, "Original Slide 2", "This slide will be hidden in slide show.");

            PowerPointSlide slide3 = presentation.AddSlide();
            ConfigureSlide(slide3, content, "Original Slide 3", "This slide will move to the front.");

            PowerPointSlide duplicate = presentation.DuplicateSlide(0);
            PowerPointTextBox duplicateBadge = duplicate.AddTextBox(
                "Duplicate of Slide 1",                global::OfficeIMO.OfficeLength.Centimeters(content.Left.ToCentimeters()),global::OfficeIMO.OfficeLength.Centimeters(content.Top.ToCentimeters() + 3.2),global::OfficeIMO.OfficeLength.Centimeters(content.Width.ToCentimeters()),global::OfficeIMO.OfficeLength.Centimeters(0.8));
            duplicateBadge.ApplyTextStyle(PowerPointTextStyle.Caption.WithColor("1F4E79"));
            duplicateBadge.SetTextMargins(global::OfficeIMO.OfficeLength.Centimeters(0.2), global::OfficeIMO.OfficeLength.Centimeters(0.1), global::OfficeIMO.OfficeLength.Centimeters(0.2), global::OfficeIMO.OfficeLength.Centimeters(0.1));

            slide2.Hidden = true;

            string sourcePath = Path.Combine(folderPath, "Slides Management - Source.pptx");
            using (PowerPointDocument sourcePresentation = PowerPointDocument.Create(sourcePath)) {
                PowerPointLayoutBox sourceContent = sourcePresentation.SlideSize.GetContentBox(OfficeIMO.OfficeLength.Centimeters(marginCm));
                PowerPointSlide sourceSlide = sourcePresentation.AddSlide();
                ConfigureSlide(sourceSlide, sourceContent, "Imported Slide", "This slide is imported from another presentation.");
                PowerPointTextBox sourceBadge = sourceSlide.AddTextBox(
                    "Source deck",                    global::OfficeIMO.OfficeLength.Centimeters(sourceContent.Left.ToCentimeters()),global::OfficeIMO.OfficeLength.Centimeters(sourceContent.Top.ToCentimeters() + 3.2),global::OfficeIMO.OfficeLength.Centimeters(sourceContent.Width.ToCentimeters()),global::OfficeIMO.OfficeLength.Centimeters(0.8));
                sourceBadge.ApplyTextStyle(PowerPointTextStyle.Caption.WithColor("1F4E79"));
                sourceBadge.SetTextMargins(global::OfficeIMO.OfficeLength.Centimeters(0.2), global::OfficeIMO.OfficeLength.Centimeters(0.1), global::OfficeIMO.OfficeLength.Centimeters(0.2), global::OfficeIMO.OfficeLength.Centimeters(0.1));
                sourcePresentation.Save();

                PowerPointSlide imported = presentation.ImportSlide(sourcePresentation, 0, insertAt: 1);
                PowerPointTextBox importBadge = imported.AddTextBox(
                    "Imported from Slides Management - Source.pptx",                    global::OfficeIMO.OfficeLength.Centimeters(content.Left.ToCentimeters()),global::OfficeIMO.OfficeLength.Centimeters(content.Top.ToCentimeters() + 3.2),global::OfficeIMO.OfficeLength.Centimeters(content.Width.ToCentimeters()),global::OfficeIMO.OfficeLength.Centimeters(0.8));
                importBadge.ApplyTextStyle(PowerPointTextStyle.Caption.WithColor("1F4E79"));
                importBadge.SetTextMargins(global::OfficeIMO.OfficeLength.Centimeters(0.2), global::OfficeIMO.OfficeLength.Centimeters(0.1), global::OfficeIMO.OfficeLength.Centimeters(0.2), global::OfficeIMO.OfficeLength.Centimeters(0.1));
            }

            presentation.MoveSlide(3, 0);

            for (int i = 0; i < presentation.Slides.Count; i++) {
                PowerPointSlide slide = presentation.Slides[i];
                string hiddenLabel = slide.Hidden ? " (hidden)" : string.Empty;
                PowerPointTextBox badge = slide.AddTextBox(
                    $"Final position: {i + 1}{hiddenLabel}",                    global::OfficeIMO.OfficeLength.Centimeters(content.Left.ToCentimeters()),global::OfficeIMO.OfficeLength.Centimeters(content.Top.ToCentimeters() + 4.2),global::OfficeIMO.OfficeLength.Centimeters(content.Width.ToCentimeters()),global::OfficeIMO.OfficeLength.Centimeters(0.8));
                badge.ApplyTextStyle(PowerPointTextStyle.Caption.WithColor("666666"));
                badge.SetTextMargins(global::OfficeIMO.OfficeLength.Centimeters(0.2), global::OfficeIMO.OfficeLength.Centimeters(0.1), global::OfficeIMO.OfficeLength.Centimeters(0.2), global::OfficeIMO.OfficeLength.Centimeters(0.1));
            }

            presentation.Save();
            if (openPowerPoint) ExampleFileLauncher.Open(filePath);
        }

        private static void ConfigureSlide(PowerPointSlide slide, PowerPointLayoutBox content, string title, string body) {
            PowerPointTextBox titleBox = slide.AddTitle(title,global::OfficeIMO.OfficeLength.Centimeters(content.Left.ToCentimeters()),global::OfficeIMO.OfficeLength.Centimeters(content.Top.ToCentimeters()),global::OfficeIMO.OfficeLength.Centimeters(content.Width.ToCentimeters()),global::OfficeIMO.OfficeLength.Centimeters(1.4));
            if (titleBox.Paragraphs.Count > 0) {
                PowerPointTextStyle.Title.WithColor("1F4E79").Apply(titleBox.Paragraphs[0]);
            }

            PowerPointTextBox description = slide.AddTextBox(body,                global::OfficeIMO.OfficeLength.Centimeters(content.Left.ToCentimeters()),global::OfficeIMO.OfficeLength.Centimeters(content.Top.ToCentimeters() + 1.9),global::OfficeIMO.OfficeLength.Centimeters(content.Width.ToCentimeters()),global::OfficeIMO.OfficeLength.Centimeters(1.2));
            description.ApplyTextStyle(PowerPointTextStyle.Body);
            description.SetTextMargins(global::OfficeIMO.OfficeLength.Centimeters(0.2), global::OfficeIMO.OfficeLength.Centimeters(0.1), global::OfficeIMO.OfficeLength.Centimeters(0.2), global::OfficeIMO.OfficeLength.Centimeters(0.1));
        }
    }
}
