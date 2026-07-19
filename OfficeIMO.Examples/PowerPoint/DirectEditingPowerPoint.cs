using System;
using System.IO;
using OfficeIMO.PowerPoint;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Examples.PowerPoint {
    /// <summary>Demonstrates the canonical direct editing API.</summary>
    public static class DirectEditingPowerPoint {
        public static void Example(string folderPath, bool openPowerPoint) {
            Console.WriteLine("[*] PowerPoint - Creating and editing a presentation");
            string filePath = Path.Combine(folderPath, "DirectEditingPowerPoint.pptx");
            string sourcePath = Path.Combine(folderPath, "DirectEditingPowerPoint-Source.pptx");

            using (PowerPointPresentation source = PowerPointPresentation.Create(sourcePath)) {
                PowerPointSlide sourceSlide = source.AddSlide();
                sourceSlide.AddTitleCm("Imported slide", 1.5, 1.5, 22.0, 1.5);
                sourceSlide.AddTextBoxCm("This slide remains fully editable after import.", 1.5, 3.5, 22.0, 1.2);
                source.Save();
            }

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                PowerPointLayoutBox content = presentation.SlideSize.GetContentBoxCm(1.5);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointTextBox title = slide.AddTitle("Direct presentation editing",
                    PowerPointLayoutBox.FromCentimeters(content.LeftCm, content.TopCm, content.WidthCm, 1.6));
                title.FontSize = 32;
                title.Color = "1F4E79";

                PowerPointTextBox bullets = slide.AddTextBox(string.Empty,
                    PowerPointLayoutBox.FromCentimeters(content.LeftCm, content.TopCm + 2.1, 11.0, 4.0));
                bullets.AddBullet("One lifecycle owner");
                bullets.AddBullet("Concrete slide and shape editing");
                bullets.AddBullet("No parallel builder vocabulary");
                bullets.TextAutoFit = PowerPointTextAutoFit.Normal;

                slide.AddShape(A.ShapeTypeValues.Rectangle,
                        PowerPointUnits.FromCentimeters(content.LeftCm + 12.0), PowerPointUnits.FromCentimeters(content.TopCm + 2.1),
                        PowerPointUnits.FromCentimeters(10.0), PowerPointUnits.FromCentimeters(4.0))
                    .Fill("E7F7FF")
                    .Stroke("007ACC", 2);
                slide.Notes.Text = "Concrete objects can be edited immediately or after reopening the deck.";

                using (PowerPointPresentation source = PowerPointPresentation.Load(sourcePath,
                           new PowerPointLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                    PowerPointSlide imported = presentation.ImportSlide(source, 0);
                    imported.AddTextBoxCm("Imported through the same presentation model.", 1.5, 5.0, 22.0, 0.8);
                }

                PowerPointSlide duplicate = presentation.DuplicateSlide(0);
                duplicate.Hidden = true;
                duplicate.ReplaceText("Direct presentation editing", "Hidden duplicate");
                presentation.MoveSlide(1, presentation.Slides.Count - 1);
                presentation.Save();
            }

            using (PowerPointPresentation edited = PowerPointPresentation.Load(filePath)) {
                edited.ReplaceText("One lifecycle owner", "One consistent lifecycle owner");
                edited.Save();
            }

            if (openPowerPoint) ExampleFileLauncher.Open(filePath);
        }
    }
}
