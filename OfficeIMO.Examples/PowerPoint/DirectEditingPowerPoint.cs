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

            using (PowerPointDocument source = PowerPointDocument.Create(sourcePath)) {
                PowerPointSlide sourceSlide = source.AddSlide();
                sourceSlide.AddTitle("Imported slide",global::OfficeIMO.OfficeLength.Centimeters(1.5),global::OfficeIMO.OfficeLength.Centimeters(1.5),global::OfficeIMO.OfficeLength.Centimeters(22.0),global::OfficeIMO.OfficeLength.Centimeters(1.5));
                sourceSlide.AddTextBox("This slide remains fully editable after import.",global::OfficeIMO.OfficeLength.Centimeters(1.5),global::OfficeIMO.OfficeLength.Centimeters(3.5),global::OfficeIMO.OfficeLength.Centimeters(22.0),global::OfficeIMO.OfficeLength.Centimeters(1.2));
                source.Save();
            }

            using (PowerPointDocument presentation = PowerPointDocument.Create(filePath)) {
                PowerPointLayoutBox content = presentation.SlideSize.GetContentBox(OfficeIMO.OfficeLength.Centimeters(1.5));
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointTextBox title = slide.AddTitle("Direct presentation editing",                    new OfficeIMO.PowerPoint.PowerPointLayoutBox(OfficeIMO.OfficeLength.Centimeters(content.Left.ToCentimeters()),OfficeIMO.OfficeLength.Centimeters(content.Top.ToCentimeters()),OfficeIMO.OfficeLength.Centimeters(content.Width.ToCentimeters()),OfficeIMO.OfficeLength.Centimeters(1.6)));
                title.FontSize = 32;
                title.Color = "1F4E79";

                PowerPointTextBox bullets = slide.AddTextBox(string.Empty,                    new OfficeIMO.PowerPoint.PowerPointLayoutBox(OfficeIMO.OfficeLength.Centimeters(content.Left.ToCentimeters()),OfficeIMO.OfficeLength.Centimeters(content.Top.ToCentimeters() + 2.1),OfficeIMO.OfficeLength.Centimeters(11.0),OfficeIMO.OfficeLength.Centimeters(4.0)));
                bullets.AddBullet("One lifecycle owner");
                bullets.AddBullet("Concrete slide and shape editing");
                bullets.AddBullet("No parallel builder vocabulary");
                bullets.TextAutoFit = PowerPointTextAutoFit.Normal;

                slide.AddShape(OfficePresetShapeType.Rectangle,
                        PowerPointUnits.FromCentimeters(content.Left.ToCentimeters() + 12.0), PowerPointUnits.FromCentimeters(content.Top.ToCentimeters() + 2.1),
                        PowerPointUnits.FromCentimeters(10.0), PowerPointUnits.FromCentimeters(4.0))
                    .Fill("E7F7FF")
                    .Stroke("007ACC", 2);
                slide.Notes.Text = "Concrete objects can be edited immediately or after reopening the deck.";

                using (PowerPointDocument source = PowerPointDocument.Load(sourcePath,
                           new PowerPointLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                    PowerPointSlide imported = presentation.ImportSlide(source, 0);
                    imported.AddTextBox("Imported through the same presentation model.",global::OfficeIMO.OfficeLength.Centimeters(1.5),global::OfficeIMO.OfficeLength.Centimeters(5.0),global::OfficeIMO.OfficeLength.Centimeters(22.0),global::OfficeIMO.OfficeLength.Centimeters(0.8));
                }

                PowerPointSlide duplicate = presentation.DuplicateSlide(0);
                duplicate.Hidden = true;
                duplicate.ReplaceText("Direct presentation editing", "Hidden duplicate");
                presentation.MoveSlide(1, presentation.Slides.Count - 1);
                presentation.Save();
            }

            using (PowerPointDocument edited = PowerPointDocument.Load(filePath)) {
                edited.ReplaceText("One lifecycle owner", "One consistent lifecycle owner");
                edited.Save();
            }

            if (openPowerPoint) ExampleFileLauncher.Open(filePath);
        }
    }
}
