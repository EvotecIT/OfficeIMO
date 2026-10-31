using System;
using System.IO;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.Examples.PowerPoint {
    /// <summary>
    /// Demonstrates that notes masters are referenced correctly when creating a presentation.
    /// </summary>
    public static class NotesMasterPowerPoint {
        public static void Example_PowerPointNotesMaster(string folderPath, bool openPowerPoint) {
            Console.WriteLine("[*] PowerPoint - Notes master");
            string filePath = Path.Combine(folderPath, "NotesMaster.pptx");
            using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
            const double marginCm = 1.5;
            PowerPointLayoutBox content = presentation.SlideSize.GetContentBox(OfficeIMO.OfficeLength.Centimeters(marginCm));

            PowerPointSlide slide = presentation.AddSlide();
            PowerPointTextBox title = slide.AddTitle("Speaker Notes",global::OfficeIMO.OfficeLength.Centimeters(content.Left.ToCentimeters()),global::OfficeIMO.OfficeLength.Centimeters(content.Top.ToCentimeters()),global::OfficeIMO.OfficeLength.Centimeters(content.Width.ToCentimeters()),global::OfficeIMO.OfficeLength.Centimeters(1.4));
            if (title.Paragraphs.Count > 0) {
                PowerPointTextStyle.Title.WithColor("1F4E79").Apply(title.Paragraphs[0]);
            }

            PowerPointTextBox body = slide.AddTextBox(
                "This slide includes speaker notes that remain hidden during presentation.",                global::OfficeIMO.OfficeLength.Centimeters(content.Left.ToCentimeters()),global::OfficeIMO.OfficeLength.Centimeters(content.Top.ToCentimeters() + 1.9),global::OfficeIMO.OfficeLength.Centimeters(content.Width.ToCentimeters()),global::OfficeIMO.OfficeLength.Centimeters(1.6));
            body.ApplyTextStyle(PowerPointTextStyle.Body);
            body.SetTextMargins(global::OfficeIMO.OfficeLength.Centimeters(0.2), global::OfficeIMO.OfficeLength.Centimeters(0.2), global::OfficeIMO.OfficeLength.Centimeters(0.2), global::OfficeIMO.OfficeLength.Centimeters(0.2));

            slide.Notes.Text = "Notes are stored in the notes section of the slide.\n" +
                               "Use them for presenter-only reminders.";

            presentation.Save();
            if (openPowerPoint) ExampleFileLauncher.Open(filePath);
        }
    }
}
