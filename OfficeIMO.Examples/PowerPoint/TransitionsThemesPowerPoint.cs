using System;
using System.Collections.Generic;
using System.IO;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.Examples.PowerPoint {
    /// <summary>
    /// Demonstrates the default theme and slide transitions.
    /// </summary>
    public static class TransitionsThemesPowerPoint {
        public static void Example_TransitionsThemes(string folderPath, bool openPowerPoint) {
            Console.WriteLine("[*] PowerPoint - Transitions and themes");
            string filePath = Path.Combine(folderPath, "Transitions and Themes.pptx");
            using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
            const double marginCm = 1.5;
            const double titleHeightCm = 1.4;
            const double bodyGapCm = 0.8;
            PowerPointLayoutBox content = presentation.SlideSize.GetContentBoxCm(marginCm);
            double bodyTopCm = content.Top.ToCentimeters() + titleHeightCm + bodyGapCm;
            double bodyHeightCm = content.Height.ToCentimeters() - titleHeightCm - bodyGapCm;

            PowerPointSlide intro = presentation.AddSlide();
            PowerPointTextBox introTitle = intro.AddTitleCm("Transitions & Themes", marginCm, marginCm, content.Width.ToCentimeters(), titleHeightCm);
            introTitle.FontSize = 32;
            introTitle.Color = "1F4E79";
            intro.AddTextBoxCm("Built with the default Office theme and slide transitions.", marginCm, 3.2, content.Width.ToCentimeters(), 1.1);
            intro.AddTextBoxCm($"Theme: {presentation.ThemeName}", marginCm, 4.6, content.Width.ToCentimeters(), 1.0);
            intro.Transition = PowerPointSlideTransition.Fade;

            IReadOnlyList<PowerPointSlideTransition> transitions = new[] {
                PowerPointSlideTransition.Fade,
                PowerPointSlideTransition.Wipe,
                PowerPointSlideTransition.BlindsVertical,
                PowerPointSlideTransition.BlindsHorizontal,
                PowerPointSlideTransition.CombHorizontal,
                PowerPointSlideTransition.CombVertical,
                PowerPointSlideTransition.PushLeft,
                PowerPointSlideTransition.PushRight,
                PowerPointSlideTransition.PushUp,
                PowerPointSlideTransition.PushDown,
                PowerPointSlideTransition.Cut,
                PowerPointSlideTransition.Flash,
                PowerPointSlideTransition.WarpIn,
                PowerPointSlideTransition.WarpOut,
                PowerPointSlideTransition.Prism,
                PowerPointSlideTransition.FerrisLeft,
                PowerPointSlideTransition.FerrisRight,
                PowerPointSlideTransition.Morph
            };

            foreach (PowerPointSlideTransition transition in transitions) {
                PowerPointSlide slide = presentation.AddSlide();
                slide.Transition = transition;

                PowerPointTextBox title = slide.AddTitleCm($"{transition} transition", marginCm, marginCm, content.Width.ToCentimeters(), titleHeightCm);
                title.FontSize = 28;
                title.Color = "1F4E79";

                PowerPointLayoutBox body = new OfficeIMO.PowerPoint.PowerPointLayoutBox(OfficeIMO.OfficeLength.Centimeters(content.Left.ToCentimeters()),OfficeIMO.OfficeLength.Centimeters(bodyTopCm),OfficeIMO.OfficeLength.Centimeters(content.Width.ToCentimeters()),OfficeIMO.OfficeLength.Centimeters(bodyHeightCm));
                PowerPointLayoutBox[] columns = body.SplitColumns(2,OfficeIMO.OfficeLength.Centimeters(1.0));

                PowerPointAutoShape leftPanel = slide.AddRectangleCm(columns[0].Left.ToCentimeters(), columns[0].Top.ToCentimeters(), columns[0].Width.ToCentimeters(),
                    columns[0].Height.ToCentimeters(), "Left Panel");
                leftPanel.Fill("E7F7FF").Stroke("007ACC", 1.5);

                PowerPointAutoShape rightPanel = slide.AddRectangleCm(columns[1].Left.ToCentimeters(), columns[1].Top.ToCentimeters(), columns[1].Width.ToCentimeters(),
                    columns[1].Height.ToCentimeters(), "Right Panel");
                rightPanel.Fill("FFF4E5").Stroke("C48A00", 1.5);

                PowerPointTextBox leftLabel = slide.AddTextBoxCm("Panel A", columns[0].Left.ToCentimeters() + 0.4, columns[0].Top.ToCentimeters() + 0.4,
                    columns[0].Width.ToCentimeters() - 0.8, 1.0);
                leftLabel.FontSize = 16;
                leftLabel.Color = "1F4E79";

                PowerPointTextBox rightLabel = slide.AddTextBoxCm("Panel B", columns[1].Left.ToCentimeters() + 0.4, columns[1].Top.ToCentimeters() + 0.4,
                    columns[1].Width.ToCentimeters() - 0.8, 1.0);
                rightLabel.FontSize = 16;
                rightLabel.Color = "1F4E79";

                PowerPointTextBox note = slide.AddTextBoxCm($"SlideTransition.{transition}", marginCm, bodyTopCm - 0.6,
                    content.Width.ToCentimeters(), 0.6);
                note.FontSize = 14;
                note.Color = "666666";
            }

            presentation.Save();
            if (openPowerPoint) ExampleFileLauncher.Open(filePath);
        }
    }
}
