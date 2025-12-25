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
            using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
            const double marginCm = 1.5;
            const double titleHeightCm = 1.4;
            const double bodyGapCm = 0.8;
            PowerPointLayoutBox content = presentation.SlideSize.GetContentBoxCm(marginCm);
            double bodyTopCm = content.TopCm + titleHeightCm + bodyGapCm;
            double bodyHeightCm = content.HeightCm - titleHeightCm - bodyGapCm;

            PowerPointSlide intro = presentation.AddSlide();
            PowerPointTextBox introTitle = intro.AddTitleCm("Transitions & Themes", marginCm, marginCm, content.WidthCm, titleHeightCm);
            introTitle.FontSize = 32;
            introTitle.Color = "1F4E79";
            intro.AddTextBoxCm("Built with the default Office theme and slide transitions.", marginCm, 3.2, content.WidthCm, 1.1);
            intro.AddTextBoxCm($"Theme: {presentation.ThemeName}", marginCm, 4.6, content.WidthCm, 1.0);
            intro.Transition = SlideTransition.Fade;

            IReadOnlyList<SlideTransition> transitions = new[] {
                SlideTransition.Fade,
                SlideTransition.Wipe,
                SlideTransition.BlindsVertical,
                SlideTransition.BlindsHorizontal,
                SlideTransition.CombHorizontal,
                SlideTransition.CombVertical,
                SlideTransition.PushLeft,
                SlideTransition.PushRight,
                SlideTransition.PushUp,
                SlideTransition.PushDown,
                SlideTransition.Cut,
                SlideTransition.Flash,
                SlideTransition.WarpIn,
                SlideTransition.WarpOut,
                SlideTransition.Prism,
                SlideTransition.FerrisLeft,
                SlideTransition.FerrisRight,
                SlideTransition.Morph
            };

            foreach (SlideTransition transition in transitions) {
                PowerPointSlide slide = presentation.AddSlide();
                slide.Transition = transition;

                PowerPointTextBox title = slide.AddTitleCm($"{transition} transition", marginCm, marginCm, content.WidthCm, titleHeightCm);
                title.FontSize = 28;
                title.Color = "1F4E79";

                PowerPointLayoutBox body = PowerPointLayoutBox.FromCentimeters(content.LeftCm, bodyTopCm, content.WidthCm, bodyHeightCm);
                PowerPointLayoutBox[] columns = body.SplitColumnsCm(2, 1.0);

                PowerPointAutoShape leftPanel = slide.AddRectangleCm(columns[0].LeftCm, columns[0].TopCm, columns[0].WidthCm,
                    columns[0].HeightCm, "Left Panel");
                leftPanel.Fill("E7F7FF").Stroke("007ACC", 1.5);

                PowerPointAutoShape rightPanel = slide.AddRectangleCm(columns[1].LeftCm, columns[1].TopCm, columns[1].WidthCm,
                    columns[1].HeightCm, "Right Panel");
                rightPanel.Fill("FFF4E5").Stroke("C48A00", 1.5);

                PowerPointTextBox leftLabel = slide.AddTextBoxCm("Panel A", columns[0].LeftCm + 0.4, columns[0].TopCm + 0.4,
                    columns[0].WidthCm - 0.8, 1.0);
                leftLabel.FontSize = 16;
                leftLabel.Color = "1F4E79";

                PowerPointTextBox rightLabel = slide.AddTextBoxCm("Panel B", columns[1].LeftCm + 0.4, columns[1].TopCm + 0.4,
                    columns[1].WidthCm - 0.8, 1.0);
                rightLabel.FontSize = 16;
                rightLabel.Color = "1F4E79";

                PowerPointTextBox note = slide.AddTextBoxCm($"SlideTransition.{transition}", marginCm, bodyTopCm - 0.6,
                    content.WidthCm, 0.6);
                note.FontSize = 14;
                note.Color = "666666";
            }

            presentation.Save();
            Helpers.Open(filePath, openPowerPoint);
        }
    }
}
