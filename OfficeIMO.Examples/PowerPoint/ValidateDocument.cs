using System;
using System.IO;
using OfficeIMO.PowerPoint;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Examples.PowerPoint {
    /// <summary>
    /// Demonstrates validating a PowerPoint presentation.
    /// </summary>
    public static class ValidateDocument {
        public static void Example(string folderPath, bool openPowerPoint) {
            Console.WriteLine("[*] PowerPoint - Validate document");
            string filePath = Path.Combine(folderPath, "ValidateDocument.pptx");
            const double marginCm = 1.5;
            const double titleHeightCm = 1.3;
            const double bodyGapCm = 0.8;

            using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
            PowerPointLayoutBox content = presentation.SlideSize.GetContentBoxCm(marginCm);
            PowerPointLayoutBox titleBox = PowerPointLayoutBox.FromCentimeters(
                content.LeftCm, content.TopCm, content.WidthCm, titleHeightCm);
            PowerPointLayoutBox bodyBox = PowerPointLayoutBox.FromCentimeters(
                content.LeftCm,
                content.TopCm + titleHeightCm + bodyGapCm,
                content.WidthCm,
                content.HeightCm - titleHeightCm - bodyGapCm);

            PowerPointSlide slide = presentation.AddSlide();
            slide.AddTitle("Validation Example", titleBox);
            PowerPointTextBox intro = slide.AddTextBox(
                "Creates a deck, saves it, and validates the Open XML package.",
                bodyBox);
            intro.ApplyTextStyle(PowerPointTextStyle.Body.WithColor("1F4E79"));
            intro.TextAutoFit = PowerPointTextAutoFit.Normal;
            intro.SetTextMarginsCm(0.3, 0.2, 0.3, 0.2);

            presentation.Save();
            var errors = presentation.ValidateDocument();

            PowerPointSlide results = presentation.AddSlide();
            results.AddTitle("Validation Result", titleBox);
            string status = errors.Count == 0 ? "Pass" : "Fail";
            string details = errors.Count == 0 ? "No validation errors found." : $"{errors.Count} validation issue(s).";

            var summaryRows = new[] {
                new ValidationRow { Check = "Open XML validation", Status = status, Details = details },
                new ValidationRow { Check = "Slides generated", Status = "Info", Details = $"{presentation.Slides.Count} slides" },
                new ValidationRow { Check = "Theme", Status = "Info", Details = presentation.ThemeName ?? "Default" }
            };

            PowerPointTable table = results.AddTableCm(
                summaryRows,
                o => {
                    o.HeaderCase = HeaderCase.Title;
                    o.PinFirst("Check");
                },
                includeHeaders: true,
                leftCm: bodyBox.LeftCm,
                topCm: bodyBox.TopCm,
                widthCm: bodyBox.WidthCm,
                heightCm: bodyBox.HeightCm);
            table.ApplyStyle(PowerPointTableStylePreset.Default);
            table.SetColumnWidthsEvenly();
            for (int c = 0; c < table.Columns; c++) {
                PowerPointTableCell header = table.GetCell(0, c);
                header.Bold = true;
                header.HorizontalAlignment = A.TextAlignmentTypeValues.Center;
                header.VerticalAlignment = A.TextAnchoringTypeValues.Center;
            }

            presentation.Save();
            Helpers.Open(filePath, openPowerPoint);
        }

        private sealed class ValidationRow {
            public string Check { get; set; } = string.Empty;
            public string Status { get; set; } = string.Empty;
            public string Details { get; set; } = string.Empty;
        }
    }
}
