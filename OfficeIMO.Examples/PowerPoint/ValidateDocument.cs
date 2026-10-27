using System;
using System.IO;
using OfficeIMO.PowerPoint;
using OfficeIMO.Drawing;
using OfficeIMO.Data;
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

            using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
            PowerPointLayoutBox content = presentation.SlideSize.GetContentBoxCm(marginCm);
            PowerPointLayoutBox titleBox = new OfficeIMO.PowerPoint.PowerPointLayoutBox(
                OfficeIMO.OfficeLength.Centimeters(content.Left.ToCentimeters()),OfficeIMO.OfficeLength.Centimeters(content.Top.ToCentimeters()),OfficeIMO.OfficeLength.Centimeters(content.Width.ToCentimeters()),OfficeIMO.OfficeLength.Centimeters(titleHeightCm));
            PowerPointLayoutBox bodyBox = new OfficeIMO.PowerPoint.PowerPointLayoutBox(
                OfficeIMO.OfficeLength.Centimeters(content.Left.ToCentimeters()),                OfficeIMO.OfficeLength.Centimeters(content.Top.ToCentimeters() + titleHeightCm + bodyGapCm),                OfficeIMO.OfficeLength.Centimeters(content.Width.ToCentimeters()),                OfficeIMO.OfficeLength.Centimeters(content.Height.ToCentimeters() - titleHeightCm - bodyGapCm));

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
                leftCm: bodyBox.Left.ToCentimeters(),
                topCm: bodyBox.Top.ToCentimeters(),
                widthCm: bodyBox.Width.ToCentimeters(),
                heightCm: bodyBox.Height.ToCentimeters());
            table.ApplyStyle(PowerPointTableStylePreset.Default);
            table.SetColumnWidthsEvenly();
            for (int c = 0; c < table.Columns; c++) {
                PowerPointTableCell header = table.GetCell(0, c);
                header.Bold = true;
                header.HorizontalAlignment = PowerPointTextAlignment.Center;
                header.VerticalAlignment = PowerPointTextVerticalAlignment.Center;
            }

            presentation.Save();
            if (openPowerPoint) ExampleFileLauncher.Open(filePath);
        }

        private sealed class ValidationRow {
            public string Check { get; set; } = string.Empty;
            public string Status { get; set; } = string.Empty;
            public string Details { get; set; } = string.Empty;
        }
    }
}
