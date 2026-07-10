using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;

namespace OfficeIMO.Examples.PowerPoint {
    /// <summary>
    /// Generates one inspectable proof bundle: editable PPTX, PNG/SVG previews, PDF, and JSON preflight.
    /// </summary>
    public static class EndToEndPowerPointProof {
        public static void Example_EndToEndPowerPointProof(string folderPath, bool openPowerPoint) {
            Console.WriteLine("[*] PowerPoint - end-to-end proof bundle");
            string outputFolder = Path.Combine(folderPath, "PowerPoint End To End Proof");
            Directory.CreateDirectory(outputFolder);
            string presentationPath = Path.Combine(outputFolder, "Executive Delivery Review.pptx");
            string pdfPath = Path.Combine(outputFolder, "Executive Delivery Review.pdf");
            string reportPath = Path.Combine(outputFolder, "Executive Delivery Review.preflight.json");

            using PowerPointPresentation presentation = PowerPointPresentation.Create(presentationPath);
            presentation.SlideSize.SetPreset(PowerPointSlideSizePreset.Screen16x9);
            PowerPointDesignBrief brief = PowerPointDesignBrief
                .FromBrand("#0B7FAB", "e2e-proof", "executive delivery review")
                .WithIdentity("Delivery Review", eyebrow: "OFFICEIMO", footerLeft: "DELIVERY",
                    footerRight: "Generated proof");
            PowerPointDeckComposer deck = presentation.UseDesigner(brief, alternativeIndex: 0);
            PowerPointDeckPlan plan = new PowerPointDeckPlan()
                .AddSection("Executive delivery review", "Editable, measured, and ready for inspection")
                .AddProcess("Delivery path", "Long content continues deterministically",
                    Enumerable.Range(1, 8).Select(index => new PowerPointProcessStep(
                        "Phase " + index, "Evidence and owner for delivery phase " + index + ".")))
                .AddCardGrid("Decision signals", "A semantic story rendered as native shapes",
                    Enumerable.Range(1, 8).Select(index => new PowerPointCardContent(
                        "Signal " + index, new[] { "Owner assigned", "Evidence available" })));

            deck.AddSlidesWithContinuation(plan);

            DeliveryRow[] rows = Enumerable.Range(1, 17)
                .Select(index => new DeliveryRow("Workstream " + index,
                    index % 3 == 0 ? "At risk" : "On track", 50 + index * 2))
                .ToArray();
            var columns = new[] {
                PowerPointTableColumn<DeliveryRow>.Create("Workstream", row => row.Name).WithWidthCm(9),
                PowerPointTableColumn<DeliveryRow>.Create("Status", row => row.Status).WithWidthCm(5),
                PowerPointTableColumn<DeliveryRow>.Create("Complete", row => row.Percent + "%")
            };
            presentation.AddTableSlides(rows, columns, new PowerPointTablePaginationOptions {
                TableBounds = PowerPointLayoutBox.FromCentimeters(1.5, 3.0, 22.4, 9.5),
                MinimumRowHeightPoints = 27,
                ConfigureSlide = (slide, context) => {
                    PowerPointTextBox title = slide.AddTitleCm(
                        context.IsContinuation ? "Delivery detail (continued)" : "Delivery detail",
                        1.5, 1.0, 22.4, 1.3);
                    if (title.Paragraphs.Count > 0) {
                        PowerPointTextStyle.Title.WithColor("0B3954").Apply(title.Paragraphs[0]);
                    }
                    slide.Notes.Text = "Table page " + (context.PageIndex + 1) + " of " + context.PageCount + ".";
                },
                ConfigureTable = (table, _) => table.BandedRows = true
            });

            var preflightOptions = new PowerPointDeckPreflightOptions {
                MinimumReadableFontSizePoints = 8,
                DetectShapeCollisions = false
            };
            PowerPointDeckPreflightReport report = presentation.Preflight(preflightOptions);
            report.SaveJson(reportPath);
            presentation.Save();

            for (int slideIndex = 0; slideIndex < presentation.Slides.Count; slideIndex++) {
                string stem = "Slide " + (slideIndex + 1).ToString("00");
                presentation.Slides[slideIndex].SaveAsPng(Path.Combine(outputFolder, stem + ".png"));
                presentation.Slides[slideIndex].SaveAsSvg(Path.Combine(outputFolder, stem + ".svg"));
            }
            presentation.SaveAsPdf(pdfPath);

            List<ValidationErrorInfo> errors = presentation.ValidateDocument();
            if (errors.Count > 0) {
                throw new InvalidOperationException("End-to-end proof deck has " + errors.Count +
                    " Open XML validation error(s): " +
                    string.Join("; ", errors.Take(3).Select(error => error.Description)));
            }

            Console.WriteLine("    Deck: " + presentationPath);
            Console.WriteLine("    Slides: " + presentation.Slides.Count);
            Console.WriteLine("    Preflight: " + report.ErrorCount + " errors, " + report.WarningCount + " warnings");
            Console.WriteLine("    Proof: PNG, SVG, PDF, JSON, and Open XML validation");
            Helpers.Open(presentationPath, openPowerPoint);
        }

        private sealed class DeliveryRow {
            internal DeliveryRow(string name, string status, int percent) {
                Name = name;
                Status = status;
                Percent = percent;
            }

            internal string Name { get; }
            internal string Status { get; }
            internal int Percent { get; }
        }
    }
}
