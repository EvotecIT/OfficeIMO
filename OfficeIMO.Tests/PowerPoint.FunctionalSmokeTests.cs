using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using PptImagePartType = OfficeIMO.PowerPoint.ImagePartType;

namespace OfficeIMO.Tests {
    public class PowerPointFunctionalSmokeTests {
        private static readonly byte[] OnePixelPng = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMB/6X4nKkAAAAASUVORK5CYII=");

        [Fact]
        public void CanBuildRichDeckAndValidate() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    presentation.SetThemeColorForAllMasters(PowerPointThemeColor.Accent1, "4472C4");

                    PowerPointSlide slide = presentation.AddSlide(SlideLayoutValues.TitleOnly);
                    slide.AddTitle("Functional Smoke Test");
                    PowerPointTextBox box = slide.AddTextBox("Agenda", PowerPointUnits.Cm(1), PowerPointUnits.Cm(3),
                        PowerPointUnits.Cm(6), PowerPointUnits.Cm(3));
                    box.AddBullets(new[] { "Shapes", "Images", "Tables", "Charts" });

                    PowerPointAutoShape rect = slide.AddRectangle(PowerPointUnits.Cm(1), PowerPointUnits.Cm(7),
                        PowerPointUnits.Cm(4), PowerPointUnits.Cm(2), "Card");
                    rect.FillColor = "E7F7FF";
                    rect.OutlineColor = "007ACC";

                    using (var imageStream = new MemoryStream(OnePixelPng)) {
                        slide.AddPicture(imageStream, PptImagePartType.Png,
                            PowerPointUnits.Cm(8), PowerPointUnits.Cm(1),
                            PowerPointUnits.Cm(2), PowerPointUnits.Cm(2));
                    }

                    PowerPointTableStyleInfo style = presentation.TableStyles
                        .FirstOrDefault(s => !string.IsNullOrWhiteSpace(s.StyleId));
                    Assert.False(string.IsNullOrWhiteSpace(style.StyleId));
                    string styleName = string.IsNullOrWhiteSpace(style.Name) ? style.StyleId : style.Name;

                    PowerPointTable table = slide.AddTable(rows: 2, columns: 2, styleName: styleName,
                        left: PowerPointUnits.Cm(8), top: PowerPointUnits.Cm(4),
                        width: PowerPointUnits.Cm(6), height: PowerPointUnits.Cm(3),
                        firstRow: true, bandedRows: true);
                    table.GetCell(0, 0).Text = "Header";
                    table.GetCell(1, 0).Text = "Value";
                    table.GetCell(0, 0).SetTextAutoFit(PowerPointTextAutoFit.Normal,
                        new PowerPointTextAutoFitOptions(fontScalePercent: 80, lineSpaceReductionPercent: 10));

                    PowerPointChart chart = slide.AddChart(PowerPointUnits.Cm(1), PowerPointUnits.Cm(10),
                        PowerPointUnits.Cm(10), PowerPointUnits.Cm(5));
                    chart.SetTitle("Sales");

                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    var validator = new OpenXmlValidator(FileFormatVersions.Microsoft365);
                    var errors = validator.Validate(document).ToList();
                    Assert.True(errors.Count == 0, FormatValidationErrors(errors));

                    PresentationPart part = document.PresentationPart!;
                    Assert.NotNull(part.TableStylesPart?.TableStyleList);
                    SlidePart slidePart = part.SlideParts.First();
                    Assert.True(slidePart.ImageParts.Any());
                    Assert.True(slidePart.ChartParts.Any());
                    Assert.True(slidePart.Slide.Descendants<A.Table>().Any());
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Open(filePath)) {
                    PowerPointSlide slide = presentation.Slides.First();
                    Assert.True(slide.Pictures.Any());
                    Assert.True(slide.Tables.Any());
                    Assert.True(slide.Charts.Any());
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        private static string FormatValidationErrors(IEnumerable<ValidationErrorInfo> errors) {
            return string.Join(Environment.NewLine + Environment.NewLine,
                errors.Select(error =>
                    $"Description: {error.Description}\n" +
                    $"Id: {error.Id}\n" +
                    $"ErrorType: {error.ErrorType}\n" +
                    $"Part: {error.Part?.Uri}\n" +
                    $"Path: {error.Path?.XPath}"));
        }
    }
}
