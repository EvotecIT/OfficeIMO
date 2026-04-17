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
using C = DocumentFormat.OpenXml.Drawing.Charts;
using PptImagePartType = OfficeIMO.PowerPoint.ImagePartType;

namespace OfficeIMO.Tests {
    public class PowerPointFunctionalSmokeTests {
        private static readonly byte[] OnePixelPng = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMB/6X4nKkAAAAASUVORK5CYII=");

        [Fact]
        public void CanBuildRichDeckAndValidate() {
            string filePath = CreateTempFilePath(".pptx");
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

        [Fact]
        public void CanBuildModernThemeDeckAndValidate() {
            string filePath = CreateTempFilePath(".pptx");
            string backgroundPath = CreateTempFilePath(".png");

            try {
                File.WriteAllBytes(backgroundPath, OnePixelPng);

                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    presentation.SlideSize.SetPreset(PowerPointSlideSizePreset.Screen16x9);
                    presentation.ThemeName = "OfficeIMO Modern Smoke";
                    presentation.SetThemeColorsForAllMasters(new Dictionary<PowerPointThemeColor, string> {
                        [PowerPointThemeColor.Dark1] = "161411",
                        [PowerPointThemeColor.Light1] = "F8F5EF",
                        [PowerPointThemeColor.Dark2] = "253746",
                        [PowerPointThemeColor.Light2] = "EFE8DA",
                        [PowerPointThemeColor.Accent1] = "156082",
                        [PowerPointThemeColor.Accent2] = "F26A3D",
                        [PowerPointThemeColor.Accent3] = "8CB369",
                        [PowerPointThemeColor.Accent4] = "6B6EA8",
                        [PowerPointThemeColor.Accent5] = "D6A84F",
                        [PowerPointThemeColor.Accent6] = "6C8EAD"
                    });
                    presentation.SetThemeFontsForAllMasters(new PowerPointThemeFontSet(
                        majorLatin: "Aptos Display",
                        minorLatin: "Aptos",
                        majorEastAsian: "Yu Gothic",
                        minorEastAsian: "Yu Gothic",
                        majorComplexScript: "Arial",
                        minorComplexScript: "Arial"));

                    PowerPointSlide cover = presentation.AddSlide(SlideLayoutValues.TitleOnly);
                    cover.BackgroundColor = "F8F5EF";
                    cover.Transition = SlideTransition.Morph;
                    PowerPointTextBox title = cover.AddTitle("Commercial Snapshot",
                        new PowerPointLayoutBox(PowerPointUnits.Cm(1.4), PowerPointUnits.Cm(1.0),
                            PowerPointUnits.Cm(18.0), PowerPointUnits.Cm(1.2)));
                    title.FontSize = 32;
                    title.Color = "161411";

                    PowerPointAutoShape hero = cover.AddRectangleCm(1.4, 3.0, 16.0, 3.0, "Hero Card");
                    hero.FillColor = "156082";
                    hero.FillTransparency = 8;
                    hero.OutlineColor = "156082";
                    hero.SetShadow("000000", blurPoints: 10, distancePoints: 4, angleDegrees: 90, transparencyPercent: 70);

                    PowerPointTextBox insight = cover.AddTextBox("A clean, generated deck with theme colors, font scheme, background, effects, chart, and table.",
                        new PowerPointLayoutBox(PowerPointUnits.Cm(2.0), PowerPointUnits.Cm(3.6),
                            PowerPointUnits.Cm(14.8), PowerPointUnits.Cm(1.3)));
                    insight.FontSize = 18;
                    insight.Color = "FFFFFF";

                    PowerPointAutoShape accent = cover.AddEllipseCm(17.8, 0.7, 2.5, 2.5, "Glow Accent");
                    accent.FillColor = "F26A3D";
                    accent.FillTransparency = 18;
                    accent.OutlineColor = "F26A3D";
                    accent.SetGlow("F26A3D", radiusPoints: 8, transparencyPercent: 30);

                    PowerPointSlide dashboard = presentation.AddSlide(SlideLayoutValues.Blank);
                    dashboard.SetBackgroundImage(backgroundPath);
                    dashboard.Transition = SlideTransition.Fade;

                    PowerPointAutoShape panel = dashboard.AddRectangleCm(0.9, 0.7, 24.0, 12.0, "Dashboard Surface");
                    panel.FillColor = "F8F5EF";
                    panel.FillTransparency = 3;
                    panel.OutlineColor = "EFE8DA";
                    panel.SetSoftEdges(1.5);

                    PowerPointChartData data = new(
                        new[] { "Jan", "Feb", "Mar", "Apr" },
                        new[] {
                            new PowerPointChartSeries("Revenue", new[] { 12d, 18d, 21d, 28d }),
                            new PowerPointChartSeries("Profit", new[] { 4d, 7d, 9d, 13d })
                        });
                    PowerPointChart chart = dashboard.AddLineChartCm(data, 1.4, 1.3, 14.5, 7.0);
                    chart.SetTitle("Momentum Over Time");
                    chart.SetLegend(C.LegendPositionValues.Bottom);
                    chart.SetChartAreaStyle(fillColor: "FFFFFF", lineColor: "EFE8DA");
                    chart.SetPlotAreaStyle(fillColor: "FFFFFF", lineColor: "FFFFFF");
                    chart.SetSeriesLineColor("Revenue", "156082", widthPoints: 2.5);
                    chart.SetSeriesLineColor("Profit", "F26A3D", widthPoints: 2.5);
                    chart.SetValueAxisGridlines(showMajor: true, lineColor: "D8D5CC", lineWidthPoints: 0.75);

                    PowerPointTable table = dashboard.AddTable(rows: 3, columns: 2,
                        styleName: presentation.TableStyles.First().Name,
                        left: PowerPointUnits.Cm(16.8), top: PowerPointUnits.Cm(1.5),
                        width: PowerPointUnits.Cm(7.0), height: PowerPointUnits.Cm(3.8),
                        firstRow: true, bandedRows: true);
                    table.GetCell(0, 0).Text = "Metric";
                    table.GetCell(0, 1).Text = "Value";
                    table.GetCell(1, 0).Text = "Revenue";
                    table.GetCell(1, 1).Text = "28";
                    table.GetCell(2, 0).Text = "Profit";
                    table.GetCell(2, 1).Text = "13";

                    PowerPointAutoShape callout = dashboard.AddRectangleCm(16.8, 6.1, 7.0, 3.1, "Observation Card");
                    callout.FillColor = "EFE8DA";
                    callout.FillTransparency = 0;
                    callout.OutlineColor = "6B6EA8";
                    callout.OutlineWidthPoints = 1.2;
                    callout.SetReflection(blurPoints: 2, distancePoints: 1, startOpacityPercent: 20, endOpacityPercent: 0);

                    PowerPointTextBox calloutText = dashboard.AddTextBox("Observation\nThe revenue trend stays healthy through Q2.",
                        new PowerPointLayoutBox(PowerPointUnits.Cm(17.3), PowerPointUnits.Cm(6.5),
                            PowerPointUnits.Cm(6.0), PowerPointUnits.Cm(2.2)));
                    calloutText.FontSize = 16;
                    calloutText.Color = "161411";

                    dashboard.Notes.Text = "Modern deck smoke test validates theme, background, effects, chart and table.";
                    presentation.Save();
                    var packageErrors = presentation.ValidateDocument().ToList();
                    Assert.True(packageErrors.Count == 0, FormatValidationErrors(packageErrors));
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    var validator = new OpenXmlValidator(FileFormatVersions.Microsoft365);
                    var errors = validator.Validate(document).ToList();
                    Assert.True(errors.Count == 0, FormatValidationErrors(errors));

                    PresentationPart part = document.PresentationPart!;
                    Assert.Contains(part.SlideParts, slidePart => slidePart.ImageParts.Any());
                    Assert.Contains(part.SlideParts, slidePart => slidePart.ChartParts.Any());
                    Assert.Contains(part.SlideMasterParts, master =>
                        master.ThemePart?.Theme?.Name?.Value == "OfficeIMO Modern Smoke");
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
                if (File.Exists(backgroundPath)) {
                    File.Delete(backgroundPath);
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

        private static string CreateTempFilePath(string extension) {
            string path = Path.GetTempFileName();
            File.Delete(path);
            return Path.ChangeExtension(path, extension);
        }
    }
}
