using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeIMO.PowerPoint;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Examples.PowerPoint {
    /// <summary>
    /// Demonstrates a modern PowerPoint deck with theme colors, theme fonts, backgrounds, transitions,
    /// effects, charts, a table, notes, and validation.
    /// </summary>
    public static class ModernPowerPointDeck {
        private const string Ink = "161411";
        private const string Paper = "F8F5EF";
        private const string Linen = "EFE8DA";
        private const string Teal = "156082";
        private const string Orange = "F26A3D";
        private const string Sage = "8CB369";
        private const string Indigo = "6B6EA8";
        private const string Gold = "D6A84F";

        public static void Example_ModernPowerPointDeck(string folderPath, bool openPowerPoint) {
            Console.WriteLine("[*] PowerPoint - Modern themed deck");
            string filePath = Path.Combine(folderPath, "Modern PowerPoint Deck.pptx");
            string backgroundImagePath = Path.Combine(AppContext.BaseDirectory, "Images", "BackgroundImage.png");

            using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
            presentation.SlideSize.SetPreset(PowerPointSlideSizePreset.Screen16x9);
            presentation.ThemeName = "OfficeIMO Modern";
            presentation.SetThemeColorsForAllMasters(new Dictionary<PowerPointThemeColor, string> {
                [PowerPointThemeColor.Dark1] = Ink,
                [PowerPointThemeColor.Light1] = Paper,
                [PowerPointThemeColor.Dark2] = "253746",
                [PowerPointThemeColor.Light2] = Linen,
                [PowerPointThemeColor.Accent1] = Teal,
                [PowerPointThemeColor.Accent2] = Orange,
                [PowerPointThemeColor.Accent3] = Sage,
                [PowerPointThemeColor.Accent4] = Indigo,
                [PowerPointThemeColor.Accent5] = Gold,
                [PowerPointThemeColor.Accent6] = "6C8EAD"
            });
            presentation.SetThemeFontsForAllMasters(new PowerPointThemeFontSet(
                majorLatin: "Aptos Display",
                minorLatin: "Aptos",
                majorEastAsian: "Yu Gothic",
                minorEastAsian: "Yu Gothic",
                majorComplexScript: "Arial",
                minorComplexScript: "Arial"));

            AddCoverSlide(presentation, backgroundImagePath);
            AddPerformanceSlide(presentation);
            AddChannelMixSlide(presentation);
            AddGrowthSlide(presentation);

            presentation.Save();
            List<DocumentFormat.OpenXml.Validation.ValidationErrorInfo> errors = presentation.ValidateDocument();
            if (errors.Count > 0) {
                string details = string.Join(Environment.NewLine, errors.Take(5).Select(error => error.Description));
                throw new InvalidOperationException($"PowerPoint validation failed with {errors.Count} error(s).{Environment.NewLine}{details}");
            }

            Console.WriteLine($"    Saved: {filePath}");
            Console.WriteLine("    Validation: no Open XML errors found.");
            Helpers.Open(filePath, openPowerPoint);
        }

        private static void AddCoverSlide(PowerPointPresentation presentation, string backgroundImagePath) {
            PowerPointSlide slide = presentation.AddSlide();
            slide.BackgroundColor = Paper;
            slide.Transition = SlideTransition.Morph;
            ApplyBackgroundImage(slide, backgroundImagePath);
            AddWash(slide);

            PowerPointAutoShape rail = slide.AddRectangleCm(1.0, 1.1, 0.18, 10.8, "Accent Rail");
            rail.FillColor = Teal;
            rail.OutlineColor = Teal;

            PowerPointTextBox eyebrow = slide.AddTextBoxCm("Commercial snapshot", 1.5, 1.1, 9.0, 0.6);
            eyebrow.FontSize = 11;
            eyebrow.Color = Teal;

            PowerPointTextBox title = slide.AddTextBoxCm("Modern decks with OfficeIMO.PowerPoint", 1.5, 2.1, 17.5, 2.5);
            title.FontSize = 34;
            title.Color = Ink;
            title.TextAutoFit = PowerPointTextAutoFit.Normal;
            title.SetTextMarginsCm(0, 0, 0, 0);

            PowerPointAutoShape card = slide.AddRectangleCm(1.5, 5.6, 15.8, 3.5, "Hero Story Card");
            card.FillColor = Teal;
            card.FillTransparency = 8;
            card.OutlineColor = Teal;
            card.SetShadow("000000", blurPoints: 10, distancePoints: 4, angleDegrees: 90, transparencyPercent: 72);

            PowerPointTextBox story = slide.AddTextBoxCm(
                "Themes, backgrounds, effects, editable charts, tables, notes, and transitions in one generated presentation.",
                2.1, 6.35, 14.4, 1.5);
            story.FontSize = 19;
            story.Color = "FFFFFF";
            story.TextAutoFit = PowerPointTextAutoFit.Normal;

            PowerPointAutoShape glow = slide.AddEllipseCm(24.4, 1.0, 4.5, 4.5, "Glow Accent");
            glow.FillColor = Orange;
            glow.FillTransparency = 18;
            glow.OutlineColor = Orange;
            glow.SetGlow(Orange, radiusPoints: 8, transparencyPercent: 28);

            PowerPointAutoShape pill = slide.AddRectangleCm(22.2, 13.2, 8.8, 1.1, "Footer Pill");
            pill.FillColor = Linen;
            pill.FillTransparency = 0;
            pill.OutlineColor = Linen;
            pill.SetSoftEdges(1.2);

            PowerPointTextBox footer = slide.AddTextBoxCm("Validated PPTX output", 22.7, 13.45, 7.8, 0.45);
            footer.FontSize = 11;
            footer.Color = Ink;
            footer.TextVerticalAlignment = A.TextAnchoringTypeValues.Center;
            slide.Notes.Text = "Opening slide introduces the feature set demonstrated by this deck.";
        }

        private static void AddPerformanceSlide(PowerPointPresentation presentation) {
            PowerPointSlide slide = presentation.AddSlide();
            slide.BackgroundColor = Paper;
            slide.Transition = SlideTransition.Fade;
            AddSlideTitle(slide, "Monthly performance", "Column chart, KPI table, shape effects, and speaker notes.");

            PowerPointAutoShape chartPanel = AddPanel(slide, 1.3, 3.1, 16.2, 9.9, "Chart Surface");
            chartPanel.SetShadow("000000", blurPoints: 7, distancePoints: 2, angleDegrees: 90, transparencyPercent: 84);

            PowerPointChartData data = new(
                new[] { "Jan", "Feb", "Mar", "Apr", "May", "Jun" },
                new[] {
                    new PowerPointChartSeries("Sales", new[] { 10d, 12d, 16d, 19d, 24d, 27d }),
                    new PowerPointChartSeries("Profit", new[] { 4d, 5d, 7d, 9d, 11d, 13d })
                });
            PowerPointChart chart = slide.AddChartCm(data, 1.8, 3.7, 15.1, 8.5);
            chart.SetTitle("Sales vs Profit")
                .SetLegend(C.LegendPositionValues.Bottom)
                .SetChartAreaStyle(fillColor: "FFFFFF", lineColor: "FFFFFF")
                .SetPlotAreaStyle(fillColor: "FFFFFF", lineColor: "FFFFFF")
                .SetSeriesFillColor("Sales", Teal)
                .SetSeriesFillColor("Profit", Orange)
                .SetValueAxisGridlines(showMajor: true, lineColor: "D8D5CC", lineWidthPoints: 0.75)
                .SetValueAxisLabelTextStyle(fontSizePoints: 9, color: "534F49", fontName: "Aptos")
                .SetCategoryAxisLabelTextStyle(fontSizePoints: 9, color: "534F49", fontName: "Aptos");

            PowerPointAutoShape takeawayCard = AddPanel(slide, 18.5, 3.1, 11.4, 3.6, "Takeaway Card");
            takeawayCard.FillColor = Linen;
            takeawayCard.OutlineColor = Orange;
            takeawayCard.OutlineWidthPoints = 1.4;
            takeawayCard.SetReflection(blurPoints: 2, distancePoints: 1, startOpacityPercent: 15, endOpacityPercent: 0);
            AddLabel(slide, "Takeaway", 19.1, 3.55, 5.0, 0.55, Orange, 12, bold: true);
            AddBody(slide, "Sales nearly triple from January to June while profit expands at a steadier pace.", 19.1, 4.45, 9.8, 1.5, 18);

            PowerPointTable table = slide.AddTableCm(
                new[] {
                    new KpiRow { Metric = "Sales", Value = "27", Status = "up 170%" },
                    new KpiRow { Metric = "Profit", Value = "13", Status = "up 225%" },
                    new KpiRow { Metric = "Runway", Value = "Q3", Status = "healthy" }
                },
                options => {
                    options.HeaderCase = HeaderCase.Title;
                    options.PinFirst("Metric");
                },
                includeHeaders: true,
                leftCm: 18.5,
                topCm: 7.6,
                widthCm: 11.4,
                heightCm: 4.5);
            table.ApplyStyle(PowerPointTableStylePreset.Default);
            table.SetColumnWidthsPoints(105, 75, 125);
            StyleTable(table);

            slide.Notes.Text = "Performance slide: keep the conversation on the widening gap between revenue growth and profitability.";
        }

        private static void AddChannelMixSlide(PowerPointPresentation presentation) {
            PowerPointSlide slide = presentation.AddSlide();
            slide.BackgroundColor = Paper;
            slide.Transition = SlideTransition.Fade;
            AddSlideTitle(slide, "Channel mix", "Doughnut chart, modern cards, and soft visual hierarchy.");

            PowerPointAutoShape chartPanel = AddPanel(slide, 1.3, 3.0, 15.3, 10.0, "Channel Chart Panel");
            chartPanel.FillColor = "FFFFFF";
            chartPanel.SetShadow("000000", blurPoints: 7, distancePoints: 2, angleDegrees: 90, transparencyPercent: 84);

            PowerPointChartData data = new(
                new[] { "Direct", "Partner", "Online", "Events" },
                new[] { new PowerPointChartSeries("Share", new[] { 42d, 27d, 22d, 9d }) });
            PowerPointChart chart = slide.AddDoughnutChartCm(data, 2.2, 3.9, 13.4, 7.8);
            chart.SetTitle("Revenue Share by Channel")
                .SetLegend(C.LegendPositionValues.Bottom)
                .SetDataLabels(showPercent: true)
                .SetDataLabelPosition(C.DataLabelPositionValues.BestFit)
                .SetChartAreaStyle(fillColor: "FFFFFF", lineColor: "FFFFFF")
                .SetSeriesFillColor(0, Teal);

            AddMetricCard(slide, "42%", "Direct", "Primary engine", 18.2, 3.1, Teal);
            AddMetricCard(slide, "27%", "Partner", "Strong co-sell base", 18.2, 6.1, Orange);
            AddMetricCard(slide, "22%", "Online", "Efficient pipeline", 18.2, 9.1, Sage);

            PowerPointAutoShape noteCard = AddPanel(slide, 25.2, 3.1, 5.1, 8.9, "Story Card");
            noteCard.FillColor = Linen;
            noteCard.OutlineColor = Indigo;
            noteCard.OutlineWidthPoints = 1.2;
            AddLabel(slide, "Story", 25.75, 3.75, 3.8, 0.5, Indigo, 12, bold: true);
            AddBody(slide, "Enterprise remains anchor-led, but online demand now deserves a dedicated program.", 25.75, 4.75, 4.1, 4.0, 16);

            slide.Notes.Text = "Channel slide: direct is still dominant, but online has enough share to justify deliberate investment.";
        }

        private static void AddGrowthSlide(PowerPointPresentation presentation) {
            PowerPointSlide slide = presentation.AddSlide();
            slide.BackgroundColor = Paper;
            slide.Transition = SlideTransition.Fade;
            AddSlideTitle(slide, "Growth trend", "Line chart with styled series, markers, gridlines, and an observation card.");

            PowerPointAutoShape stripe = slide.AddRectangleCm(1.3, 4.2, 0.36, 8.8, "Accent Stripe");
            stripe.FillColor = Orange;
            stripe.OutlineColor = Orange;

            PowerPointAutoShape chartPanel = AddPanel(slide, 2.2, 3.2, 20.7, 9.6, "Trend Chart Panel");
            chartPanel.FillColor = "FFFFFF";
            chartPanel.SetShadow("000000", blurPoints: 7, distancePoints: 2, angleDegrees: 90, transparencyPercent: 84);

            PowerPointChartData data = new(
                new[] { "Jan", "Feb", "Mar", "Apr", "May", "Jun" },
                new[] {
                    new PowerPointChartSeries("Sales", new[] { 10d, 14d, 18d, 21d, 25d, 28d }),
                    new PowerPointChartSeries("Profit", new[] { 4d, 6d, 8d, 10d, 12d, 14d })
                });
            PowerPointChart chart = slide.AddLineChartCm(data, 2.8, 3.8, 19.5, 8.4);
            chart.SetTitle("Momentum Over Time")
                .SetLegend(C.LegendPositionValues.Bottom)
                .SetChartAreaStyle(fillColor: "FFFFFF", lineColor: "FFFFFF")
                .SetPlotAreaStyle(fillColor: "FFFFFF", lineColor: "FFFFFF")
                .SetSeriesLineColor("Sales", Teal, widthPoints: 2.75)
                .SetSeriesLineColor("Profit", Orange, widthPoints: 2.75)
                .SetSeriesMarker("Sales", C.MarkerStyleValues.Diamond, size: 8, fillColor: Teal, lineColor: Teal)
                .SetSeriesMarker("Profit", C.MarkerStyleValues.Square, size: 8, fillColor: Orange, lineColor: Orange)
                .SetValueAxisGridlines(showMajor: true, lineColor: "D8D5CC", lineWidthPoints: 0.75)
                .SetValueAxisLabelTextStyle(fontSizePoints: 9, color: "534F49", fontName: "Aptos")
                .SetCategoryAxisLabelTextStyle(fontSizePoints: 9, color: "534F49", fontName: "Aptos");

            PowerPointAutoShape callout = AddPanel(slide, 24.2, 4.8, 7.4, 5.9, "Observation Card");
            callout.FillColor = Linen;
            callout.OutlineColor = Indigo;
            callout.OutlineWidthPoints = 1.2;
            callout.SetReflection(blurPoints: 2, distancePoints: 1, startOpacityPercent: 18, endOpacityPercent: 0);
            AddLabel(slide, "Observation", 25.0, 5.45, 5.5, 0.55, Indigo, 13, bold: true);
            AddBody(slide, "The slope stays healthy through Q2, which makes the line view easier to read than a scatter plot here.", 25.0, 6.6, 5.6, 3.0, 16);

            slide.Notes.Text = "Growth slide: this intentionally uses a line chart because the month-to-month sequence matters.";
        }

        private static void AddSlideTitle(PowerPointSlide slide, string title, string subtitle) {
            PowerPointTextBox titleBox = slide.AddTextBoxCm(title, 1.3, 0.85, 17.5, 0.95);
            titleBox.FontSize = 23;
            titleBox.Color = Ink;
            titleBox.SetTextMarginsCm(0, 0, 0, 0);

            PowerPointTextBox subtitleBox = slide.AddTextBoxCm(subtitle, 1.35, 1.85, 16.5, 0.5);
            subtitleBox.FontSize = 10;
            subtitleBox.Color = "6A625B";
            subtitleBox.SetTextMarginsCm(0, 0, 0, 0);
        }

        private static PowerPointAutoShape AddPanel(PowerPointSlide slide, double leftCm, double topCm, double widthCm, double heightCm, string name) {
            PowerPointAutoShape panel = slide.AddRectangleCm(leftCm, topCm, widthCm, heightCm, name);
            panel.FillColor = "FFFFFF";
            panel.FillTransparency = 0;
            panel.OutlineColor = Linen;
            panel.SetSoftEdges(0.9);
            return panel;
        }

        private static void AddMetricCard(PowerPointSlide slide, string value, string label, string detail, double leftCm, double topCm, string color) {
            PowerPointAutoShape card = AddPanel(slide, leftCm, topCm, 5.7, 2.25, label + " Card");
            card.OutlineColor = color;
            card.OutlineWidthPoints = 1.1;

            PowerPointAutoShape chip = slide.AddEllipseCm(leftCm + 0.45, topCm + 0.42, 0.55, 0.55, label + " Chip");
            chip.FillColor = color;
            chip.OutlineColor = color;

            AddLabel(slide, value, leftCm + 1.25, topCm + 0.25, 3.8, 0.65, Ink, 18, bold: true);
            AddLabel(slide, label, leftCm + 1.25, topCm + 0.95, 3.8, 0.45, color, 11, bold: true);
            AddBody(slide, detail, leftCm + 1.25, topCm + 1.45, 3.95, 0.45, 10);
        }

        private static PowerPointTextBox AddLabel(PowerPointSlide slide, string text, double leftCm, double topCm, double widthCm, double heightCm, string color, int fontSize, bool bold) {
            PowerPointTextBox box = slide.AddTextBoxCm(text, leftCm, topCm, widthCm, heightCm);
            box.FontSize = fontSize;
            box.Color = color;
            box.Bold = bold;
            box.SetTextMarginsCm(0, 0, 0, 0);
            return box;
        }

        private static PowerPointTextBox AddBody(PowerPointSlide slide, string text, double leftCm, double topCm, double widthCm, double heightCm, int fontSize) {
            PowerPointTextBox box = slide.AddTextBoxCm(text, leftCm, topCm, widthCm, heightCm);
            box.FontSize = fontSize;
            box.Color = Ink;
            box.TextAutoFit = PowerPointTextAutoFit.Normal;
            box.SetTextMarginsCm(0, 0, 0, 0);
            return box;
        }

        private static void StyleTable(PowerPointTable table) {
            for (int c = 0; c < table.Columns; c++) {
                PowerPointTableCell header = table.GetCell(0, c);
                header.Bold = true;
                header.HorizontalAlignment = A.TextAlignmentTypeValues.Center;
                header.VerticalAlignment = A.TextAnchoringTypeValues.Center;
            }
        }

        private static void ApplyBackgroundImage(PowerPointSlide slide, string imagePath) {
            if (File.Exists(imagePath)) {
                slide.SetBackgroundImage(imagePath);
            }
        }

        private static void AddWash(PowerPointSlide slide) {
            PowerPointAutoShape wash = slide.AddRectangleCm(0, 0, 33.87, 19.05, "Background Wash");
            wash.FillColor = Paper;
            wash.FillTransparency = 6;
            wash.OutlineColor = Paper;
        }

        private sealed class KpiRow {
            public string Metric { get; set; } = string.Empty;
            public string Value { get; set; } = string.Empty;
            public string Status { get; set; } = string.Empty;
        }
    }
}
