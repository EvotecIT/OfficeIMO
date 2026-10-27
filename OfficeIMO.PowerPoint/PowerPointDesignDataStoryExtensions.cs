using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Drawing.Charts;
using OfficeIMO.Drawing;

namespace OfficeIMO.PowerPoint {
    internal static partial class PowerPointDesignExtensions {
        /// <summary>Adds an editable native chart with hero or insight-rail narrative composition.</summary>
        public static PowerPointSlide AddDesignerChartStorySlide(this PowerPointDocument presentation,
            string title, string? subtitle, PowerPointChartStoryContent content,
            PowerPointDesignTheme? theme = null, PowerPointChartStorySlideOptions? options = null) {
            if (presentation == null) throw new ArgumentNullException(nameof(presentation));
            if (string.IsNullOrWhiteSpace(title)) throw new ArgumentException("Title cannot be empty.", nameof(title));
            if (content == null) throw new ArgumentNullException(nameof(content));
            PowerPointDesignTheme resolvedTheme = ResolveTheme(theme);
            PowerPointChartStorySlideOptions resolved = options ?? new PowerPointChartStorySlideOptions();
            PowerPointSlide slide = AddDesignerSlide(presentation, resolved);
            double width = presentation.SlideSize.Width.ToCentimeters();
            double height = presentation.SlideSize.Height.ToCentimeters();
            PrepareLightStorySlide(slide, resolvedTheme, resolved, title, subtitle, width, height);
            PowerPointChartStoryLayoutVariant variant = ResolveChartStoryVariant(resolved, content);
            if (variant == PowerPointChartStoryLayoutVariant.InsightRail) {
                AddChartInsightRail(slide, resolvedTheme, content, width, height);
            } else {
                AddChartHero(slide, resolvedTheme, content, width, height);
            }
            return FinalizeDesignerAccessibility(slide, title);
        }

        /// <summary>Adds an editable side-by-side or decision-matrix comparison.</summary>
        public static PowerPointSlide AddDesignerComparisonSlide(this PowerPointDocument presentation,
            string title, string? subtitle, IEnumerable<PowerPointComparisonItem> items,
            PowerPointDesignTheme? theme = null, PowerPointComparisonSlideOptions? options = null) {
            if (presentation == null) throw new ArgumentNullException(nameof(presentation));
            List<PowerPointComparisonItem> values = (items ?? throw new ArgumentNullException(nameof(items))).ToList();
            if (values.Count < 2 || values.Count > 4)
                throw new ArgumentOutOfRangeException(nameof(items), "Comparison slides support two to four options.");
            PowerPointDesignTheme resolvedTheme = ResolveTheme(theme);
            PowerPointComparisonSlideOptions resolved = options ?? new PowerPointComparisonSlideOptions();
            PowerPointSlide slide = AddDesignerSlide(presentation, resolved);
            double width = presentation.SlideSize.Width.ToCentimeters();
            double height = presentation.SlideSize.Height.ToCentimeters();
            PrepareLightStorySlide(slide, resolvedTheme, resolved, title, subtitle, width, height);
            PowerPointComparisonLayoutVariant variant = ResolveComparisonVariant(resolved, values.Count);
            if (variant == PowerPointComparisonLayoutVariant.DecisionMatrix)
                AddComparisonMatrix(slide, resolvedTheme, values, width, height);
            else
                AddComparisonColumns(slide, resolvedTheme, values, width, height, resolved);
            return FinalizeDesignerAccessibility(slide, title);
        }

        /// <summary>Adds a native appendix table with full-width or notes-rail composition.</summary>
        public static PowerPointSlide AddDesignerAppendixTableSlide(this PowerPointDocument presentation,
            string title, string? subtitle, PowerPointTableData data, PowerPointDesignTheme? theme = null,
            PowerPointAppendixTableSlideOptions? options = null) {
            if (presentation == null) throw new ArgumentNullException(nameof(presentation));
            if (data == null) throw new ArgumentNullException(nameof(data));
            if (data.Rows.Count > 14)
                throw new ArgumentOutOfRangeException(nameof(data),
                    "Appendix-table slides support up to 14 rows; paginate larger tables before rendering.");
            PowerPointDesignTheme resolvedTheme = ResolveTheme(theme);
            PowerPointAppendixTableSlideOptions resolved = options ?? new PowerPointAppendixTableSlideOptions();
            PowerPointSlide slide = AddDesignerSlide(presentation, resolved);
            double width = presentation.SlideSize.Width.ToCentimeters();
            double height = presentation.SlideSize.Height.ToCentimeters();
            PrepareLightStorySlide(slide, resolvedTheme, resolved, title, subtitle, width, height);
            PowerPointAppendixTableLayoutVariant variant = ResolveAppendixVariant(resolved, data);
            if (variant == PowerPointAppendixTableLayoutVariant.NotesRail)
                AddAppendixNotesRail(slide, resolvedTheme, data, width, height);
            else
                AddAppendixFullWidth(slide, resolvedTheme, data, width, height);
            return FinalizeDesignerAccessibility(slide, title);
        }

        internal static PowerPointChartStoryLayoutVariant ResolveChartStoryVariant(
            PowerPointChartStorySlideOptions options, PowerPointChartStoryContent content) {
            if (options.Variant != PowerPointChartStoryLayoutVariant.Auto) return options.Variant;
            if (content.Insights.Count >= 2 || options.DesignIntent.LayoutStrategy == PowerPointAutoLayoutStrategy.Compact)
                return PowerPointChartStoryLayoutVariant.InsightRail;
            return PowerPointChartStoryLayoutVariant.ChartHero;
        }

        internal static PowerPointComparisonLayoutVariant ResolveComparisonVariant(
            PowerPointComparisonSlideOptions options, int itemCount) {
            if (options.Variant != PowerPointComparisonLayoutVariant.Auto) return options.Variant;
            return itemCount > 2 || options.DesignIntent.LayoutStrategy == PowerPointAutoLayoutStrategy.Compact
                ? PowerPointComparisonLayoutVariant.DecisionMatrix
                : PowerPointComparisonLayoutVariant.SideBySide;
        }

        internal static PowerPointAppendixTableLayoutVariant ResolveAppendixVariant(
            PowerPointAppendixTableSlideOptions options, PowerPointTableData data) {
            if (options.Variant != PowerPointAppendixTableLayoutVariant.Auto) return options.Variant;
            return data.Notes.Count > 0
                ? PowerPointAppendixTableLayoutVariant.NotesRail
                : PowerPointAppendixTableLayoutVariant.FullWidth;
        }

        private static void AddChartHero(PowerPointSlide slide, PowerPointDesignTheme theme,
            PowerPointChartStoryContent content, double width, double height) {
            PowerPointLayoutBox chartBounds = new OfficeIMO.PowerPoint.PowerPointLayoutBox(OfficeIMO.OfficeLength.Centimeters(1.5),OfficeIMO.OfficeLength.Centimeters(3.55),OfficeIMO.OfficeLength.Centimeters(width - 3),                OfficeIMO.OfficeLength.Centimeters(height - 5.7));
            PowerPointChart chart = AddStoryChart(slide, content, chartBounds);
            chart.SetLegend(OfficeChartLegendPosition.Bottom);
            AddStoryCaption(slide, theme, content.Caption, content.Provenance, 1.55, height - 1.95, width - 3.1);
        }

        private static void AddChartInsightRail(PowerPointSlide slide, PowerPointDesignTheme theme,
            PowerPointChartStoryContent content, double width, double height) {
            PowerPointLayoutBox body = new OfficeIMO.PowerPoint.PowerPointLayoutBox(OfficeIMO.OfficeLength.Centimeters(1.5),OfficeIMO.OfficeLength.Centimeters(3.55),OfficeIMO.OfficeLength.Centimeters(width - 3),OfficeIMO.OfficeLength.Centimeters(height - 5.25));
            PowerPointLayoutBox[] columns = body.SplitColumns(2,OfficeIMO.OfficeLength.Centimeters(0.7));
            PowerPointLayoutBox chartBounds = new PowerPointLayoutBox(OfficeIMO.OfficeLength.Emus(columns[0].Left.ToEmus()),OfficeIMO.OfficeLength.Emus(columns[0].Top.ToEmus()),                OfficeIMO.OfficeLength.Emus((long)(body.Width.ToEmus() * 0.66)),OfficeIMO.OfficeLength.Emus(columns[0].Height.ToEmus()));
            PowerPointChart chart = AddStoryChart(slide, content, chartBounds);
            chart.SetLegend(OfficeChartLegendPosition.Bottom);
            double railLeft = chartBounds.Right.ToCentimeters() + 0.7;
            double railWidth = body.Right.ToCentimeters() - railLeft;
            PowerPointAutoShape rail = slide.AddRectangleCm(railLeft, body.Top.ToCentimeters(), railWidth, body.Height.ToCentimeters(),
                "Chart Insight Rail");
            rail.FillColor = theme.SurfaceColor;
            rail.OutlineColor = theme.PanelBorderColor;
            double rowHeight = Math.Max(1.25, (body.Height.ToCentimeters() - 1.2) / Math.Max(1, content.Insights.Count));
            for (int index = 0; index < content.Insights.Count; index++) {
                string accent = GetAccent(theme, index);
                slide.AddRectangleCm(railLeft + 0.35, body.Top.ToCentimeters() + 0.45 + index * rowHeight, 0.12,
                    rowHeight - 0.35, "Insight Accent " + (index + 1)).FillColor = accent;
                AddText(slide, content.Insights[index], railLeft + 0.65,
                    body.Top.ToCentimeters() + 0.45 + index * rowHeight, railWidth - 1.0, rowHeight - 0.25, 11,
                    theme.PrimaryTextColor, theme.BodyFontName, bold: index == 0);
            }
            AddStoryCaption(slide, theme, content.Caption, content.Provenance, 1.55, height - 1.62, width - 3.1);
        }

        private static PowerPointChart AddStoryChart(PowerPointSlide slide, PowerPointChartStoryContent content,
            PowerPointLayoutBox bounds) {
            return slide.AddChart(content.ChartKind, content.SharedData, bounds.Left.ToEmus(), bounds.Top.ToEmus(),
                bounds.Width.ToEmus(), bounds.Height.ToEmus(), new PowerPointChartAccessibilityOptions {
                    Name = "Chart Story",
                    AlternativeText = content.AlternativeText ?? "Editable chart story",
                    DataSummary = content.DataSummary,
                    IncludeDataSummaryInAlternativeText = true
                });
        }

        private static void AddComparisonColumns(PowerPointSlide slide, PowerPointDesignTheme theme,
            IReadOnlyList<PowerPointComparisonItem> items, double width, double height,
            PowerPointComparisonSlideOptions options) {
            PowerPointLayoutBox body = new OfficeIMO.PowerPoint.PowerPointLayoutBox(OfficeIMO.OfficeLength.Centimeters(1.5),OfficeIMO.OfficeLength.Centimeters(3.65),OfficeIMO.OfficeLength.Centimeters(width - 3),OfficeIMO.OfficeLength.Centimeters(height - 5.25));
            List<PowerPointCardContent> cards = items.Select(item => new PowerPointCardContent(item.Title,
                BuildComparisonBullets(item))).ToList();
            var cardOptions = new PowerPointCardGridSlideOptions { MaxColumns = items.Count,
                Variant = PowerPointCardGridLayoutVariant.SoftTiles, DesignIntent = options.DesignIntent };
            AddCardGrid(slide, theme, cards, cardOptions, PowerPointCardGridLayoutVariant.SoftTiles, body);
        }

        private static IEnumerable<string> BuildComparisonBullets(PowerPointComparisonItem item) {
            if (!string.IsNullOrWhiteSpace(item.Summary)) yield return item.Summary!;
            foreach (string value in item.Strengths) yield return "+ " + value;
            foreach (string value in item.Tradeoffs) yield return "- " + value;
        }

        private static void AddComparisonMatrix(PowerPointSlide slide, PowerPointDesignTheme theme,
            IReadOnlyList<PowerPointComparisonItem> items, double width, double height) {
            var headers = new List<string> { "Decision lens" };
            headers.AddRange(items.Select(item => item.Title));
            var rows = new List<IReadOnlyList<string>> {
                new[] { "Summary" }.Concat(items.Select(item => item.Summary ?? string.Empty)).ToArray(),
                new[] { "Strengths" }.Concat(items.Select(item => string.Join("; ", item.Strengths))).ToArray(),
                new[] { "Tradeoffs" }.Concat(items.Select(item => string.Join("; ", item.Tradeoffs))).ToArray()
            };
            var data = new PowerPointTableData(headers, rows);
            AddNativeTable(slide, data, new OfficeIMO.PowerPoint.PowerPointLayoutBox(OfficeIMO.OfficeLength.Centimeters(1.5),OfficeIMO.OfficeLength.Centimeters(3.75),OfficeIMO.OfficeLength.Centimeters(width - 3),OfficeIMO.OfficeLength.Centimeters(height - 5.45)),
                theme);
        }

        private static void AddAppendixFullWidth(PowerPointSlide slide, PowerPointDesignTheme theme,
            PowerPointTableData data, double width, double height) {
            AddNativeTable(slide, data, new OfficeIMO.PowerPoint.PowerPointLayoutBox(OfficeIMO.OfficeLength.Centimeters(1.5),OfficeIMO.OfficeLength.Centimeters(3.55),OfficeIMO.OfficeLength.Centimeters(width - 3),OfficeIMO.OfficeLength.Centimeters(height - 5.25)),
                theme);
            AddStoryCaption(slide, theme, data.Caption, data.Provenance, 1.55, height - 1.62, width - 3.1);
        }

        private static void AddAppendixNotesRail(PowerPointSlide slide, PowerPointDesignTheme theme,
            PowerPointTableData data, double width, double height) {
            PowerPointLayoutBox body = new OfficeIMO.PowerPoint.PowerPointLayoutBox(OfficeIMO.OfficeLength.Centimeters(1.5),OfficeIMO.OfficeLength.Centimeters(3.55),OfficeIMO.OfficeLength.Centimeters(width - 3),OfficeIMO.OfficeLength.Centimeters(height - 5.25));
            PowerPointLayoutBox[] columns = body.SplitColumns(2,OfficeIMO.OfficeLength.Centimeters(0.7));
            PowerPointLayoutBox table = new PowerPointLayoutBox(OfficeIMO.OfficeLength.Emus(columns[0].Left.ToEmus()),OfficeIMO.OfficeLength.Emus(columns[0].Top.ToEmus()),                OfficeIMO.OfficeLength.Emus((long)(body.Width.ToEmus() * 0.72)),OfficeIMO.OfficeLength.Emus(columns[0].Height.ToEmus()));
            AddNativeTable(slide, data, table, theme);
            double notesLeft = table.Right.ToCentimeters() + 0.7;
            PowerPointAutoShape rail = slide.AddRectangleCm(notesLeft, body.Top.ToCentimeters(), body.Right.ToCentimeters() - notesLeft,
                body.Height.ToCentimeters(), "Appendix Notes Rail");
            rail.FillColor = theme.SurfaceColor;
            rail.OutlineColor = theme.PanelBorderColor;
            AddText(slide, "Notes", notesLeft + 0.45, body.Top.ToCentimeters() + 0.4, body.Right.ToCentimeters() - notesLeft - 0.9,
                0.55, 14, theme.AccentDarkColor, theme.HeadingFontName, bold: true);
            AddText(slide, string.Join("\n", data.Notes.Select(note => "• " + note)), notesLeft + 0.45,
                body.Top.ToCentimeters() + 1.15, body.Right.ToCentimeters() - notesLeft - 0.9, body.Height.ToCentimeters() - 1.55, 10,
                theme.SecondaryTextColor, theme.BodyFontName);
            AddStoryCaption(slide, theme, data.Caption, data.Provenance, 1.55, height - 1.62, width - 3.1);
        }

        private static PowerPointTable AddNativeTable(PowerPointSlide slide, PowerPointTableData data,
            PowerPointLayoutBox bounds, PowerPointDesignTheme theme) {
            int rowCount = data.Rows.Count + 1;
            PowerPointTable table = slide.AddTable(rowCount, data.Headers.Count, bounds.Left.ToEmus(), bounds.Top.ToEmus(),
                bounds.Width.ToEmus(), bounds.Height.ToEmus());
            table.Name = "Semantic Appendix Table";
            table.HeaderRow = true;
            table.BandedRows = true;
            for (int column = 0; column < data.Headers.Count; column++) {
                PowerPointTableCell cell = table.GetCell(0, column);
                cell.Text = data.Headers[column];
                cell.FillColor = theme.AccentDarkColor;
                cell.Bold = true;
                PowerPointTextRun? run = cell.Runs.FirstOrDefault();
                if (run != null) {
                    run.Color = theme.AccentContrastColor;
                    run.FontName = theme.BodyFontName;
                }
            }
            for (int row = 0; row < data.Rows.Count; row++) {
                for (int column = 0; column < data.Headers.Count; column++) {
                    PowerPointTableCell cell = table.GetCell(row + 1, column);
                    cell.Text = data.Rows[row][column];
                    if (row % 2 == 1) cell.FillColor = theme.SurfaceColor;
                }
            }
            double rowHeight = bounds.Height.ToPoints() / Math.Max(1, rowCount);
            foreach (PowerPointTableRow row in table.RowItems) row.HeightPoints = rowHeight;
            return table;
        }

        private static void AddStoryCaption(PowerPointSlide slide, PowerPointDesignTheme theme, string? caption,
            string? provenance, double left, double top, double width) {
            if (!string.IsNullOrWhiteSpace(caption)) {
                AddText(slide, caption!, left, top, width * 0.65, 0.42, 9,
                    theme.SecondaryTextColor, theme.BodyFontName, bold: true);
            }
            if (!string.IsNullOrWhiteSpace(provenance)) {
                PowerPointTextBox source = AddText(slide, "Source: " + provenance, left + width * 0.65, top,
                    width * 0.35, 0.42, 8, theme.MutedTextColor, theme.BodyFontName);
                RightAlignText(source);
            }
        }
    }
}
