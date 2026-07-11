using System;
using System.Linq;

namespace OfficeIMO.PowerPoint {
    internal static partial class PowerPointDesignExtensions {
        internal static PowerPointSlide FinalizeDesignerAccessibility(PowerPointSlide slide, string? requestedTitle = null,
            string language = "en-US") {
            if (slide == null) throw new ArgumentNullException(nameof(slide));
            PowerPointTextBox? detectedTitle = slide.TextBoxes
                .Where(box => !string.IsNullOrWhiteSpace(box.Text))
                .OrderBy(box => box.TopPoints)
                .ThenByDescending(box => box.FontSize ?? 0)
                .FirstOrDefault();
            foreach (PowerPointShape shape in slide.Shapes) {
                if (HasDesignerDecorativeName(shape.Name) &&
                    shape.ShapeContentType != PowerPointShapeContentType.TextBox) {
                    shape.Decorative = true;
                    shape.Description = null;
                    continue;
                }

                if (shape is PowerPointTextBox textBox && !string.IsNullOrWhiteSpace(textBox.Text)) {
                    textBox.SetLanguage(language);
                    textBox.Title = ReferenceEquals(textBox, detectedTitle)
                        ? "Slide title"
                        : CreateAccessibleTitle(textBox.Text, "Text");
                } else if (shape is PowerPointTable table) {
                    table.SetLanguage(language);
                    table.HeaderRow = true;
                    table.Title ??= "Data table";
                    table.Description ??= "Data table with " + table.Rows + " row(s) and " + table.Columns + " column(s).";
                } else if (shape is PowerPointChart chart) {
                    chart.Title ??= "Data chart";
                    if (string.IsNullOrWhiteSpace(chart.Description) && chart.TryGetOfficeSnapshot(out _)) {
                        chart.Description = chart.CreateDataSummary();
                    }
                } else if (shape is PowerPointSmartArt smartArt) {
                    smartArt.Title ??= "Diagram";
                    smartArt.Description ??= "Diagram with " + smartArt.NodeCount + " editable node(s).";
                } else if (shape.ShapeContentType == PowerPointShapeContentType.Picture ||
                           shape.ShapeContentType == PowerPointShapeContentType.Media) {
                    if (string.IsNullOrWhiteSpace(shape.Description)) {
                        shape.Decorative = true;
                    } else {
                        shape.Title ??= CreateAccessibleTitle(shape.Description, "Visual");
                    }
                } else if (!string.IsNullOrWhiteSpace(shape.Name)) {
                    shape.Title ??= shape.Name;
                }
            }

            if (detectedTitle != null && !string.IsNullOrWhiteSpace(requestedTitle)) {
                detectedTitle.Title = "Slide title";
            }
            return slide;
        }

        private static bool HasDesignerDecorativeName(string? value) {
            if (string.IsNullOrWhiteSpace(value)) return false;
            string name = value!.ToLowerInvariant();
            string[] hints = { "accent", "background", "plane", "diagonal", "halo", "rail", "connector",
                "rule", "motif", "dot", "wash", "plate", "underline", "marker", "divider" };
            return hints.Any(name.Contains);
        }

        private static string CreateAccessibleTitle(string? value, string fallback) {
            string normalized = string.Join(" ", (value ?? string.Empty)
                .Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries));
            if (normalized.Length == 0) return fallback;
            return normalized.Length <= 80 ? normalized : normalized.Substring(0, 77) + "...";
        }
    }
}
