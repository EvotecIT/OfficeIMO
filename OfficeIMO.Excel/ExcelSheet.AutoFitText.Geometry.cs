using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel;

public partial class ExcelSheet {
    /// <summary>Projects measured text boxes using OOXML's upward/downward angle encoding.</summary>
    private static (double Width, double Height) ProjectAutoFitTextBox(
        string text, double width, double height, uint rotation, ExcelTextMeasurer measurer, ExcelTextMeasurer.Style style) {
        if (rotation == 255U) {
            double stackedWidth = 0;
            int elements = 0;
            foreach (string element in OfficeTextElements.Enumerate(text)) {
                if (element == "\r" || element == "\n" || element == "\r\n") continue;
                stackedWidth = Math.Max(stackedWidth, measurer.MeasureWidthOrDefault(element, style, 0));
                elements++;
            }
            return (stackedWidth, height * Math.Max(1, elements));
        }
        double degrees = rotation <= 90U ? rotation : rotation <= 180U ? 90D - rotation : 0D;
        double radians = Math.Abs(degrees) * Math.PI / 180D;
        double cosine = Math.Cos(radians);
        double sine = Math.Sin(radians);
        return (width * cosine + height * sine, width * sine + height * cosine);
    }

    private static double GetAutoFitIndentPixels(Alignment? alignment, ExcelTextMeasurer measurer, ExcelTextMeasurer.Style style) =>
        (alignment?.Indent?.Value ?? 0U) * 3D * measurer.MeasureWidthOrDefault(" ", style, 0);
}
