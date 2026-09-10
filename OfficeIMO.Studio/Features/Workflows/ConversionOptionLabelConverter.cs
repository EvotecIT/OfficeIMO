using System.Globalization;
using Avalonia.Data.Converters;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Html.Pdf;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Studio.Infrastructure.Localization;
using OfficeIMO.Word.Pdf;

namespace OfficeIMO.Studio.Features.Workflows;

/// <summary>Localized labels for the canonical conversion modes exposed by the workbench.</summary>
public sealed class ConversionOptionLabelConverter : IValueConverter {
    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture) {
        string label = value switch {
            PdfWordImportMode.EditableContent or PdfPowerPointImportMode.EditableContent => "Editable content",
            PdfWordImportMode.VisualPages or PdfPowerPointImportMode.VisualPages => "Visual pages",
            PdfPowerPointImportMode.HybridVisualAndEditableTables => "Visual pages with editable tables",
            PdfPowerPointImportMode.EditableTables => "Editable tables only",
            ExcelPdfWorksheetLayoutMode.WorksheetCanvas => "Worksheet canvas",
            ExcelPdfWorksheetLayoutMode.FlowTable => "Flowing tables",
            PdfHtmlProfile.PositionedReview => "Positioned pages",
            PdfHtmlProfile.Semantic => "Semantic document",
            _ => string.Empty
        };
        return StudioLocalization.Current.GetOrDefault("Conversion.Mode." + value, label);
    }
    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture) => throw new NotSupportedException();
}
