using OfficeIMO.Drawing;
using OfficeIMO.Web.Converter.Models;

namespace OfficeIMO.Web.Converter.Services;

public static class ConversionRouteCatalog {
    public static IReadOnlyList<ConversionRoute> All { get; } =
        OfficeConversionCapabilityCatalog.BrowserRoutes.Select(CreateBrowserRoute).ToArray();

    public static ConversionRoute Default => All[0];

    public static ConversionRoute Find(string? id) =>
        All.FirstOrDefault(route => string.Equals(route.Id, id, StringComparison.OrdinalIgnoreCase)) ?? Default;

    private static ConversionRoute CreateBrowserRoute(OfficeConversionCapability route) =>
        new(
            route.Id,
            route.Source == "Markdown" ? "MD" : route.Source,
            route.Target == "Markdown" ? "MD" : route.Target,
            GetTitle(route),
            route.Description,
            route.InputKind == OfficeConversionInputKind.File ? ConversionInputKind.File : ConversionInputKind.Text,
            string.Join(",", route.SourceExtensions),
            route.Api,
            GetAccentClass(route));

    private static string GetTitle(OfficeConversionCapability route) => route.Id switch {
        "docx-pdf" => "Word to PDF",
        "xlsx-pdf" => "Excel to PDF",
        "pptx-pdf" => "PowerPoint to PDF",
        "pdf-docx" => "PDF to Word",
        "pdf-xlsx" => "PDF tables to Excel",
        "pdf-pptx" => "PDF to PowerPoint",
        "markdown-docx" => "Markdown to Word",
        _ => route.Source + " to " + route.Target
    };

    private static string GetAccentClass(OfficeConversionCapability route) => route.Source switch {
        "DOCX" => "ocx-route-card--word",
        "XLSX" => "ocx-route-card--excel",
        "PPTX" => "ocx-route-card--powerpoint",
        "PDF" => "ocx-route-card--pdf",
        "Markdown" => "ocx-route-card--markdown",
        _ => "ocx-route-card--html"
    };
}
