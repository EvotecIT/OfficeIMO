namespace OfficeIMO;

public static partial class OfficeConversionCapabilityCatalog {
    internal static string? PreferredRouteId(string sourceExtension, string targetExtension) {
        // Excel 2003 XML stays the default for .xml to .xlsx; data XML is chosen with RouteId "xml-xlsx".
        if (sourceExtension.Equals(".xml", StringComparison.OrdinalIgnoreCase) && targetExtension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase)) return "spreadsheet-xml-xlsx";
        if (!targetExtension.Equals(".pdf", StringComparison.OrdinalIgnoreCase)) return null;
        if (sourceExtension.Equals(".txt", StringComparison.OrdinalIgnoreCase)) return "txt-pdf";
        // Native diagram pages retain physical layout; Reader remains an explicit semantic route.
        return sourceExtension.Equals(".vsdx", StringComparison.OrdinalIgnoreCase)
            || sourceExtension.Equals(".vstx", StringComparison.OrdinalIgnoreCase)
            || sourceExtension.Equals(".vssx", StringComparison.OrdinalIgnoreCase)
            || sourceExtension.Equals(".vsdm", StringComparison.OrdinalIgnoreCase)
            || sourceExtension.Equals(".vstm", StringComparison.OrdinalIgnoreCase)
            || sourceExtension.Equals(".vssm", StringComparison.OrdinalIgnoreCase)
            || sourceExtension.Equals(".vdx", StringComparison.OrdinalIgnoreCase)
            || sourceExtension.Equals(".vsx", StringComparison.OrdinalIgnoreCase)
            || sourceExtension.Equals(".vtx", StringComparison.OrdinalIgnoreCase) ? "visio-pdf" : null;
    }
}
