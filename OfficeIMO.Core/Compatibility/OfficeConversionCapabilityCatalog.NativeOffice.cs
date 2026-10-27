using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO;

public static partial class OfficeConversionCapabilityCatalog {
    private static void AddNativeOfficeRoutes(List<OfficeConversionCapability> routes) {
        AddNativeFamily(routes, "word", "Word", new[] { ".doc", ".dot", ".docx", ".docm", ".dotx", ".dotm" },
            new[] { ".doc", ".docx", ".docm", ".dotx", ".dotm" });
        AddNativeFamily(routes, "excel", "Excel", new[] { ".xls", ".xlt", ".xla", ".xlm", ".xlw", ".xlsx", ".xlsm", ".xltx", ".xltm", ".xlam", ".xlsb" },
            new[] { ".xls", ".xlsx", ".xlsm", ".xltx", ".xltm", ".xlam", ".xlsb" });
        AddNativeFamily(routes, "powerpoint", "PowerPoint", new[] { ".ppt", ".pot", ".pps", ".ppa", ".pptx", ".pptm", ".potx", ".potm", ".ppsx", ".ppsm", ".ppam" },
            new[] { ".ppt", ".pot", ".pps", ".pptx", ".pptm", ".potx", ".potm", ".ppsx", ".ppsm", ".ppam" });
        routes.Add(Route("xls-pdf", "XLS", "PDF", OfficeConversionInputKind.File, new[] { ".xls", ".xlt", ".xla", ".xlm", ".xlw" }, ".pdf", "OfficeIMO.Excel.Pdf",
            "ExcelDocument.Load(source, options, out var importReport).SaveAsPdf(destination)", "Project supported binary workbook content and render it with separate import and PDF findings.", OfficeConversionFidelityKind.FixedLayout, "OfficeConversionReport"));
        routes.Add(Route("xlsb-pdf", "XLSB", "PDF", OfficeConversionInputKind.File, new[] { ".xlsb" }, ".pdf", "OfficeIMO.Excel.Pdf",
            "ExcelDocument.Load(source, options, out var importReport).SaveAsPdf(destination)", "Project supported XLSB records and render workbook content with import and PDF findings.", OfficeConversionFidelityKind.FixedLayout, "OfficeConversionReport"));
        routes.Add(Route("ppt-pdf", "PPT", "PDF", OfficeConversionInputKind.File, new[] { ".ppt", ".pot", ".pps", ".ppa" }, ".pdf", "OfficeIMO.PowerPoint.Pdf",
            "PowerPointDocument.Load(source, options, out var importReport).SaveAsPdf(destination)", "Project supported binary presentation content and render slides with import and PDF findings.", OfficeConversionFidelityKind.FixedLayout, "OfficeConversionReport"));
    }

    private static void AddNativeFamily(List<OfficeConversionCapability> routes, string id, string name, string[] sources, string[] targets) {
        foreach (string target in targets) routes.Add(Route("native-" + id + "-" + target.Substring(1), name, target.Substring(1).ToUpperInvariant(),
            OfficeConversionInputKind.File, sources.Where(source => source != target).ToArray(), target, "OfficeIMO",
            "OfficeConverter.Convert(source, destination, options)", "Convert native " + name + " variants through their owning codecs with explicit preservation and loss policy.",
            OfficeConversionFidelityKind.Editable, "OfficeConversionReport"));
    }
}
