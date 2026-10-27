namespace OfficeIMO;

public sealed partial class OfficeConversionOptions {
    internal OfficeConversionOptions ForRoute(string routeId) => ForRoute(OfficeConverter.FindExecutionRoute(routeId)
        ?? throw new ArgumentException("Choose an executable conversion route.", nameof(routeId)));

    internal OfficeConversionOptions ForRoute(OfficeConversionExecutionRoute route) {
        OfficeConversionOptions copy = Snapshot();
        string id = route.Id;
        copy.RouteId = id;
        if (id is not ("doc-pdf" or "docx-pdf")) copy.Word = null;
        if (id is not ("xlsx-pdf" or "xls-pdf" or "xlsb-pdf")) copy.Excel = null;
        if (id is not ("pptx-pdf" or "ppt-pdf")) copy.PowerPoint = null;
        if (id != "markdown-pdf") copy.Markdown = null;
        if (id != "html-pdf") copy.Html = null;
        if (id != "docx-html") copy.WordHtml = null;
        if (id != "docx-markdown") copy.WordMarkdown = null;
        if (id != "markdown-docx") copy.MarkdownWord = null;
        if (id != "markdown-html") copy.MarkdownHtml = null;
        if (id != "html-markdown") copy.HtmlMarkdown = null;
        if (id != "html-docx") copy.HtmlWord = null;
        if (id != "pdf-docx") copy.PdfWord = null;
        if (id != "pdf-xlsx") copy.PdfExcel = null;
        if (id != "pdf-pptx") copy.PdfPowerPoint = null;
        if (id != "pdf-html") copy.PdfHtml = null;
        if (id != "txt-pdf") copy.PlainText = null;
        if (!route.SourceExtensions.Any(extension => OfficeConverter.FormatForExtension(extension) == OfficeDocumentFormat.Word)) copy.WordLoad = null;
        if (!route.SourceExtensions.Any(extension => OfficeConverter.FormatForExtension(extension) == OfficeDocumentFormat.Excel)) copy.ExcelLoad = null;
        if (!route.SourceExtensions.Any(extension => OfficeConverter.FormatForExtension(extension) == OfficeDocumentFormat.PowerPoint)) copy.PowerPointLoad = null;
        if (!route.SourceExtensions.Any(extension => OfficeConverter.FormatForExtension(extension) == OfficeDocumentFormat.Pdf)) copy.PdfLoad = null;
        if (!route.SourceExtensions.Any(extension => OfficeConverter.FormatForExtension(extension) == OfficeDocumentFormat.Html)) copy.HtmlLoad = null;
        if (OfficeConverter.FormatForExtension(route.TargetExtension) != OfficeDocumentFormat.Word) copy.WordSave = null;
        if (OfficeConverter.FormatForExtension(route.TargetExtension) != OfficeDocumentFormat.Excel) copy.ExcelSave = null;
        if (OfficeConverter.FormatForExtension(route.TargetExtension) != OfficeDocumentFormat.PowerPoint) copy.PowerPointSave = null;
        if (id is not ("email-pdf" or "epub-pdf" or "reader-visio-pdf")) { copy.Reader = null; copy.ReaderPdf = null; copy.ReaderPipeline = null; }
        if (OfficeConverter.IsBuiltinRoute(id)) copy.RegisteredSettings = null;
        return copy;
    }
}
