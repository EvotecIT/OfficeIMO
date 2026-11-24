using OfficeIMO.Notebook;
using OfficeIMO.Notebook.Html;
using OfficeIMO.Notebook.Pdf;

if (args.Length != 3) {
    throw new ArgumentException("Usage: OfficeNotebookExport <pdf|widget-html> <notebook> <destination>");
}

using var notebook = NotebookDocument.Load(args[1]);
switch (args[0]) {
    case "pdf":
        await notebook.SaveAsPdfAsync(args[2], new NotebookPdfExportOptions {
            IncludeCode = false,
            IncludeExecutionCounts = false
        });
        break;
    case "widget-html":
        await notebook.SaveAsHtmlAsync(args[2], new NotebookHtmlExportOptions {
            IncludeInteractiveWidgets = true,
            ProjectionOptions = new() { IncludeCode = false, IncludeExecutionCounts = false }
        });
        break;
    default:
        throw new ArgumentException("Unknown export format.", nameof(args));
}
