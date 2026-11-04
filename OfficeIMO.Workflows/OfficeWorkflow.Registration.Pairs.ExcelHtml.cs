using OfficeIMO.Excel;
using OfficeIMO.Excel.Html;
using OfficeIMO.Html;

namespace OfficeIMO.Workflows;

public static partial class OfficeWorkflow {
    private sealed partial class PairRoute {
        private async Task<OfficeConversionReport> ConvertExcelHtmlAsync(Stream input, Stream output,
            OfficeConversionOptions options, OfficeWorkflowRouteSettings settings, CancellationToken token) {
            if (Id == "xlsx-html") {
                using ExcelDocument document = await ExcelDocument.LoadAsync(input, OfficeConverter.ExcelInputOptions(options), token).ConfigureAwait(false);
                ExcelHtmlExportOptions conversion = settings.ExcelHtml ?? new();
                conversion.FailOnLoss |= options.FailOnLoss;
                return await RetainSourceAsync(document.CreateLoadReport(options.SourceName!, TimeSpan.Zero),
                    () => document.SaveAsHtmlAsync(output, conversion, token), conversion.FailOnLoss).ConfigureAwait(false);
            }
            HtmlConversionDocument html = await HtmlConversionDocument.LoadAsync(input, options.HtmlLoad, cancellationToken: token).ConfigureAwait(false);
            HtmlExcelExportOptions import = settings.HtmlExcel ?? new();
            import.FailOnLoss |= options.FailOnLoss;
            return await SaveModelAsync(await html.ToExcelAsync(import, token).ConfigureAwait(false),
                model => model.SaveAsync(output, null, token), token).ConfigureAwait(false);
        }
    }
}
