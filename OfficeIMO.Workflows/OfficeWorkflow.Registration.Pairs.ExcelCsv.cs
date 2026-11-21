using OfficeIMO.CSV;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Csv;

namespace OfficeIMO.Workflows;

public static partial class OfficeWorkflow {
    private sealed partial class PairRoute {
        // CSV and workbooks convert through the Excel.Csv adapter: delimited values become one typed worksheet,
        // and the first worksheet's used range becomes delimited values.
        private async Task<OfficeConversionReport> ConvertExcelCsvAsync(Stream input, Stream output,
            OfficeConversionOptions options, CancellationToken token) {
            if (Id == "xlsx-csv") {
                using ExcelDocument workbook = await ExcelDocument.LoadAsync(input, OfficeConverter.ExcelInputOptions(options), token).ConfigureAwait(false);
                if (workbook.Sheets.Count == 0) throw new InvalidDataException("The workbook has no worksheets to export as CSV.");
                var export = new ExcelCsvExportOptions { FailOnLoss = options.FailOnLoss };
                return await RetainSourceAsync(workbook.CreateLoadReport(options.SourceName!, TimeSpan.Zero),
                    () => workbook.Sheets[0].SaveAsCsvAsync(output, export, token), options.FailOnLoss).ConfigureAwait(false);
            }
            bool tabSeparated = string.Equals(Path.GetExtension(options.SourceName), ".tsv", StringComparison.OrdinalIgnoreCase);
            using CsvDocument csv = await CsvDocument.LoadAsync(input,
                CsvProfiles.CreateLoadOptions(tabSeparated ? CsvProfile.Tsv : CsvProfile.Spreadsheet), token).ConfigureAwait(false);
            return await SaveModelAsync(await csv.ToExcelAsync(new CsvExcelExportOptions { FailOnLoss = options.FailOnLoss }, token).ConfigureAwait(false),
                model => model.SaveAsync(output, null, token), token).ConfigureAwait(false);
        }
    }
}
