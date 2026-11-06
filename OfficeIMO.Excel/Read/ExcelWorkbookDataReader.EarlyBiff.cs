using System.Data.Common;
using System.Threading;
using OfficeIMO.Excel.Legacy;

namespace OfficeIMO.Excel {
    public sealed partial class ExcelWorkbookDataReader {
        private static readonly ExcelReaderOptions EarlyBiffReadDefaults = new();

        private static bool CanReadEarlyBiffRows(ExcelReaderOptions options) =>
            CanUseLegacyFastPath(options)
            && !options.NumericAsDecimal
            // These limits belong to the existing projected package's string/XML caches.
            && options.MaxSharedStringItems == EarlyBiffReadDefaults.MaxSharedStringItems
            && options.MaxSharedStringItemCharacters == EarlyBiffReadDefaults.MaxSharedStringItemCharacters
            && options.MaxSharedStringCharacters == EarlyBiffReadDefaults.MaxSharedStringCharacters
            && options.MaxXmlDataReaderBufferedCharacters == EarlyBiffReadDefaults.MaxXmlDataReaderBufferedCharacters;

        private static ExcelWorkbookDataReader OpenEarlyBiffRows(byte[] bytes, ExcelReaderOptions options) {
            var limits = new OfficeLegacyImportLimits();
            limits.MaxInputBytes = (int)Math.Min(limits.MaxInputBytes, options.MaxInputBytes);
            if (bytes.Length > limits.MaxInputBytes) {
                string limitName = limits.MaxInputBytes < options.MaxInputBytes
                    ? nameof(ExcelLoadOptions.MaxLegacyInputBytes) : nameof(ExcelReaderOptions.MaxInputBytes);
                throw ExcelReadLimitFailure.Create(limitName, limits.MaxInputBytes,
                    "Early BIFF input exceeds the configured byte limit.");
            }
            (LegacySpreadsheetModel model, _) = EarlyBiffWorkbook.Parse(bytes, limits, options.CancellationToken);
            if (model.Sheets.Count > options.MaxWorksheets) {
                throw ExcelReadLimitFailure.Create(nameof(ExcelReaderOptions.MaxWorksheets), options.MaxWorksheets,
                    "Early BIFF input exceeds the configured worksheet limit.");
            }
            var names = new string[model.Sheets.Count];
            var existing = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int index = 0; index < names.Length; index++) {
                options.CancellationToken.ThrowIfCancellationRequested();
                LegacySpreadsheetSheet sheet = model.Sheets[index];
                names[index] = ExcelDocument.PrepareStrictWorksheetName(sheet.Name, existing);
                existing.Add(names[index]);
                foreach (LegacySpreadsheetCell cell in sheet.Cells) {
                    options.CancellationToken.ThrowIfCancellationRequested();
                    if (cell.Value is string text && string.IsNullOrWhiteSpace(cell.Formula))
                        CoerceValueHelper.ValidateSharedStringLength(text, "value");
                }
            }
            var owner = new EarlyBiffRowOwner(model);
            try {
                IReadOnlyList<SheetSelection> sheets = SelectSheets(names, options);
                return new ExcelWorkbookDataReader(sheets,
                    (index, cancellationToken) => owner.OpenSheet(sheets[index].WorkbookIndex, options, cancellationToken),
                    owner, options.Culture, options.TypeConverter, options.StrictTypedMapping,
                    options.MappingErrorValuePolicy, options.CancellationToken);
            } catch {
                owner.Dispose();
                throw;
            }
        }

        private sealed class EarlyBiffRowOwner : IDisposable {
            private LegacySpreadsheetModel? _model;

            internal EarlyBiffRowOwner(LegacySpreadsheetModel model) => _model = model;

            internal DbDataReader OpenSheet(int index, ExcelReaderOptions options, CancellationToken cancellationToken) {
                LegacySpreadsheetModel model = _model ?? throw new ObjectDisposedException(nameof(EarlyBiffRowOwner));
                return ExcelSheetReader.ReadEarlyBiffRows(model.Sheets[index],
                    model.DateSystem ?? ExcelDateSystem.NineteenHundred, options, cancellationToken);
            }

            public void Dispose() => _model = null;
        }
    }
}
