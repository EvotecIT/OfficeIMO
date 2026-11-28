using OfficeIMO.Core.Internal;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Excel.Xlsb.Write;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private static readonly Lazy<byte[]> DefaultNativeXlsbStyles =
            new(() => XlsbStylesheetPartWriter.CreateFromMarkup(CreatedStylesXml(theme: null)));
        private static readonly Lazy<byte[]> DefaultNativeXlsbTheme = new(() => {
            byte[] declaration = Encoding.UTF8.GetBytes("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
            byte[] theme = DefaultThemeBytes.Value;
            byte[] part = new byte[declaration.Length + theme.Length];
            Buffer.BlockCopy(declaration, 0, part, 0, declaration.Length);
            Buffer.BlockCopy(theme, 0, part, declaration.Length, theme.Length);
            return part;
        });

        /// <summary>Encodes plain worksheet values from a new workbook's deferred store before opening the SDK.</summary>
        private bool TryPrepareOwnedXlsbCreation(Stream destination, ExcelSaveOptions? options,
            CancellationToken cancellationToken, out ArraySegment<byte> package) {
            package = default;
            if (destination.GetType() != typeof(MemoryStream) || destination.Length != 0 || destination.Position != 0
                || !IsSdkDeferred || !_deferredCreated || options?.DisableFastPackageWriter == true) return false;

            ArraySegment<byte> output = WithDeferredOpenGate(() => {
                if (_deferredWorkbook is not { Sheets.Count: 1 } workbook || !_deferredCreated
                    || _defaultFontTheme != null || _packagePropertiesDirty || _customDocumentPropertiesDirty
                    || _requiresSavePreflight || HasCalculationSaveWork(options)
                    || workbook.Sheets[0] is not { IsWorksheet: true, IsNewSheet: true } entry) {
                    return default(ArraySegment<byte>);
                }
                cancellationToken.ThrowIfCancellationRequested();
                var edits = TakeDeferredEdits();
                try {
                    if (edits.Length != 1 || !ReferenceEquals(edits[0].Sheet.DeferredEntry, entry)
                        || edits[0].Store.Count < ExcelSheet.PendingDirectCellValueMinimumCellCount
                        || edits[0].Store.Tabular is not { HasCompactRectangle: true, HasCellStyles: false } cells
                        || !HasDefaultOwnedXlsbFormatting(workbook)) {
                        return default;
                    }

                    var source = new ExcelDirectTabularSource(entry.Name, DirectDataSetTableModel.FromCellValueBuffer(cells),
                        includeHeaders: false, preserveMissingValues: true, useCellValueNumberFormats: true);
                    if (!XlsbDirectTabularPlan.CanCapture(source)) return default;
                    var dateSystem = workbook.Date1904 ? ExcelDateSystem.NineteenFour : ExcelDateSystem.NineteenHundred;
                    XlsbSharedStringTable? sharedStrings = options?.XlsbUseSharedStrings == true ? new XlsbSharedStringTable() : null;
                    if (!XlsbDirectTabularPlan.TryCreateOwned(dateSystem, source, cells, sharedStrings, cancellationToken, out var plan))
                        return default;

                    using (plan) {
                        long limit = GetXlsbMaterializationLimit(options);
                        using var staged = new OfficeBoundedMemoryStream(limit);
                        XlsbNewPackageWriter.WriteOwnedTabular(staged, dateSystem, source.SheetName,
                            DefaultNativeXlsbStyles.Value, DefaultNativeXlsbTheme.Value, sharedStrings, plan, limit, cancellationToken);
                        return new ArraySegment<byte>(staged.GetBuffer(), 0, checked((int)staged.Length));
                    }
                } finally {
                    // A stream save writes a copy. The captured cells remain editable in their original worksheet.
                    RestoreDeferredEdits(edits);
                }
            });

            package = output;
            return output.Array != null;
        }

        private bool HasDefaultOwnedXlsbFormatting(ExcelDeferredWorkbook workbook) {
            using var source = new MemoryStream(_deferredPackageBytes!, writable: false);
            using OfficeOpcPackage package = OfficeOpcPackage.Open(source, leaveOpen: true);
            if (package.GetContentType(workbook.WorkbookPartName) != CreatedMainContentType(ExcelDocumentType.Workbook))
                return false;

            string? themePart = package.GetRelatedPartName(workbook.WorkbookPartName, ThemeRelationshipType);
            if (themePart == null || !package.PartExists(themePart)) return false;
            using OfficePooledBytes theme = package.ReadPart(themePart);
            return theme.Array.AsSpan(0, theme.Length).SequenceEqual(DefaultThemeBytes.Value);
        }
    }
}
