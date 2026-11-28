namespace OfficeIMO.Excel {
    /// <summary>
    /// A worksheet of a workbook whose Open XML SDK package has not been opened. It knows its name and part from the
    /// workbook the load read, and binds to the SDK's sheet element and part when the document opens the SDK.
    /// </summary>
    public partial class ExcelSheet {
        private ExcelDeferredSheet? _deferredEntry;

        /// <summary>The changes this sheet applies through the model once the SDK package has opened.</summary>
        internal ExcelDeferredCellStore? DeferredCellsForReplay { get; set; }

        internal ExcelSheet(ExcelDocument excelDocument, ExcelDeferredSheet entry) {
            _excelDocument = excelDocument;
            _deferredEntry = entry;
            excelDocument.RegisterSheetWrapper(this);
            Interlocked.Increment(ref _instancesCreated);
        }

        /// <summary>The workbook entry this sheet was read from while the SDK package was not open; null once bound.</summary>
        internal ExcelDeferredSheet? DeferredEntry => _worksheetPartValue == null ? _deferredEntry : null;

        /// <summary>Opens the document's SDK package, which binds this sheet, and returns the sheet.</summary>
        private ExcelSheet BindDeferredSheet() {
            if (_worksheetPartValue != null) return this;
            if (_deferredEntry == null) throw new InvalidOperationException("The worksheet is not bound to a workbook part.");
            _excelDocument.OpenDeferredSdk(null);
            if (_worksheetPartValue == null) {
                _excelDocument.ThrowIfDisposed();
                throw new InvalidOperationException("The worksheet '" + _deferredEntry.Name + "' is no longer part of the workbook.");
            }
            return this;
        }
    }
}
