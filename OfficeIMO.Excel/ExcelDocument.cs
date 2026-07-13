using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Excel.Utilities;
using OfficeIMO.Drawing;
using OfficeIMO.Drawing.Internal;
using System.IO.Packaging;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using System;
using System.Diagnostics;
using System.IO;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Represents an Excel document and provides methods for creating,
    /// loading and saving spreadsheets.
    /// </summary>
    public partial class ExcelDocument : IDisposable, IAsyncDisposable {
        private const int StreamBufferSize = 4096;
        private static readonly System.Text.RegularExpressions.Regex _multipleUnderscoresRegex =
            new System.Text.RegularExpressions.Regex("_+", System.Text.RegularExpressions.RegexOptions.Compiled);

        private static readonly Lazy<byte[]> DefaultThemeBytes = new(() => LoadEmbeddedResource("OfficeIMO.Excel.Resources.theme1.xml"));
        // Allocated only when an operation actually needs a serialized apply stage
        internal ReaderWriterLockSlim? _lock;
        internal List<UInt32Value> id = new List<UInt32Value>() { 0 };
        private readonly Dictionary<string, int> _sharedStringCache = new Dictionary<string, int>();
        private Dictionary<string, bool>? _sharedStringLineBreakCache;
        private readonly object _sharedStringLock = new object();
        private int _sharedStringTableCount = -1;
        // Workbook-level cache of table names for fast uniqueness checks
        private HashSet<string>? _tableNameCache;
        private readonly object _tableMetadataLock = new object();
        private uint? _nextTableId;
        private System.Collections.Generic.IEqualityComparer<string> _tableNameComparer = System.StringComparer.OrdinalIgnoreCase;
        private List<ExcelSheet>? _cachedSheets;
        private bool _sheetCacheDirty = true;
        private bool _customDocumentPropertiesDirty;
        private DocumentPersistenceMode _persistenceMode = DocumentPersistenceMode.Explicit;

        /// <summary>
        /// Enables caching of <see cref="ExcelSheet"/> wrappers for faster repeat access at the cost of higher memory usage.
        /// Set to <see langword="false"/> to avoid holding references to every sheet in very large workbooks.
        /// </summary>
        public bool SheetCachingEnabled { get; set; } = true;

        /// <summary>
        /// Controls how workbook-level table name uniqueness is compared.
        /// Defaults to <see cref="StringComparer.OrdinalIgnoreCase"/>. Changing this will reset the
        /// internal cache and rebuild it on next use. Set it once before adding tables for predictable behavior.
        /// </summary>
        public System.Collections.Generic.IEqualityComparer<string> TableNameComparer {
            get => _tableNameComparer;
            set {
                if (value == null) throw new System.ArgumentNullException(nameof(value));
                if (!object.ReferenceEquals(_tableNameComparer, value)) {
                    _tableNameComparer = value;
                    _tableNameCache = null; // rebuild lazily on next use with the new comparer
                }
            }
        }

        /// <summary>
        /// Optional default chart style preset applied to charts created in this workbook.
        /// </summary>
        public ExcelChartStylePreset? DefaultChartStylePreset { get; set; }

        /// <summary>
        /// Execution policy for controlling parallel vs sequential operations.
        /// </summary>
        public ExecutionPolicy Execution { get; } = new();

        // Default strategy mirrors CoerceValueHelper's behaviour and uses LocalDateTime so that
        // serial values are aligned with Excel's local time interpretation.
        private Func<DateTimeOffset, DateTime> _dateTimeOffsetWriteStrategy = static dto => dto.LocalDateTime;

        /// <summary>
        /// Controls how <see cref="DateTimeOffset"/> values are converted to <see cref="DateTime"/>
        /// before being written to worksheet cells. Defaults to <see cref="DateTimeOffset.LocalDateTime"/>.
        /// </summary>
        /// <remarks>
        /// The delegate influences the numeric serial value stored in the cell but does not automatically
        /// change number formats. Apply the desired cell formatting separately.
        /// </remarks>
        public Func<DateTimeOffset, DateTime> DateTimeOffsetWriteStrategy {
            get => _dateTimeOffsetWriteStrategy;
            set => _dateTimeOffsetWriteStrategy = value ?? throw new ArgumentNullException(nameof(value));
        }

        internal ReaderWriterLockSlim EnsureLock()
            => _lock ??= new ReaderWriterLockSlim(); // default: NoRecursion

        internal bool EnsureWorkbookThemeAndStyles() {
            var workbookPart = _spreadSheetDocument?.WorkbookPart ?? _workBookPart;
            bool changed = false;

            if (!workbookPart.GetPartsOfType<ThemePart>().Any()) {
                ThemePart themePart = workbookPart.AddNewPart<ThemePart>();
                using var themeStream = new MemoryStream(DefaultThemeBytes.Value);
                themePart.FeedData(themeStream);
                changed = true;
            }

            var stylesPart = workbookPart.WorkbookStylesPart;
            if (stylesPart == null) {
                stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
                changed = true;
            }

            if (stylesPart.Stylesheet == null) {
                stylesPart.Stylesheet = CreateDefaultStylesheet();
                stylesPart.Stylesheet.Save();
                changed = true;
            }

            return changed;
        }

        private static Stylesheet CreateDefaultStylesheet() {
            var stylesheet = new Stylesheet();

            stylesheet.Fonts = new Fonts(new Font(new FontSize { Val = 11D }, new FontName { Val = "Calibri" }));
            stylesheet.Fonts.Count = (uint)stylesheet.Fonts.Count();

            stylesheet.Fills = new Fills(
                new Fill(new PatternFill { PatternType = PatternValues.None }),
                new Fill(new PatternFill { PatternType = PatternValues.Gray125 })
            );
            stylesheet.Fills.Count = (uint)stylesheet.Fills.Count();

            stylesheet.Borders = new Borders(new Border());
            stylesheet.Borders.Count = (uint)stylesheet.Borders.Count();

            stylesheet.CellStyleFormats = new CellStyleFormats(new CellFormat {
                NumberFormatId = 0U,
                FontId = 0U,
                FillId = 0U,
                BorderId = 0U
            });
            stylesheet.CellStyleFormats.Count = (uint)stylesheet.CellStyleFormats.Count();

            stylesheet.CellFormats = new CellFormats(new CellFormat {
                NumberFormatId = 0U,
                FontId = 0U,
                FillId = 0U,
                BorderId = 0U,
                FormatId = 0U
            });
            stylesheet.CellFormats.Count = (uint)stylesheet.CellFormats.Count();

            stylesheet.CellStyles = new CellStyles(new CellStyle {
                Name = "Normal",
                FormatId = 0U,
                BuiltinId = 0U
            });
            stylesheet.CellStyles.Count = (uint)stylesheet.CellStyles.Count();

            stylesheet.DifferentialFormats = new DifferentialFormats { Count = 0U };
            stylesheet.TableStyles = new TableStyles {
                Count = 0U,
                DefaultTableStyle = "TableStyleMedium2",
                DefaultPivotStyle = "PivotStyleLight16"
            };

            return stylesheet;
        }

        private static byte[] LoadEmbeddedResource(string resourceName) {
            var assembly = typeof(ExcelDocument).Assembly;
            using Stream? stream = assembly.GetManifestResourceStream(resourceName);
            if (stream == null) {
                throw new InvalidOperationException($"Missing embedded resource '{resourceName}'.");
            }

            using var buffer = new MemoryStream();
            stream.CopyTo(buffer);
            return buffer.ToArray();
        }

        private void MarkSheetCacheDirty()
        {
            _sheetCacheDirty = true;
            _cachedSheets = null;
            MarkPackageDirty();
        }

        private List<Sheet> ReadSheetElements()
        {
            var sheets = _spreadSheetDocument?.WorkbookPart?.Workbook?.Sheets;
            if (sheets == null) {
                return new List<Sheet>();
            }

            return sheets.OfType<Sheet>().ToList();
        }

        private void UpdateSheetIdCache(List<Sheet> elements)
        {
            id.Clear();
            id.Add(0);
            foreach (Sheet s in elements) {
                var sheetId = s.SheetId;
                if (sheetId != null && !id.Contains(sheetId)) {
                    id.Add(sheetId);
                }
            }
        }

        private List<ExcelSheet> MaterializeSheets(List<Sheet> elements)
        {
            List<ExcelSheet> listExcel = new List<ExcelSheet>(elements.Count);
            foreach (Sheet s in elements) {
                if (s.Id?.Value == null
                    || _spreadSheetDocument?.WorkbookPart?.GetPartById(s.Id.Value) is not WorksheetPart) {
                    continue;
                }

                listExcel.Add(new ExcelSheet(this, _spreadSheetDocument!, s));
            }

            return listExcel;
        }

        private void RebuildSheetCacheLocked()
        {
            var elements = ReadSheetElements();
            UpdateSheetIdCache(elements);
            _cachedSheets = SheetCachingEnabled ? MaterializeSheets(elements) : null;
            _sheetCacheDirty = false;
        }

        private void EnsureSheetCacheInitialized(ReaderWriterLockSlim? lck)
        {
            if (!(_sheetCacheDirty || _cachedSheets == null)) return;

            if (Locking.IsNoLock || lck is null || lck.IsWriteLockHeld) {
                RebuildSheetCacheLocked();
                return;
            }

            lck.EnterWriteLock();
            try {
                if (_sheetCacheDirty || _cachedSheets == null) {
                    RebuildSheetCacheLocked();
                }
            } finally {
                lck.ExitWriteLock();
            }
        }

        private List<ExcelSheet> CloneSheetCache()
        {
            if (_cachedSheets == null) {
                return new List<ExcelSheet>();
            }

            return new List<ExcelSheet>(_cachedSheets);
        }

        private List<ExcelSheet> BuildSheetsWithoutCaching()
        {
            var elements = ReadSheetElements();
            UpdateSheetIdCache(elements);
            return MaterializeSheets(elements);
        }

        internal void InvalidateSheetCache()
        {
            Locking.ExecuteWrite(EnsureLock(), MarkSheetCacheDirty);
        }

        /// <summary>
        /// Gets a list of worksheets contained in the document.
        /// </summary>
        public List<ExcelSheet> Sheets {
            get {
                MaterializeDeferredDataSetImport();
                var lck = EnsureLock();
                if (Locking.IsNoLock || lck is null) {
                    if (SheetCachingEnabled) {
                        EnsureSheetCacheInitialized(lck);
                        return CloneSheetCache();
                    }

                    return BuildSheetsWithoutCaching();
                }

                if (!SheetCachingEnabled) {
                    lck.EnterReadLock();
                    try {
                        return BuildSheetsWithoutCaching();
                    } finally {
                        lck.ExitReadLock();
                    }
                }

                lck.EnterReadLock();
                try {
                    if (!(_sheetCacheDirty || _cachedSheets == null)) {
                        return CloneSheetCache();
                    }
                } finally {
                    lck.ExitReadLock();
                }

                lck.EnterUpgradeableReadLock();
                try {
                    if (_sheetCacheDirty || _cachedSheets == null) {
                        EnsureSheetCacheInitialized(lck);
                    }

                    return CloneSheetCache();
                } finally {
                    lck.ExitUpgradeableReadLock();
                }
            }
        }

        /// <summary>
        /// Underlying Open XML spreadsheet document instance.
        /// </summary>
        internal SpreadsheetDocument _spreadSheetDocument = null!;

        /// <summary>Gets the underlying Open XML package for advanced integration scenarios.</summary>
        public SpreadsheetDocument OpenXmlDocument => _spreadSheetDocument;
        private WorkbookPart _workBookPart = null!;
        private SharedStringTablePart? _sharedStringTablePart;
        private bool _sharedStringTableDirty;
        private Stream? _packageStream;
        private Stream? _sourceStream;
        private Stream? _ownedOpenStream;
        private bool _copyPackageToSourceOnDispose;
        private bool _copyPackageToFilePathOnDispose;
        private bool _leaveSourceStreamOpen = true;
        private bool _packageContentTypesKnownNormalized;
        private bool _requiresSavePreflight;
        private bool _packageDirty = true;
        private bool _packagePropertiesDirty;
        private bool _preserveDirectDataSetSaveCandidateForNextDirtyMark;
        private int _directDataSetSaveCandidatePreservationDepth;
        private int _materializedDirectDataSetFastSaveModelPreservationDepth;
        private byte[]? _unchangedPackageBytes;
        private bool _simplePackageContentKnown;
        private DirectDataSetSaveCandidate? _directDataSetSaveCandidate;
        private DirectDataSetWorkbookModel? _materializedDirectDataSetFastSaveModel;
        private bool _materializedDirectDataSetFastSaveModelHasMaterializedWorksheet;
        private ExcelSheet? _directDataSetMetadataSourceSheet;
        private ExcelSheet? _pendingDirectCellValueSheet;
        private bool _materializingDeferredDataSetImport;
        private bool _preserveMaterializedDirectDataSetFastSaveModelForNextDirtyMark;
        private int _directDataSetExternalCellMutationPreservationDepth;

        /// <summary>
        /// Diagnostics for the most recent save operation.
        /// </summary>
        internal ExcelSaveDiagnostics LastSaveDiagnostics { get; private set; } = ExcelSaveDiagnostics.Standard("Workbook has not been saved yet.");

        private const int StreamCopyBufferSize = 81920;

        internal void MarkRequiresSavePreflight() {
            _requiresSavePreflight = true;
            MarkPackageDirty();
        }

        internal void MarkPackageDirty() {
            bool preserveMaterializedDirectModel = _preserveMaterializedDirectDataSetFastSaveModelForNextDirtyMark
                || _materializedDirectDataSetFastSaveModelPreservationDepth > 0;
            if (!preserveMaterializedDirectModel && !_materializingDeferredDataSetImport) {
                _materializedDirectDataSetFastSaveModel = null;
                _materializedDirectDataSetFastSaveModelHasMaterializedWorksheet = false;
            }

            if (IsPackageDirtyWithoutPendingSaveCandidate) {
                _preserveMaterializedDirectDataSetFastSaveModelForNextDirtyMark = false;
                return;
            }

            _packageDirty = true;
            _unchangedPackageBytes = null;
            bool preserveDirectCandidate = _preserveDirectDataSetSaveCandidateForNextDirtyMark
                || _directDataSetSaveCandidatePreservationDepth > 0;
            try {
                if (_directDataSetSaveCandidate?.IsDeferred == true
                    && !_materializingDeferredDataSetImport
                    && !preserveDirectCandidate) {
                    MaterializeDeferredDataSetImport();
                } else if (!preserveDirectCandidate) {
                    ClearDirectDataSetSaveCandidate();
                }
            } finally {
                _preserveDirectDataSetSaveCandidateForNextDirtyMark = false;
                _preserveMaterializedDirectDataSetFastSaveModelForNextDirtyMark = false;
            }
        }

        internal void MarkPackagePropertiesDirty() {
            _packagePropertiesDirty = true;
            MarkPackageDirty();
        }

        internal bool IsPackageDirty => _packageDirty;

        internal bool IsPackageDirtyWithoutPendingSaveCandidate
            => _packageDirty
                && _unchangedPackageBytes == null
                && _directDataSetSaveCandidate == null
                && !_preserveDirectDataSetSaveCandidateForNextDirtyMark
                && _directDataSetSaveCandidatePreservationDepth == 0;

        internal bool HasPackagePropertiesDirty => _packagePropertiesDirty;

        internal bool IsMaterializingDeferredDataSetImport => _materializingDeferredDataSetImport;

        internal bool CanDeferDirectCellValuesAppendCandidate
            => _spreadSheetDocument != null && _sourceStream != null;

        internal bool IsPreservingDirectDataSetExternalCellMutation
            => _directDataSetExternalCellMutationPreservationDepth > 0;

        internal bool HasDirectDataSetFastSaveState
            => _materializedDirectDataSetFastSaveModel != null
               || _directDataSetSaveCandidate?.IsValid == true;

        internal void PreserveDirectDataSetSaveCandidateForNextDirtyMark() {
            _preserveDirectDataSetSaveCandidateForNextDirtyMark = true;
        }

        internal IDisposable PreserveDirectDataSetSaveCandidateDuringDirtyMarks() {
            _directDataSetSaveCandidatePreservationDepth++;
            return new DirectDataSetSaveCandidatePreservationScope(this);
        }

        internal IDisposable PreserveDirectDataSetFastSaveStateDuringDirtyMarks() {
            _directDataSetSaveCandidatePreservationDepth++;
            _materializedDirectDataSetFastSaveModelPreservationDepth++;
            return new DirectDataSetFastSaveStatePreservationScope(this);
        }

        internal IDisposable? PreserveDirectDataSetFastSaveStateForExternalCellMutation(ExcelSheet sheet, int row, int column) {
            var model = _materializedDirectDataSetFastSaveModel;
            bool hasMaterializedModel = model != null;
            if (sheet == null || !ReferenceEquals(sheet.Document, this) || row <= 0 || column <= 0) {
                return null;
            }

            if (model == null) {
                var candidate = _directDataSetSaveCandidate;
                if (candidate == null || !candidate.IsValid) {
                    return null;
                }

                model = candidate.Model;
            }

            for (int i = 0; i < model.Sheets.Count; i++) {
                var sheetModel = model.Sheets[i];
                if (!string.Equals(sheetModel.SheetName, sheet.Name, StringComparison.Ordinal)) {
                    continue;
                }

                int lastDirectRow = sheetModel.Table.RowCount + (sheetModel.IncludeHeaders ? 1 : 0);
                if (row > lastDirectRow || column > sheetModel.Table.ColumnCount) {
                    _directDataSetMetadataSourceSheet = sheet;
                    _directDataSetSaveCandidatePreservationDepth++;
                    _materializedDirectDataSetFastSaveModelPreservationDepth++;
                    _directDataSetExternalCellMutationPreservationDepth++;
                    return new DirectDataSetExternalCellMutationPreservationScope(this);
                }

                if (hasMaterializedModel) {
                    _materializedDirectDataSetFastSaveModel = null;
                    _materializedDirectDataSetFastSaveModelHasMaterializedWorksheet = false;
                } else {
                    ClearDirectDataSetSaveCandidate();
                }

                _materializingDeferredDataSetImport = true;
                try {
                    MaterializeDirectDataSetModel(model);
                } finally {
                    _materializingDeferredDataSetImport = false;
                }

                return null;
            }

            return null;
        }

        private sealed class DirectDataSetSaveCandidatePreservationScope : IDisposable {
            private ExcelDocument? _document;

            internal DirectDataSetSaveCandidatePreservationScope(ExcelDocument document) {
                _document = document;
            }

            public void Dispose() {
                var document = _document;
                if (document == null) {
                    return;
                }

                _document = null;
                if (document._directDataSetSaveCandidatePreservationDepth > 0) {
                    document._directDataSetSaveCandidatePreservationDepth--;
                }
            }
        }

        private sealed class DirectDataSetFastSaveStatePreservationScope : IDisposable {
            private ExcelDocument? _document;

            internal DirectDataSetFastSaveStatePreservationScope(ExcelDocument document) {
                _document = document;
            }

            public void Dispose() {
                var document = _document;
                if (document == null) {
                    return;
                }

                _document = null;
                if (document._directDataSetSaveCandidatePreservationDepth > 0) {
                    document._directDataSetSaveCandidatePreservationDepth--;
                }

                if (document._materializedDirectDataSetFastSaveModelPreservationDepth > 0) {
                    document._materializedDirectDataSetFastSaveModelPreservationDepth--;
                }
            }
        }

        private sealed class DirectDataSetExternalCellMutationPreservationScope : IDisposable {
            private ExcelDocument? _document;

            internal DirectDataSetExternalCellMutationPreservationScope(ExcelDocument document) {
                _document = document;
            }

            public void Dispose() {
                var document = _document;
                if (document == null) {
                    return;
                }

                _document = null;
                if (document._directDataSetSaveCandidatePreservationDepth > 0) {
                    document._directDataSetSaveCandidatePreservationDepth--;
                }

                if (document._materializedDirectDataSetFastSaveModelPreservationDepth > 0) {
                    document._materializedDirectDataSetFastSaveModelPreservationDepth--;
                }

                if (document._directDataSetExternalCellMutationPreservationDepth > 0) {
                    document._directDataSetExternalCellMutationPreservationDepth--;
                }
            }
        }
    }
}
