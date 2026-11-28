using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private SpreadsheetDocument? _workbookPartOwner;
        private WorkbookPart? _knownWorkbookPart;

        /// <summary>
        /// The SDK resolves the main part through feature lookups that allocate, and this runs per cell. Only exposed
        /// SDK access and a document type change can replace the part of an open document.
        /// </summary>
        internal WorkbookPart WorkbookPartRoot {
            get {
                SpreadsheetDocument? document = _spreadSheetDocument;
                if (ReferenceEquals(document, _workbookPartOwner) && _knownWorkbookPart != null && !_openXmlDocumentExposed) {
                    return _knownWorkbookPart;
                }

                WorkbookPart? part = document?.WorkbookPart;
                if (part != null) {
                    _workbookPartOwner = document;
                    _knownWorkbookPart = part;
                }
                return part ?? _workBookPart ?? throw new InvalidOperationException("WorkbookPart is null.");
            }
        }

        private void ForgetKnownWorkbookPart() {
            _workbookPartOwner = null;
            _knownWorkbookPart = null;
        }

        internal Workbook WorkbookRoot {
            get => WorkbookPartRoot.Workbook ?? throw new InvalidOperationException("Workbook is null.");
            set => WorkbookPartRoot.Workbook = value;
        }
    }

    public partial class ExcelSheet {
        private WorkbookPart WorkbookPartRoot => _excelDocument.WorkbookPartRoot;

        private Workbook WorkbookRoot {
            get => _excelDocument.WorkbookRoot;
            set => _excelDocument.WorkbookRoot = value;
        }

        internal Worksheet WorksheetRoot {
            get {
                if (!_worksheetPart.IsRootElementLoaded) LoadWorksheetRoot();
                return _worksheetPart.Worksheet ?? throw new InvalidOperationException("Worksheet is null.");
            }
            private set => _worksheetPart.Worksheet = value;
        }

        private void LoadWorksheetRoot() {
            // A part read from a package stream can be opened by one reader at a time, so threads load the root in turn.
            lock (_worksheetPart) {
                if (_worksheetPart.IsRootElementLoaded) return;
                CaptureOriginalDynamicSpillFingerprintIfSafe();
                if (_excelDocument.AllowsWorksheetGridLoader
                    && ExcelWorksheetGridLoader.TryLoad(_worksheetPart, storeCells: true, out int emptyStyleCells, out string? firstEmptyStyleCell)) {
                    Persistence.UsesCellStore = true;
                    if (emptyStyleCells > 0) RecordEmptyCellStyleRepair(emptyStyleCells, firstEmptyStyleCell);
                    return;
                }
                _ = _worksheetPart.Worksheet;
            }
        }

        private WorksheetCommentsPart? WorksheetCommentsPartRoot => _worksheetPart.WorksheetCommentsPart;
    }

    // The table APIs still use the editing model. Open it only when one of those APIs is called.
    internal sealed partial class ExcelDocumentReader {
        private IDisposable? _tableDocument;
        private Stream? _tableStream;

        private WorkbookPart WorkbookPartRoot {
            get {
                if (_liveWorkbook?.SdkPart is WorkbookPart live) return live;
                if (_tableDocument is not SpreadsheetDocument document) {
                    Stream stream = _package!.OpenSource();
                    try {
                        document = OpenTableWorkbook(ref stream);
                        _tableDocument = document;
                        _tableStream = stream;
                    } catch {
                        stream.Dispose();
                        throw;
                    }
                }
                return document.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart is null.");
            }
        }

        private SpreadsheetDocument OpenTableWorkbook(ref Stream stream) {
            try {
                return OpenValidatedTableWorkbook(stream);
            } catch (Exception exception) when (exception is OpenXmlPackageException || CanRetryWithNormalizedContentTypes(exception)) {
                stream.Dispose();
                using Stream source = _package!.OpenSource();
                stream = OpenNormalizedPackageStream(ReadInput(source, _opt), _opt.CancellationToken);
                return OpenValidatedTableWorkbook(stream);
            }
        }

        private static SpreadsheetDocument OpenValidatedTableWorkbook(Stream stream) {
            SpreadsheetDocument document = OfficeIMO.OpenXml.Internal.OfficeOpenXmlReadOnlyPackage.OpenSpreadsheet(stream, new OpenSettings());
            try {
                // The SDK validates the typed relationships on first access to WorkbookPart.
                _ = document.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart is null.");
                return document;
            } catch {
                document.Dispose();
                throw;
            }
        }

        private Workbook WorkbookRoot =>
            WorkbookPartRoot.Workbook ?? throw new InvalidOperationException("Workbook is null.");

        private void DisposeTableWorkbook() {
            try {
                _tableDocument?.Dispose();
            } finally {
                _tableStream?.Dispose();
            }
        }
    }
}
