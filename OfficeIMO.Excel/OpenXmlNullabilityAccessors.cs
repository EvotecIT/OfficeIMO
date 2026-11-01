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

        private Worksheet WorksheetRoot {
            get {
                CaptureOriginalDynamicSpillFingerprintIfSafe();
                return _worksheetPart.Worksheet ?? throw new InvalidOperationException("Worksheet is null.");
            }
            set => _worksheetPart.Worksheet = value;
        }

        private WorksheetCommentsPart? WorksheetCommentsPartRoot => _worksheetPart.WorksheetCommentsPart;
    }

    internal sealed partial class ExcelDocumentReader {
        private WorkbookPart WorkbookPartRoot =>
            _doc.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart is null.");

        private Workbook WorkbookRoot =>
            WorkbookPartRoot.Workbook ?? throw new InvalidOperationException("Workbook is null.");
    }

    internal sealed partial class ExcelSheetReader {
        private Worksheet WorksheetRoot =>
            _wsPart.Worksheet ?? throw new InvalidOperationException("Worksheet is null.");
    }
}
