using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        internal WorkbookPart WorkbookPartRoot =>
            _spreadSheetDocument?.WorkbookPart ?? _workBookPart ?? throw new InvalidOperationException("WorkbookPart is null.");

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
            get => _worksheetPart.Worksheet ?? throw new InvalidOperationException("Worksheet is null.");
            set => _worksheetPart.Worksheet = value;
        }

        private WorksheetCommentsPart? WorksheetCommentsPartRoot => _worksheetPart.WorksheetCommentsPart;
    }

    public sealed partial class ExcelDocumentReader {
        private WorkbookPart WorkbookPartRoot =>
            _doc.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart is null.");

        private Workbook WorkbookRoot =>
            WorkbookPartRoot.Workbook ?? throw new InvalidOperationException("Workbook is null.");
    }

    public sealed partial class ExcelSheetReader {
        private Worksheet WorksheetRoot =>
            _wsPart.Worksheet ?? throw new InvalidOperationException("Worksheet is null.");
    }
}
