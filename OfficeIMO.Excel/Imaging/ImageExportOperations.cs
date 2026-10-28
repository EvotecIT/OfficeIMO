using OfficeIMO.Drawing;

namespace OfficeIMO.Excel {
    public sealed partial class ExcelChart {
        private OfficeImageExportOperation<ExcelImageExportOptions> CreateImageExportOperation(
            ExcelImageExportOptions? options, OfficeImageExportFormat format) {
            ExcelImageExportOptions snapshot = options?.Clone() ?? new ExcelImageExportOptions();
            return new OfficeImageExportOperation<ExcelImageExportOptions>(snapshot, format,
                OfficeImageExportFileConflictPolicy.FailIfExists,
                (encoding, effective) => this.ExportImage(encoding, effective),
                (encoding, effective, token) => this.ExportImage(encoding, effective, token));
        }
    }
    public sealed partial class ExcelRange {
        private OfficeImageExportOperation<ExcelImageExportOptions> CreateImageExportOperation(
            ExcelImageExportOptions? options, OfficeImageExportFormat format) {
            ExcelImageExportOptions snapshot = options?.Clone() ?? new ExcelImageExportOptions();
            return new OfficeImageExportOperation<ExcelImageExportOptions>(snapshot, format,
                OfficeImageExportFileConflictPolicy.FailIfExists,
                (encoding, effective) => this.ExportImage(encoding, effective),
                (encoding, effective, token) => this.ExportImage(encoding, effective, token));
        }
    }
    public partial class ExcelSheet {
        private OfficeImageExportOperation<ExcelWorksheetImageExportOptions> CreateImageExportOperation(
            ExcelWorksheetImageExportOptions? options, OfficeImageExportFormat format) {
            ExcelWorksheetImageExportOptions snapshot = options?.CloneWorksheet() ?? new ExcelWorksheetImageExportOptions();
            return new OfficeImageExportOperation<ExcelWorksheetImageExportOptions>(snapshot, format,
                OfficeImageExportFileConflictPolicy.FailIfExists,
                (encoding, effective) => this.ExportImage(encoding, effective),
                (encoding, effective, token) => this.ExportImage(encoding, effective, token));
        }
    }
    public partial class ExcelDocument {
        private OfficeImageExportBatchOperation<ExcelWorkbookImageExportOptions> CreateImageExportBatchOperation(
            ExcelWorkbookImageExportOptions? options, OfficeImageExportFormat format) {
            ExcelWorkbookImageExportOptions snapshot = options?.CloneWorkbook() ?? new ExcelWorkbookImageExportOptions();
            return new OfficeImageExportBatchOperation<ExcelWorkbookImageExportOptions>(snapshot, format,
                OfficeImageExportFileConflictPolicy.FailIfExists, this.ExportImages,
                exportEach: (encoding, effective, consumer, token) =>
                    this.ExportImages(encoding, consumer, effective, token));
        }
    }
}
