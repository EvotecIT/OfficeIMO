using OfficeIMO.Drawing;

namespace OfficeIMO.Excel {
    public sealed partial class ExcelChart {
        private OfficeImageExportOperation<ExcelImageExportOptions> CreateImageExportOperation(
            ExcelImageExportOptions? options, OfficeImageExportFormat format) {
            ExcelImageExportOptions snapshot = options?.Clone() ?? new ExcelImageExportOptions();
            return new OfficeImageExportOperation<ExcelImageExportOptions>(snapshot, format,
                snapshot.FileConflictPolicy,
                (encoding, effective) => this.ExportImage(encoding, effective),
                (encoding, effective, token) => this.ExportImage(encoding, effective, token));
        }
    }
    public sealed partial class ExcelRange {
        private OfficeImageExportOperation<ExcelImageExportOptions> CreateImageExportOperation(
            ExcelImageExportOptions? options, OfficeImageExportFormat format) {
            ExcelImageExportOptions snapshot = options?.Clone() ?? new ExcelImageExportOptions();
            return new OfficeImageExportOperation<ExcelImageExportOptions>(snapshot, format,
                snapshot.FileConflictPolicy,
                (encoding, effective) => this.ExportImage(encoding, effective),
                (encoding, effective, token) => this.ExportImage(encoding, effective, token));
        }
    }
    public partial class ExcelSheet {
        private OfficeImageExportOperation<ExcelImageExportOptions> CreateImageExportOperation(
            ExcelImageExportOptions? options, OfficeImageExportFormat format) {
            ExcelImageExportOptions snapshot = options?.Clone() ?? new ExcelImageExportOptions();
            return new OfficeImageExportOperation<ExcelImageExportOptions>(snapshot, format,
                snapshot.FileConflictPolicy,
                (encoding, effective) => this.ExportImage(encoding, effective),
                (encoding, effective, token) => this.ExportImage(encoding, effective, token));
        }
    }
    public partial class ExcelSheet {
        private OfficeImageExportBatchOperation<ExcelImageExportOptions> CreateImageExportBatchOperation(
            ExcelImageExportOptions? options, OfficeImageExportFormat format) {
            ExcelImageExportOptions snapshot = options?.Clone() ?? new ExcelImageExportOptions();
            return new OfficeImageExportBatchOperation<ExcelImageExportOptions>(snapshot, format,
                snapshot.FileConflictPolicy, this.ExportImages,
                exportEach: (encoding, effective, consumer, token) =>
                    this.ExportImages(encoding, consumer, effective, token));
        }
    }
    public partial class ExcelDocument {
        /// <summary>Saves selected worksheets with per-image metadata for in-process consumers.</summary>
        internal IReadOnlyList<OfficeImageExportResult> SaveImageResults(string folderPath, ExcelImageExportOptions options) =>
            CreateImageExportBatchOperation(options, options.Format).Save(folderPath);

        private OfficeImageExportBatchOperation<ExcelImageExportOptions> CreateImageExportBatchOperation(
            ExcelImageExportOptions? options, OfficeImageExportFormat format) {
            ExcelImageExportOptions snapshot = options?.Clone() ?? new ExcelImageExportOptions();
            return new OfficeImageExportBatchOperation<ExcelImageExportOptions>(snapshot, format,
                snapshot.FileConflictPolicy, this.ExportImages,
                exportEach: (encoding, effective, consumer, token) =>
                    this.ExportImages(encoding, consumer, effective, token));
        }
    }
}
