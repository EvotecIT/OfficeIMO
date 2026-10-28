using OfficeIMO.Drawing;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        private OfficeImageExportOperation<PowerPointImageExportOptions> CreateImageExportOperation(
            PowerPointImageExportOptions? options, OfficeImageExportFormat format) {
            PowerPointImageExportOptions snapshot = options?.Clone() ?? new PowerPointImageExportOptions();
            return new OfficeImageExportOperation<PowerPointImageExportOptions>(snapshot, format,
                OfficeImageExportFileConflictPolicy.FailIfExists,
                (encoding, effective) => this.ExportImage(encoding, effective),
                (encoding, effective, token) => this.ExportImage(encoding, effective, token));
        }
    }
    public sealed partial class PowerPointDocument {
        private OfficeImageExportBatchOperation<PowerPointDocumentImageExportOptions> CreateImageExportBatchOperation(
            PowerPointDocumentImageExportOptions? options, OfficeImageExportFormat format) {
            PowerPointDocumentImageExportOptions snapshot = options?.CloneDocument() ?? new PowerPointDocumentImageExportOptions();
            return new OfficeImageExportBatchOperation<PowerPointDocumentImageExportOptions>(snapshot, format,
                OfficeImageExportFileConflictPolicy.FailIfExists, this.ExportImages,
                exportEach: (encoding, effective, consumer, token) =>
                    this.ExportImages(encoding, consumer, effective, token));
        }
    }
}
