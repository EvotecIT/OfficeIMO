using OfficeIMO.Drawing;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        private OfficeImageExportOperation<PowerPointImageExportOptions> CreateImageExportOperation(
            PowerPointImageExportOptions? options, OfficeImageExportFormat format) {
            PowerPointImageExportOptions snapshot = options?.Clone() ?? new PowerPointImageExportOptions();
            return new OfficeImageExportOperation<PowerPointImageExportOptions>(snapshot, format,
                snapshot.FileConflictPolicy,
                (encoding, effective) => this.ToImage(encoding, effective),
                (encoding, effective, token) => this.ToImage(encoding, effective, token));
        }
    }
    public sealed partial class PowerPointDocument {
        /// <summary>Saves selected slides with per-image metadata for in-process consumers.</summary>
        internal IReadOnlyList<OfficeImageExportResult> SaveImageResults(string folderPath, PowerPointImageExportOptions options) =>
            CreateImageExportBatchOperation(options, options.Format).Save(folderPath);

        private OfficeImageExportBatchOperation<PowerPointImageExportOptions> CreateImageExportBatchOperation(
            PowerPointImageExportOptions? options, OfficeImageExportFormat format) {
            PowerPointImageExportOptions snapshot = options?.Clone() ?? new PowerPointImageExportOptions();
            return new OfficeImageExportBatchOperation<PowerPointImageExportOptions>(snapshot, format,
                snapshot.FileConflictPolicy, this.ToImages,
                exportEach: (encoding, effective, consumer, token) =>
                    this.ToImages(encoding, consumer, effective, token));
        }
    }
}
