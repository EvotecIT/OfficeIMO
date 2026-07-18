using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel {
    public sealed partial class ExcelRange {
        /// <summary>
        /// Creates a format-neutral visual snapshot for this range.
        /// </summary>
        public ExcelRangeVisualSnapshot CreateVisualSnapshot(ExcelImageExportOptions? options = null) =>
            ExcelRangeVisualSnapshotBuilder.Build(Sheet, Address, NormalizeOptions(options));

        /// <summary>
        /// Exports this range as a supported raster format or SVG.
        /// </summary>
        public OfficeImageExportResult ExportImage(OfficeImageExportFormat format, ExcelImageExportOptions? options = null) {
            ExcelImageExportOptions resolved = NormalizeOptions(options);
            ExcelRangeVisualSnapshot snapshot = ExcelRangeVisualSnapshotBuilder.Build(Sheet, Address, resolved);
            return ExcelRangeImageRenderer.Render(snapshot, format, resolved);
        }

        /// <summary>
        /// Renders this range to dependency-free PNG bytes.
        /// </summary>
        public byte[] ToPng(ExcelImageExportOptions? options = null) =>
            ExportImage(OfficeImageExportFormat.Png, options).Bytes;

        /// <summary>
        /// Renders this range to dependency-free SVG text.
        /// </summary>
        public string ToSvg(ExcelImageExportOptions? options = null) =>
            Encoding.UTF8.GetString(ExportImage(OfficeImageExportFormat.Svg, options).Bytes);

        /// <summary>
        /// Saves this range as a PNG file.
        /// </summary>
        public OfficeImageExportResult SaveAsPng(string path, ExcelImageExportOptions? options = null) =>
            new ExcelRangeImageExportBuilder(this, options).AsPng().Save(path);

        /// <summary>
        /// Saves this range as an SVG file.
        /// </summary>
        public OfficeImageExportResult SaveAsSvg(string path, ExcelImageExportOptions? options = null) =>
            new ExcelRangeImageExportBuilder(this, options).AsSvg().Save(path);

        /// <summary>
        /// Writes this range as PNG to a stream.
        /// </summary>
        public OfficeImageExportResult SaveAsPng(Stream stream, ExcelImageExportOptions? options = null) =>
            new ExcelRangeImageExportBuilder(this, options).AsPng().Save(stream);

        /// <summary>
        /// Writes this range as SVG to a stream.
        /// </summary>
        public OfficeImageExportResult SaveAsSvg(Stream stream, ExcelImageExportOptions? options = null) =>
            new ExcelRangeImageExportBuilder(this, options).AsSvg().Save(stream);

        /// <summary>Asynchronously saves this range as a PNG file.</summary>
        public Task<OfficeImageExportResult> SaveAsPngAsync(
            string path,
            ExcelImageExportOptions? options = null,
            CancellationToken cancellationToken = default) =>
            new ExcelRangeImageExportBuilder(this, options).AsPng().SaveAsync(path, cancellationToken);

        /// <summary>Asynchronously saves this range as an SVG file.</summary>
        public Task<OfficeImageExportResult> SaveAsSvgAsync(
            string path,
            ExcelImageExportOptions? options = null,
            CancellationToken cancellationToken = default) =>
            new ExcelRangeImageExportBuilder(this, options).AsSvg().SaveAsync(path, cancellationToken);

        /// <summary>Asynchronously writes this range as PNG to a stream.</summary>
        public Task<OfficeImageExportResult> SaveAsPngAsync(
            Stream stream,
            ExcelImageExportOptions? options = null,
            CancellationToken cancellationToken = default) =>
            new ExcelRangeImageExportBuilder(this, options).AsPng().SaveAsync(stream, cancellationToken);

        /// <summary>Asynchronously writes this range as SVG to a stream.</summary>
        public Task<OfficeImageExportResult> SaveAsSvgAsync(
            Stream stream,
            ExcelImageExportOptions? options = null,
            CancellationToken cancellationToken = default) =>
            new ExcelRangeImageExportBuilder(this, options).AsSvg().SaveAsync(stream, cancellationToken);

        private static ExcelImageExportOptions NormalizeOptions(ExcelImageExportOptions? options) {
            ExcelImageExportOptions resolved = options?.Clone() ?? new ExcelImageExportOptions();
            resolved.ConditionalFormattingDate ??= System.DateTime.Today;
            resolved.Validate();

            return resolved;
        }

    }
}
