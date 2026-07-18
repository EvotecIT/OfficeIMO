using System;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Drawing;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        /// <summary>
        /// Creates a format-neutral visual snapshot for this slide.
        /// </summary>
        public PowerPointSlideVisualSnapshot CreateVisualSnapshot(PowerPointImageExportOptions? options = null) {
            PowerPointImageExportOptions resolved = NormalizeImageExportOptions(options);
            return PowerPointSlideImageRenderer.CreateSnapshot(this, resolved);
        }

        /// <summary>
        /// Exports this slide as a supported raster format or SVG.
        /// </summary>
        public OfficeImageExportResult ExportImage(
            OfficeImageExportFormat format,
            PowerPointImageExportOptions? options = null,
            CancellationToken cancellationToken = default) {
            PowerPointImageExportOptions resolved = NormalizeImageExportOptions(options);
            return PowerPointSlideImageRenderer.Render(this, format, resolved, cancellationToken);
        }

        /// <summary>
        /// Renders this slide to dependency-free PNG bytes.
        /// </summary>
        public byte[] ToPng(PowerPointImageExportOptions? options = null) =>
            ExportImage(OfficeImageExportFormat.Png, options).Bytes;

        /// <summary>
        /// Renders this slide to dependency-free SVG text.
        /// </summary>
        public string ToSvg(PowerPointImageExportOptions? options = null) =>
            Encoding.UTF8.GetString(ExportImage(OfficeImageExportFormat.Svg, options).Bytes);

        /// <summary>
        /// Saves this slide as a PNG file.
        /// </summary>
        public OfficeImageExportResult SaveAsPng(string path, PowerPointImageExportOptions? options = null) =>
            new PowerPointSlideImageExportBuilder(this, options).AsPng().Save(path);

        /// <summary>
        /// Saves this slide as an SVG file.
        /// </summary>
        public OfficeImageExportResult SaveAsSvg(string path, PowerPointImageExportOptions? options = null) =>
            new PowerPointSlideImageExportBuilder(this, options).AsSvg().Save(path);

        /// <summary>
        /// Writes this slide as PNG to a stream.
        /// </summary>
        public OfficeImageExportResult SaveAsPng(Stream stream, PowerPointImageExportOptions? options = null) =>
            new PowerPointSlideImageExportBuilder(this, options).AsPng().Save(stream);

        /// <summary>
        /// Writes this slide as SVG to a stream.
        /// </summary>
        public OfficeImageExportResult SaveAsSvg(Stream stream, PowerPointImageExportOptions? options = null) =>
            new PowerPointSlideImageExportBuilder(this, options).AsSvg().Save(stream);

        /// <summary>Asynchronously saves this slide as a PNG file.</summary>
        public Task<OfficeImageExportResult> SaveAsPngAsync(
            string path,
            PowerPointImageExportOptions? options = null,
            CancellationToken cancellationToken = default) =>
            new PowerPointSlideImageExportBuilder(this, options).AsPng().SaveAsync(path, cancellationToken);

        /// <summary>Asynchronously saves this slide as an SVG file.</summary>
        public Task<OfficeImageExportResult> SaveAsSvgAsync(
            string path,
            PowerPointImageExportOptions? options = null,
            CancellationToken cancellationToken = default) =>
            new PowerPointSlideImageExportBuilder(this, options).AsSvg().SaveAsync(path, cancellationToken);

        /// <summary>Asynchronously writes this slide as PNG to a stream.</summary>
        public Task<OfficeImageExportResult> SaveAsPngAsync(
            Stream stream,
            PowerPointImageExportOptions? options = null,
            CancellationToken cancellationToken = default) =>
            new PowerPointSlideImageExportBuilder(this, options).AsPng().SaveAsync(stream, cancellationToken);

        /// <summary>Asynchronously writes this slide as SVG to a stream.</summary>
        public Task<OfficeImageExportResult> SaveAsSvgAsync(
            Stream stream,
            PowerPointImageExportOptions? options = null,
            CancellationToken cancellationToken = default) =>
            new PowerPointSlideImageExportBuilder(this, options).AsSvg().SaveAsync(stream, cancellationToken);

        private static PowerPointImageExportOptions NormalizeImageExportOptions(PowerPointImageExportOptions? options) {
            PowerPointImageExportOptions resolved = options?.Clone() ?? new PowerPointImageExportOptions();
            resolved.Validate();
            return resolved;
        }

    }
}
