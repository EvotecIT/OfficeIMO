using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Drawing;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        /// <summary>Renders this slide to dependency-free JPEG bytes.</summary>
        public byte[] ToJpeg(PowerPointImageExportOptions? options = null) =>
            ExportImage(OfficeImageExportFormat.Jpeg, options).Bytes;

        /// <summary>Renders this slide to dependency-free TIFF bytes.</summary>
        public byte[] ToTiff(PowerPointImageExportOptions? options = null) =>
            ExportImage(OfficeImageExportFormat.Tiff, options).Bytes;

        /// <summary>Renders this slide to dependency-free lossless WebP bytes.</summary>
        public byte[] ToWebp(PowerPointImageExportOptions? options = null) =>
            ExportImage(OfficeImageExportFormat.Webp, options).Bytes;

        /// <summary>Saves this slide as JPEG.</summary>
        public OfficeImageExportResult SaveAsJpeg(string path, PowerPointImageExportOptions? options = null) =>
            new PowerPointSlideImageExportBuilder(this, options).AsJpeg().Save(path);

        /// <summary>Saves this slide as TIFF.</summary>
        public OfficeImageExportResult SaveAsTiff(string path, PowerPointImageExportOptions? options = null) =>
            new PowerPointSlideImageExportBuilder(this, options).AsTiff().Save(path);

        /// <summary>Saves this slide as lossless WebP.</summary>
        public OfficeImageExportResult SaveAsWebp(string path, PowerPointImageExportOptions? options = null) =>
            new PowerPointSlideImageExportBuilder(this, options).AsWebp().Save(path);

        /// <summary>Writes this slide as JPEG.</summary>
        public OfficeImageExportResult SaveAsJpeg(Stream stream, PowerPointImageExportOptions? options = null) =>
            new PowerPointSlideImageExportBuilder(this, options).AsJpeg().Save(stream);

        /// <summary>Writes this slide as TIFF.</summary>
        public OfficeImageExportResult SaveAsTiff(Stream stream, PowerPointImageExportOptions? options = null) =>
            new PowerPointSlideImageExportBuilder(this, options).AsTiff().Save(stream);

        /// <summary>Writes this slide as lossless WebP.</summary>
        public OfficeImageExportResult SaveAsWebp(Stream stream, PowerPointImageExportOptions? options = null) =>
            new PowerPointSlideImageExportBuilder(this, options).AsWebp().Save(stream);

        /// <summary>Asynchronously saves this slide as JPEG.</summary>
        public Task<OfficeImageExportResult> SaveAsJpegAsync(string path, PowerPointImageExportOptions? options = null, CancellationToken cancellationToken = default) =>
            new PowerPointSlideImageExportBuilder(this, options).AsJpeg().SaveAsync(path, cancellationToken);

        /// <summary>Asynchronously saves this slide as TIFF.</summary>
        public Task<OfficeImageExportResult> SaveAsTiffAsync(string path, PowerPointImageExportOptions? options = null, CancellationToken cancellationToken = default) =>
            new PowerPointSlideImageExportBuilder(this, options).AsTiff().SaveAsync(path, cancellationToken);

        /// <summary>Asynchronously saves this slide as lossless WebP.</summary>
        public Task<OfficeImageExportResult> SaveAsWebpAsync(string path, PowerPointImageExportOptions? options = null, CancellationToken cancellationToken = default) =>
            new PowerPointSlideImageExportBuilder(this, options).AsWebp().SaveAsync(path, cancellationToken);

        /// <summary>Asynchronously writes this slide as JPEG.</summary>
        public Task<OfficeImageExportResult> SaveAsJpegAsync(Stream stream, PowerPointImageExportOptions? options = null, CancellationToken cancellationToken = default) =>
            new PowerPointSlideImageExportBuilder(this, options).AsJpeg().SaveAsync(stream, cancellationToken);

        /// <summary>Asynchronously writes this slide as TIFF.</summary>
        public Task<OfficeImageExportResult> SaveAsTiffAsync(Stream stream, PowerPointImageExportOptions? options = null, CancellationToken cancellationToken = default) =>
            new PowerPointSlideImageExportBuilder(this, options).AsTiff().SaveAsync(stream, cancellationToken);

        /// <summary>Asynchronously writes this slide as lossless WebP.</summary>
        public Task<OfficeImageExportResult> SaveAsWebpAsync(Stream stream, PowerPointImageExportOptions? options = null, CancellationToken cancellationToken = default) =>
            new PowerPointSlideImageExportBuilder(this, options).AsWebp().SaveAsync(stream, cancellationToken);
    }
}
