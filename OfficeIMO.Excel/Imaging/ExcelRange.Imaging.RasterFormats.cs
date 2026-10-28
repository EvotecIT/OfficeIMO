using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel {
    public sealed partial class ExcelRange {
        /// <summary>Renders this range to dependency-free JPEG bytes.</summary>
        public byte[] ToJpeg(ExcelImageExportOptions? options = null) =>
            ExportImage(OfficeImageExportFormat.Jpeg, options).Bytes;

        /// <summary>Renders this range to dependency-free TIFF bytes.</summary>
        public byte[] ToTiff(ExcelImageExportOptions? options = null) =>
            ExportImage(OfficeImageExportFormat.Tiff, options).Bytes;

        /// <summary>Renders this range to dependency-free lossless WebP bytes.</summary>
        public byte[] ToWebp(ExcelImageExportOptions? options = null) =>
            ExportImage(OfficeImageExportFormat.Webp, options).Bytes;

        /// <summary>Saves this range as JPEG.</summary>
        public OfficeImageExportResult SaveAsJpeg(string path, ExcelImageExportOptions? options = null) =>
            CreateImageExportOperation(options, OfficeImageExportFormat.Jpeg).Save(path);

        /// <summary>Saves this range as TIFF.</summary>
        public OfficeImageExportResult SaveAsTiff(string path, ExcelImageExportOptions? options = null) =>
            CreateImageExportOperation(options, OfficeImageExportFormat.Tiff).Save(path);

        /// <summary>Saves this range as lossless WebP.</summary>
        public OfficeImageExportResult SaveAsWebp(string path, ExcelImageExportOptions? options = null) =>
            CreateImageExportOperation(options, OfficeImageExportFormat.Webp).Save(path);

        /// <summary>Writes this range as JPEG.</summary>
        public OfficeImageExportResult SaveAsJpeg(Stream stream, ExcelImageExportOptions? options = null) =>
            CreateImageExportOperation(options, OfficeImageExportFormat.Jpeg).Save(stream);

        /// <summary>Writes this range as TIFF.</summary>
        public OfficeImageExportResult SaveAsTiff(Stream stream, ExcelImageExportOptions? options = null) =>
            CreateImageExportOperation(options, OfficeImageExportFormat.Tiff).Save(stream);

        /// <summary>Writes this range as lossless WebP.</summary>
        public OfficeImageExportResult SaveAsWebp(Stream stream, ExcelImageExportOptions? options = null) =>
            CreateImageExportOperation(options, OfficeImageExportFormat.Webp).Save(stream);

        /// <summary>Asynchronously saves this range as JPEG.</summary>
        public Task<OfficeImageExportResult> SaveAsJpegAsync(string path, ExcelImageExportOptions? options = null, CancellationToken cancellationToken = default) =>
            CreateImageExportOperation(options, OfficeImageExportFormat.Jpeg).SaveAsync(path, cancellationToken);

        /// <summary>Asynchronously saves this range as TIFF.</summary>
        public Task<OfficeImageExportResult> SaveAsTiffAsync(string path, ExcelImageExportOptions? options = null, CancellationToken cancellationToken = default) =>
            CreateImageExportOperation(options, OfficeImageExportFormat.Tiff).SaveAsync(path, cancellationToken);

        /// <summary>Asynchronously saves this range as lossless WebP.</summary>
        public Task<OfficeImageExportResult> SaveAsWebpAsync(string path, ExcelImageExportOptions? options = null, CancellationToken cancellationToken = default) =>
            CreateImageExportOperation(options, OfficeImageExportFormat.Webp).SaveAsync(path, cancellationToken);

        /// <summary>Asynchronously writes this range as JPEG.</summary>
        public Task<OfficeImageExportResult> SaveAsJpegAsync(Stream stream, ExcelImageExportOptions? options = null, CancellationToken cancellationToken = default) =>
            CreateImageExportOperation(options, OfficeImageExportFormat.Jpeg).SaveAsync(stream, cancellationToken);

        /// <summary>Asynchronously writes this range as TIFF.</summary>
        public Task<OfficeImageExportResult> SaveAsTiffAsync(Stream stream, ExcelImageExportOptions? options = null, CancellationToken cancellationToken = default) =>
            CreateImageExportOperation(options, OfficeImageExportFormat.Tiff).SaveAsync(stream, cancellationToken);

        /// <summary>Asynchronously writes this range as lossless WebP.</summary>
        public Task<OfficeImageExportResult> SaveAsWebpAsync(Stream stream, ExcelImageExportOptions? options = null, CancellationToken cancellationToken = default) =>
            CreateImageExportOperation(options, OfficeImageExportFormat.Webp).SaveAsync(stream, cancellationToken);
    }
}
