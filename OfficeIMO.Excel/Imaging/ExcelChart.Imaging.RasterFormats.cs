using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel {
    public sealed partial class ExcelChart {
        /// <summary>Renders this chart to dependency-free JPEG bytes.</summary>
        public byte[] ToJpeg(ExcelImageExportOptions? options = null) =>
            ExportImage(OfficeImageExportFormat.Jpeg, options).Bytes;

        /// <summary>Renders this chart to dependency-free TIFF bytes.</summary>
        public byte[] ToTiff(ExcelImageExportOptions? options = null) =>
            ExportImage(OfficeImageExportFormat.Tiff, options).Bytes;

        /// <summary>Renders this chart to dependency-free lossless WebP bytes.</summary>
        public byte[] ToWebp(ExcelImageExportOptions? options = null) =>
            ExportImage(OfficeImageExportFormat.Webp, options).Bytes;

        /// <summary>Saves this chart as JPEG.</summary>
        public OfficeImageExportResult SaveAsJpeg(string path, ExcelImageExportOptions? options = null) =>
            new ExcelChartImageExportBuilder(this, options).AsJpeg().Save(path);

        /// <summary>Saves this chart as TIFF.</summary>
        public OfficeImageExportResult SaveAsTiff(string path, ExcelImageExportOptions? options = null) =>
            new ExcelChartImageExportBuilder(this, options).AsTiff().Save(path);

        /// <summary>Saves this chart as lossless WebP.</summary>
        public OfficeImageExportResult SaveAsWebp(string path, ExcelImageExportOptions? options = null) =>
            new ExcelChartImageExportBuilder(this, options).AsWebp().Save(path);

        /// <summary>Writes this chart as JPEG.</summary>
        public OfficeImageExportResult SaveAsJpeg(Stream stream, ExcelImageExportOptions? options = null) =>
            new ExcelChartImageExportBuilder(this, options).AsJpeg().Save(stream);

        /// <summary>Writes this chart as TIFF.</summary>
        public OfficeImageExportResult SaveAsTiff(Stream stream, ExcelImageExportOptions? options = null) =>
            new ExcelChartImageExportBuilder(this, options).AsTiff().Save(stream);

        /// <summary>Writes this chart as lossless WebP.</summary>
        public OfficeImageExportResult SaveAsWebp(Stream stream, ExcelImageExportOptions? options = null) =>
            new ExcelChartImageExportBuilder(this, options).AsWebp().Save(stream);

        /// <summary>Asynchronously saves this chart as JPEG.</summary>
        public Task<OfficeImageExportResult> SaveAsJpegAsync(string path, ExcelImageExportOptions? options = null, CancellationToken cancellationToken = default) =>
            new ExcelChartImageExportBuilder(this, options).AsJpeg().SaveAsync(path, cancellationToken);

        /// <summary>Asynchronously saves this chart as TIFF.</summary>
        public Task<OfficeImageExportResult> SaveAsTiffAsync(string path, ExcelImageExportOptions? options = null, CancellationToken cancellationToken = default) =>
            new ExcelChartImageExportBuilder(this, options).AsTiff().SaveAsync(path, cancellationToken);

        /// <summary>Asynchronously saves this chart as lossless WebP.</summary>
        public Task<OfficeImageExportResult> SaveAsWebpAsync(string path, ExcelImageExportOptions? options = null, CancellationToken cancellationToken = default) =>
            new ExcelChartImageExportBuilder(this, options).AsWebp().SaveAsync(path, cancellationToken);

        /// <summary>Asynchronously writes this chart as JPEG.</summary>
        public Task<OfficeImageExportResult> SaveAsJpegAsync(Stream stream, ExcelImageExportOptions? options = null, CancellationToken cancellationToken = default) =>
            new ExcelChartImageExportBuilder(this, options).AsJpeg().SaveAsync(stream, cancellationToken);

        /// <summary>Asynchronously writes this chart as TIFF.</summary>
        public Task<OfficeImageExportResult> SaveAsTiffAsync(Stream stream, ExcelImageExportOptions? options = null, CancellationToken cancellationToken = default) =>
            new ExcelChartImageExportBuilder(this, options).AsTiff().SaveAsync(stream, cancellationToken);

        /// <summary>Asynchronously writes this chart as lossless WebP.</summary>
        public Task<OfficeImageExportResult> SaveAsWebpAsync(Stream stream, ExcelImageExportOptions? options = null, CancellationToken cancellationToken = default) =>
            new ExcelChartImageExportBuilder(this, options).AsWebp().SaveAsync(stream, cancellationToken);
    }
}
