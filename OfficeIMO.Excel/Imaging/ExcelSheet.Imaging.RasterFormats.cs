using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>Renders the worksheet used range to dependency-free JPEG bytes.</summary>
        public byte[] ToJpeg(ExcelWorksheetImageExportOptions? options = null) =>
            ExportImage(OfficeImageExportFormat.Jpeg, options).Bytes;

        /// <summary>Renders the worksheet used range to dependency-free TIFF bytes.</summary>
        public byte[] ToTiff(ExcelWorksheetImageExportOptions? options = null) =>
            ExportImage(OfficeImageExportFormat.Tiff, options).Bytes;

        /// <summary>Renders the worksheet used range to dependency-free lossless WebP bytes.</summary>
        public byte[] ToWebp(ExcelWorksheetImageExportOptions? options = null) =>
            ExportImage(OfficeImageExportFormat.Webp, options).Bytes;

        /// <summary>Saves the worksheet used range as JPEG.</summary>
        public OfficeImageExportResult SaveAsJpeg(string path, ExcelWorksheetImageExportOptions? options = null) =>
            new ExcelWorksheetImageExportBuilder(this, options).AsJpeg().Save(path);

        /// <summary>Saves the worksheet used range as TIFF.</summary>
        public OfficeImageExportResult SaveAsTiff(string path, ExcelWorksheetImageExportOptions? options = null) =>
            new ExcelWorksheetImageExportBuilder(this, options).AsTiff().Save(path);

        /// <summary>Saves the worksheet used range as lossless WebP.</summary>
        public OfficeImageExportResult SaveAsWebp(string path, ExcelWorksheetImageExportOptions? options = null) =>
            new ExcelWorksheetImageExportBuilder(this, options).AsWebp().Save(path);

        /// <summary>Writes the worksheet used range as JPEG.</summary>
        public OfficeImageExportResult SaveAsJpeg(Stream stream, ExcelWorksheetImageExportOptions? options = null) =>
            new ExcelWorksheetImageExportBuilder(this, options).AsJpeg().Save(stream);

        /// <summary>Writes the worksheet used range as TIFF.</summary>
        public OfficeImageExportResult SaveAsTiff(Stream stream, ExcelWorksheetImageExportOptions? options = null) =>
            new ExcelWorksheetImageExportBuilder(this, options).AsTiff().Save(stream);

        /// <summary>Writes the worksheet used range as lossless WebP.</summary>
        public OfficeImageExportResult SaveAsWebp(Stream stream, ExcelWorksheetImageExportOptions? options = null) =>
            new ExcelWorksheetImageExportBuilder(this, options).AsWebp().Save(stream);

        /// <summary>Asynchronously saves the worksheet used range as JPEG.</summary>
        public Task<OfficeImageExportResult> SaveAsJpegAsync(string path, ExcelWorksheetImageExportOptions? options = null, CancellationToken cancellationToken = default) =>
            new ExcelWorksheetImageExportBuilder(this, options).AsJpeg().SaveAsync(path, cancellationToken);

        /// <summary>Asynchronously saves the worksheet used range as TIFF.</summary>
        public Task<OfficeImageExportResult> SaveAsTiffAsync(string path, ExcelWorksheetImageExportOptions? options = null, CancellationToken cancellationToken = default) =>
            new ExcelWorksheetImageExportBuilder(this, options).AsTiff().SaveAsync(path, cancellationToken);

        /// <summary>Asynchronously saves the worksheet used range as lossless WebP.</summary>
        public Task<OfficeImageExportResult> SaveAsWebpAsync(string path, ExcelWorksheetImageExportOptions? options = null, CancellationToken cancellationToken = default) =>
            new ExcelWorksheetImageExportBuilder(this, options).AsWebp().SaveAsync(path, cancellationToken);

        /// <summary>Asynchronously writes the worksheet used range as JPEG.</summary>
        public Task<OfficeImageExportResult> SaveAsJpegAsync(Stream stream, ExcelWorksheetImageExportOptions? options = null, CancellationToken cancellationToken = default) =>
            new ExcelWorksheetImageExportBuilder(this, options).AsJpeg().SaveAsync(stream, cancellationToken);

        /// <summary>Asynchronously writes the worksheet used range as TIFF.</summary>
        public Task<OfficeImageExportResult> SaveAsTiffAsync(Stream stream, ExcelWorksheetImageExportOptions? options = null, CancellationToken cancellationToken = default) =>
            new ExcelWorksheetImageExportBuilder(this, options).AsTiff().SaveAsync(stream, cancellationToken);

        /// <summary>Asynchronously writes the worksheet used range as lossless WebP.</summary>
        public Task<OfficeImageExportResult> SaveAsWebpAsync(Stream stream, ExcelWorksheetImageExportOptions? options = null, CancellationToken cancellationToken = default) =>
            new ExcelWorksheetImageExportBuilder(this, options).AsWebp().SaveAsync(stream, cancellationToken);
    }
}
