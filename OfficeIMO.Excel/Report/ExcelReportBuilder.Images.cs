using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Excel {
    public sealed partial class ExcelReportBuilder {
        /// <summary>Adds an image through the worksheet and returns its attached image.</summary>
        public ExcelImage AddImage(int row, int column, byte[] imageBytes, ExcelImageOptions? options = null)
            => Sheet.AddImage(row, column, imageBytes, options);

        /// <summary>Adds an image through the worksheet and returns its attached image.</summary>
        public ExcelImage AddImage(int row, int column, string path, ExcelImageOptions? options = null)
            => Sheet.AddImage(row, column, path, options);

        /// <summary>Adds an image through the worksheet and returns its attached image.</summary>
        public ExcelImage AddImage(int row, int column, Stream stream, ExcelImageOptions? options = null)
            => Sheet.AddImage(row, column, stream, options);

        /// <summary>Loads an image asynchronously through the worksheet and returns its attached image.</summary>
        public Task<ExcelImage> AddImageAsync(int row, int column, string path, ExcelImageOptions? options = null, CancellationToken cancellationToken = default)
            => Sheet.AddImageAsync(row, column, path, options, cancellationToken);

        /// <summary>Loads an image asynchronously through the worksheet and returns its attached image.</summary>
        public Task<ExcelImage> AddImageAsync(int row, int column, Stream stream, ExcelImageOptions? options = null, CancellationToken cancellationToken = default)
            => Sheet.AddImageAsync(row, column, stream, options, cancellationToken);

        /// <summary>Loads an image asynchronously through the worksheet and returns its attached image.</summary>
        public Task<ExcelImage> AddImageAsync(int row, int column, Uri uri, ExcelImageOptions? options = null, CancellationToken cancellationToken = default)
            => Sheet.AddImageAsync(row, column, uri, options, cancellationToken);

    }
}
