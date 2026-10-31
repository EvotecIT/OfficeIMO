using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>Embeds native image bytes at a one-based cell and returns the attached image.</summary>
        public ExcelImage AddImage(int row, int column, byte[] imageBytes, ExcelImageOptions? options = null) {
            var settings = (options ?? new ExcelImageOptions()).Snapshot();
            return AddImageSource(row, column, CreateByteImageSource(imageBytes, settings), settings);
        }

        /// <summary>Loads an image from a file at a one-based cell and returns the attached image.</summary>
        public ExcelImage AddImage(int row, int column, string path, ExcelImageOptions? options = null) {
            var settings = (options ?? new ExcelImageOptions()).Snapshot();
            return AddImageSource(row, column, LoadImageSource(path), settings);
        }

        /// <summary>Loads an image from a stream at a one-based cell and returns the attached image.</summary>
        public ExcelImage AddImage(int row, int column, Stream stream, ExcelImageOptions? options = null) {
            var settings = (options ?? new ExcelImageOptions()).Snapshot();
            return AddImageSource(row, column, LoadImageSource(stream, settings.LeaveOpen), settings);
        }

        /// <summary>Asynchronously loads an image at a one-based cell and returns the attached image. Only source I/O is asynchronous.</summary>
        public async Task<ExcelImage> AddImageAsync(int row, int column, string path, ExcelImageOptions? options = null, CancellationToken cancellationToken = default) {
            var settings = (options ?? new ExcelImageOptions()).Snapshot();
            var source = await LoadImageSourceAsync(path, cancellationToken).ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
            return AddImageSource(row, column, source, settings);
        }

        /// <summary>Asynchronously loads an image at a one-based cell and returns the attached image. Only source I/O is asynchronous.</summary>
        public async Task<ExcelImage> AddImageAsync(int row, int column, Stream stream, ExcelImageOptions? options = null, CancellationToken cancellationToken = default) {
            var settings = (options ?? new ExcelImageOptions()).Snapshot();
            var source = await LoadImageSourceAsync(stream, settings.LeaveOpen, cancellationToken).ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
            return AddImageSource(row, column, source, settings);
        }

        /// <summary>Asynchronously loads an image at a one-based cell and returns the attached image. Only source I/O is asynchronous.</summary>
        public async Task<ExcelImage> AddImageAsync(int row, int column, Uri uri, ExcelImageOptions? options = null, CancellationToken cancellationToken = default) {
            var settings = (options ?? new ExcelImageOptions()).Snapshot();
            var source = await LoadImageSourceAsync(uri, settings.RemoteImageOptions, cancellationToken).ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
            return AddImageSource(row, column, source, settings);
        }

        /// <summary>Embeds native image bytes at an A1 cell or two-cell range and returns the attached image.</summary>
        public ExcelImage AddImage(string cellOrRange, byte[] imageBytes, ExcelImageOptions? options = null) {
            var settings = (options ?? new ExcelImageOptions()).Snapshot();
            return AddImageSource(cellOrRange, CreateByteImageSource(imageBytes, settings), settings);
        }

        /// <summary>Loads an image from a file at an A1 cell or two-cell range and returns the attached image.</summary>
        public ExcelImage AddImage(string cellOrRange, string path, ExcelImageOptions? options = null) {
            var settings = (options ?? new ExcelImageOptions()).Snapshot();
            return AddImageSource(cellOrRange, LoadImageSource(path), settings);
        }

        /// <summary>Loads an image from a stream at an A1 cell or two-cell range and returns the attached image.</summary>
        public ExcelImage AddImage(string cellOrRange, Stream stream, ExcelImageOptions? options = null) {
            var settings = (options ?? new ExcelImageOptions()).Snapshot();
            return AddImageSource(cellOrRange, LoadImageSource(stream, settings.LeaveOpen), settings);
        }

        /// <summary>Asynchronously loads an image at an A1 cell or two-cell range and returns the attached image. Only source I/O is asynchronous.</summary>
        public async Task<ExcelImage> AddImageAsync(string cellOrRange, string path, ExcelImageOptions? options = null, CancellationToken cancellationToken = default) {
            var settings = (options ?? new ExcelImageOptions()).Snapshot();
            var source = await LoadImageSourceAsync(path, cancellationToken).ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
            return AddImageSource(cellOrRange, source, settings);
        }

        /// <summary>Asynchronously loads an image at an A1 cell or two-cell range and returns the attached image. Only source I/O is asynchronous.</summary>
        public async Task<ExcelImage> AddImageAsync(string cellOrRange, Stream stream, ExcelImageOptions? options = null, CancellationToken cancellationToken = default) {
            var settings = (options ?? new ExcelImageOptions()).Snapshot();
            var source = await LoadImageSourceAsync(stream, settings.LeaveOpen, cancellationToken).ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
            return AddImageSource(cellOrRange, source, settings);
        }

        /// <summary>Asynchronously loads an image at an A1 cell or two-cell range and returns the attached image. Only source I/O is asynchronous.</summary>
        public async Task<ExcelImage> AddImageAsync(string cellOrRange, Uri uri, ExcelImageOptions? options = null, CancellationToken cancellationToken = default) {
            var settings = (options ?? new ExcelImageOptions()).Snapshot();
            var source = await LoadImageSourceAsync(uri, settings.RemoteImageOptions, cancellationToken).ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
            return AddImageSource(cellOrRange, source, settings);
        }

        /// <summary>Embeds native image bytes at a non-negative absolute position and returns the attached image.</summary>
        public ExcelImage AddImage(OfficeLength x, OfficeLength y, byte[] imageBytes, ExcelImageOptions? options = null) {
            var settings = (options ?? new ExcelImageOptions()).Snapshot();
            return AddImageSource(x, y, CreateByteImageSource(imageBytes, settings), settings);
        }

        /// <summary>Loads an image from a file at a non-negative absolute position and returns the attached image.</summary>
        public ExcelImage AddImage(OfficeLength x, OfficeLength y, string path, ExcelImageOptions? options = null) {
            var settings = (options ?? new ExcelImageOptions()).Snapshot();
            return AddImageSource(x, y, LoadImageSource(path), settings);
        }

        /// <summary>Loads an image from a stream at a non-negative absolute position and returns the attached image.</summary>
        public ExcelImage AddImage(OfficeLength x, OfficeLength y, Stream stream, ExcelImageOptions? options = null) {
            var settings = (options ?? new ExcelImageOptions()).Snapshot();
            return AddImageSource(x, y, LoadImageSource(stream, settings.LeaveOpen), settings);
        }

        /// <summary>Asynchronously loads an image at a non-negative absolute position and returns the attached image. Only source I/O is asynchronous.</summary>
        public async Task<ExcelImage> AddImageAsync(OfficeLength x, OfficeLength y, string path, ExcelImageOptions? options = null, CancellationToken cancellationToken = default) {
            var settings = (options ?? new ExcelImageOptions()).Snapshot();
            var source = await LoadImageSourceAsync(path, cancellationToken).ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
            return AddImageSource(x, y, source, settings);
        }

        /// <summary>Asynchronously loads an image at a non-negative absolute position and returns the attached image. Only source I/O is asynchronous.</summary>
        public async Task<ExcelImage> AddImageAsync(OfficeLength x, OfficeLength y, Stream stream, ExcelImageOptions? options = null, CancellationToken cancellationToken = default) {
            var settings = (options ?? new ExcelImageOptions()).Snapshot();
            var source = await LoadImageSourceAsync(stream, settings.LeaveOpen, cancellationToken).ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
            return AddImageSource(x, y, source, settings);
        }

        /// <summary>Asynchronously loads an image at a non-negative absolute position and returns the attached image. Only source I/O is asynchronous.</summary>
        public async Task<ExcelImage> AddImageAsync(OfficeLength x, OfficeLength y, Uri uri, ExcelImageOptions? options = null, CancellationToken cancellationToken = default) {
            var settings = (options ?? new ExcelImageOptions()).Snapshot();
            var source = await LoadImageSourceAsync(uri, settings.RemoteImageOptions, cancellationToken).ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
            return AddImageSource(x, y, source, settings);
        }

    }
}
