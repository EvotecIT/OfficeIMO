using OfficeIMO.Drawing;
using OfficeIMO.Core.Internal;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private readonly struct ImageSource {
            internal ImageSource(byte[] bytes, string contentType, OfficeImageInfo? info) {
                Bytes = bytes;
                ContentType = contentType;
                Info = info;
            }
            internal byte[] Bytes { get; }
            internal string ContentType { get; }
            internal OfficeImageInfo? Info { get; }
        }

        private static ImageSource CreateByteImageSource(byte[] bytes, ExcelImageOptions options) {
            if (bytes == null || bytes.Length == 0) throw new ArgumentException("Image bytes are required.", nameof(bytes));
            _ = ToImagePartType(options.ContentType);
            // Explicit native payloads remain lossless, including vector bytes and diagnostic fixtures.
            return new ImageSource(bytes, options.ContentType,
                options.ScalePercent.HasValue ? RequireImageContent(bytes, null, nameof(bytes)) : null);
        }

        private static ImageSource ValidatedImageSource(byte[] bytes, string? name, string parameterName, string? contentType = null) {
            OfficeImageInfo info = RequireImageContent(bytes, name, parameterName);
            return new ImageSource(bytes, ResolveImageContentType(contentType, info), info);
        }

        private static void ValidateImagePath(string path) {
            if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("Image path is required.", nameof(path));
            if (!File.Exists(path)) throw new FileNotFoundException($"Image file '{path}' was not found.", path);
        }

        private static ImageSource LoadImageSource(string path) {
            ValidateImagePath(path);
            return ValidatedImageSource(File.ReadAllBytes(path), path, nameof(path));
        }

        private static ImageSource LoadImageSource(Stream stream, bool leaveOpen) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            try {
                return ValidatedImageSource(OfficeStreamReader.ReadAllBytes(stream), null, nameof(stream));
            } finally {
                if (!leaveOpen) stream.Dispose();
            }
        }

        private static async Task<ImageSource> LoadImageSourceAsync(string path, CancellationToken cancellationToken) {
            cancellationToken.ThrowIfCancellationRequested();
            ValidateImagePath(path);
            using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read,
                81920, FileOptions.Asynchronous | FileOptions.SequentialScan);
            byte[] bytes = await OfficeStreamReader.ReadAllBytesAsync(stream, cancellationToken).ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
            return ValidatedImageSource(bytes, path, nameof(path));
        }

        private static async Task<ImageSource> LoadImageSourceAsync(Stream stream, bool leaveOpen, CancellationToken cancellationToken) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            try {
                byte[] bytes = await OfficeStreamReader.ReadAllBytesAsync(stream, cancellationToken).ConfigureAwait(false);
                cancellationToken.ThrowIfCancellationRequested();
                return ValidatedImageSource(bytes, null, nameof(stream));
            } finally {
                if (!leaveOpen) stream.Dispose();
            }
        }

        private static async Task<ImageSource> LoadImageSourceAsync(Uri uri, OfficeRemoteImageLoadOptions? options, CancellationToken cancellationToken) {
            if (uri == null) throw new ArgumentNullException(nameof(uri));
            if (!uri.IsAbsoluteUri || (uri.Scheme != Uri.UriSchemeHttp && uri.Scheme != Uri.UriSchemeHttps))
                throw new ArgumentException("Use an absolute HTTP or HTTPS image URI; local files use the path overload.", nameof(uri));
            OfficeRemoteImage remote = await OfficeRemoteImageLoader.LoadAsync(uri.AbsoluteUri, options, cancellationToken).ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
            return ValidatedImageSource(remote.ToBytes(), remote.FileName, nameof(uri), remote.ContentType);
        }

        private ExcelImage AddImageSource(int row, int column, ImageSource source, ExcelImageOptions options) {
            var (width, height) = ResolveSourceSize(source, options);
            ExcelImage image = AddCellImageCore(row, column, source.Bytes, source.ContentType, width, height,
                options.OffsetX, options.OffsetY, options.Name, options.Description, options.LockAspectRatio, options.Deduplicate);
            ApplyImageMetadata(image, options.Title, options.RotationDegrees);
            return image;
        }

        private ExcelImage AddImageSource(string cellOrRange, ImageSource source, ExcelImageOptions options) {
            // A colon explicitly requests a two-cell anchor, even for A1:A1; a single A1 reference requests one-cell.
            if (cellOrRange != null && cellOrRange.IndexOf(':') >= 0) {
                if (options.Width.HasValue || options.Height.HasValue || options.ScalePercent.HasValue)
                    throw new ArgumentException("A range determines image size; do not combine it with dimensions or scaling.", nameof(options));
                return AddRangeImageCore(cellOrRange, source.Bytes, source.ContentType, options.OffsetX, options.OffsetY,
                    options.EndOffsetX, options.EndOffsetY, options.Name, options.Description, options.Title,
                    options.LockAspectRatio, options.Placement, options.RotationDegrees, options.Deduplicate);
            }
            if (!ExcelA1.TryParseCellReferenceCoordinates(cellOrRange, out int row, out int column))
                throw new ArgumentException("Use an A1 cell or range such as B2 or A1:C15.", nameof(cellOrRange));
            return AddImageSource(row, column, source, options);
        }

        private ExcelImage AddImageSource(OfficeLength x, OfficeLength y, ImageSource source, ExcelImageOptions options) {
            if (x < OfficeLength.Zero) throw new ArgumentOutOfRangeException(nameof(x));
            if (y < OfficeLength.Zero) throw new ArgumentOutOfRangeException(nameof(y));
            var (width, height) = ResolveSourceSize(source, options);
            return AddNativeAbsoluteImage(x, y, source.Bytes, source.ContentType, width, height, options.Name,
                options.Description, options.Title, options.LockAspectRatio, options.RotationDegrees, options.Deduplicate);
        }

        private static (OfficeLength Width, OfficeLength Height) ResolveSourceSize(ImageSource source, ExcelImageOptions options) {
            if (source.Info != null) return ResolveImageSize(source.Info, options.Width, options.Height, options.ScalePercent);
            OfficeLength width = options.Width ?? OfficeLength.Pixels(96);
            OfficeLength height = options.Height ?? OfficeLength.Pixels(32);
            if (width <= OfficeLength.Zero) throw new ArgumentOutOfRangeException(nameof(options.Width));
            if (height <= OfficeLength.Zero) throw new ArgumentOutOfRangeException(nameof(options.Height));
            return (width, height);
        }
    }
}
