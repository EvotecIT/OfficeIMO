using System.Globalization;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Formats a template marker value for a named marker format.
    /// </summary>
    public delegate string ExcelTemplateValueFormatter(object? value, IFormatProvider? provider);

    /// <summary>
    /// Controls how template binding handles markers that are not supplied by the values/model.
    /// </summary>
    public enum ExcelTemplateMissingValueBehavior {
        /// <summary>Leave the marker text unchanged.</summary>
        PreserveMarker,

        /// <summary>Replace the marker with an empty string.</summary>
        EmptyString,

        /// <summary>Throw an exception when a marker is missing.</summary>
        Throw
    }

    /// <summary>
    /// Options used when applying workbook or worksheet template markers.
    /// </summary>
    public sealed class ExcelTemplateOptions {
        /// <summary>Format provider used by built-in aliases and custom formatters.</summary>
        public IFormatProvider? FormatProvider { get; set; }

        /// <summary>Throws when a marker is not supplied by the values/model. Equivalent to <see cref="ExcelTemplateMissingValueBehavior.Throw"/>.</summary>
        public bool ThrowOnMissing { get; set; }

        /// <summary>Behavior used when a marker is not supplied by the values/model.</summary>
        public ExcelTemplateMissingValueBehavior MissingValueBehavior { get; set; }

        /// <summary>Named custom formatters, keyed by marker format such as "upper" in {{Name:upper}}.</summary>
        public IDictionary<string, ExcelTemplateValueFormatter> Formatters { get; } =
            new Dictionary<string, ExcelTemplateValueFormatter>(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Adds or replaces a named custom formatter and returns this options instance.
        /// </summary>
        public ExcelTemplateOptions AddFormatter(string name, ExcelTemplateValueFormatter formatter) {
            if (string.IsNullOrWhiteSpace(name)) throw new ArgumentNullException(nameof(name));
            Formatters[name.Trim()] = formatter ?? throw new ArgumentNullException(nameof(formatter));
            return this;
        }

        internal static ExcelTemplateOptions Create(IFormatProvider? provider, bool throwOnMissing) {
            return new ExcelTemplateOptions {
                FormatProvider = provider,
                ThrowOnMissing = throwOnMissing
            };
        }
    }

    /// <summary>
    /// Image value that can be bound to a whole-cell template marker.
    /// </summary>
    public sealed class ExcelTemplateImage {
        private readonly byte[] _bytes;

        private ExcelTemplateImage(byte[] bytes, string contentType, OfficeLength width, OfficeLength height, OfficeLength offsetX, OfficeLength offsetY, string? name, string? altText, bool lockAspectRatio) {
            if (width <= OfficeLength.Zero) throw new ArgumentOutOfRangeException(nameof(width));
            if (height <= OfficeLength.Zero) throw new ArgumentOutOfRangeException(nameof(height));
            _bytes = bytes.ToArray();
            ContentType = contentType;
            Width = width;
            Height = height;
            OffsetX = offsetX;
            OffsetY = offsetY;
            Name = name;
            AltText = altText;
            LockAspectRatio = lockAspectRatio;
        }

        /// <summary>Image bytes when the image is supplied directly.</summary>
        public byte[] Bytes => _bytes.ToArray();

        /// <summary>Image content type, such as image/png or image/jpeg.</summary>
        public string ContentType { get; }

        /// <summary>Image width as a document length.</summary>
        public OfficeLength Width { get; }

        /// <summary>Image height as a document length.</summary>
        public OfficeLength Height { get; }

        /// <summary>Horizontal length offset from the target cell.</summary>
        public OfficeLength OffsetX { get; }

        /// <summary>Vertical length offset from the target cell.</summary>
        public OfficeLength OffsetY { get; }

        /// <summary>Optional drawing name.</summary>
        public string? Name { get; }

        /// <summary>Optional alternative text description.</summary>
        public string? AltText { get; }

        /// <summary>Whether Excel should keep the picture aspect ratio locked.</summary>
        public bool LockAspectRatio { get; }

        /// <summary>
        /// Creates a template image from bytes.
        /// </summary>
        public static ExcelTemplateImage FromBytes(byte[] bytes, string contentType = "image/png", OfficeLength? width = null, OfficeLength? height = null, OfficeLength offsetX = default, OfficeLength offsetY = default, string? name = null, string? altText = null, bool lockAspectRatio = true) {
            if (bytes == null || bytes.Length == 0) throw new ArgumentException("Image bytes are required.", nameof(bytes));
            return new ExcelTemplateImage(bytes.ToArray(), NormalizeContentType(contentType), width ?? OfficeLength.Pixels(96), height ?? OfficeLength.Pixels(32), offsetX, offsetY, name, altText, lockAspectRatio);
        }

        /// <summary>
        /// Creates a template image from a stream.
        /// </summary>
        public static ExcelTemplateImage FromStream(Stream stream, string contentType = "image/png", OfficeLength? width = null, OfficeLength? height = null, OfficeLength offsetX = default, OfficeLength offsetY = default, string? name = null, string? altText = null, bool lockAspectRatio = true) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            using var buffer = new MemoryStream();
            stream.CopyTo(buffer);
            return FromBytes(buffer.ToArray(), contentType, width, height, offsetX, offsetY, name, altText, lockAspectRatio);
        }

        /// <summary>
        /// Asynchronously creates a template image from a remote URL.
        /// </summary>
        public static async Task<ExcelTemplateImage> FromUrlAsync(string url, OfficeLength? width = null, OfficeLength? height = null,
            OfficeLength offsetX = default, OfficeLength offsetY = default, string? name = null, string? altText = null,
            bool lockAspectRatio = true, CancellationToken cancellationToken = default) {
            return await FromUrlCoreAsync(url, null, width, height, offsetX, offsetY,
                name, altText, lockAspectRatio, cancellationToken).ConfigureAwait(false);
        }

        /// <summary>Creates a template image using an explicit remote-network policy.</summary>
        public static Task<ExcelTemplateImage> FromUrlAsync(string url, OfficeRemoteImageLoadOptions remoteImageOptions,
            OfficeLength? width = null, OfficeLength? height = null, OfficeLength offsetX = default, OfficeLength offsetY = default,
            string? name = null, string? altText = null, bool lockAspectRatio = true,
            CancellationToken cancellationToken = default) {
            if (remoteImageOptions == null) throw new ArgumentNullException(nameof(remoteImageOptions));
            return FromUrlCoreAsync(url, remoteImageOptions, width, height, offsetX,
                offsetY, name, altText, lockAspectRatio, cancellationToken);
        }

        private static async Task<ExcelTemplateImage> FromUrlCoreAsync(string url,
            OfficeRemoteImageLoadOptions? remoteImageOptions, OfficeLength? width, OfficeLength? height,
            OfficeLength offsetX, OfficeLength offsetY, string? name, string? altText, bool lockAspectRatio,
            CancellationToken cancellationToken) {
            OfficeRemoteImage remote = await OfficeRemoteImageLoader.LoadAsync(
                url,
                remoteImageOptions,
                cancellationToken).ConfigureAwait(false);
            return FromBytes(remote.ToBytes(), remote.ContentType, width, height, offsetX, offsetY,
                name, altText, lockAspectRatio);
        }

        internal bool TryAddToSheet(ExcelSheet sheet, int row, int column) {
            sheet.AddImage(row, column, _bytes, ContentType, Width, Height, OffsetX, OffsetY, Name, AltText, LockAspectRatio);
            return true;
        }

        private static string NormalizeContentType(string? contentType) {
            if (string.IsNullOrWhiteSpace(contentType)) {
                return OfficeImageInfo.GetMimeType(OfficeImageFormat.Png);
            }

            return OfficeImageInfo.TryNormalizeImageContentType(contentType, out string normalizedContentType)
                ? normalizedContentType
                : contentType!.Trim();
        }
    }
}
