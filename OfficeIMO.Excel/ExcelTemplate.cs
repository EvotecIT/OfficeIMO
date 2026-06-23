using System.Globalization;
using System.Reflection;
using System.Text.RegularExpressions;
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
        private ExcelTemplateImage(byte[]? bytes, string? url, string contentType, int widthPixels, int heightPixels, int offsetXPixels, int offsetYPixels, string? name, string? altText, bool lockAspectRatio) {
            Bytes = bytes;
            Url = url;
            ContentType = contentType;
            WidthPixels = widthPixels;
            HeightPixels = heightPixels;
            OffsetXPixels = offsetXPixels;
            OffsetYPixels = offsetYPixels;
            Name = name;
            AltText = altText;
            LockAspectRatio = lockAspectRatio;
        }

        /// <summary>Image bytes when the image is supplied directly.</summary>
        public byte[]? Bytes { get; }

        /// <summary>Remote image URL when the image should be downloaded during binding.</summary>
        public string? Url { get; }

        /// <summary>Image content type, such as image/png or image/jpeg.</summary>
        public string ContentType { get; }

        /// <summary>Image width in pixels.</summary>
        public int WidthPixels { get; }

        /// <summary>Image height in pixels.</summary>
        public int HeightPixels { get; }

        /// <summary>Horizontal pixel offset from the target cell.</summary>
        public int OffsetXPixels { get; }

        /// <summary>Vertical pixel offset from the target cell.</summary>
        public int OffsetYPixels { get; }

        /// <summary>Optional drawing name.</summary>
        public string? Name { get; }

        /// <summary>Optional alternative text description.</summary>
        public string? AltText { get; }

        /// <summary>Whether Excel should keep the picture aspect ratio locked.</summary>
        public bool LockAspectRatio { get; }

        /// <summary>
        /// Creates a template image from bytes.
        /// </summary>
        public static ExcelTemplateImage FromBytes(byte[] bytes, string contentType = "image/png", int widthPixels = 96, int heightPixels = 32, int offsetXPixels = 0, int offsetYPixels = 0, string? name = null, string? altText = null, bool lockAspectRatio = true) {
            if (bytes == null || bytes.Length == 0) throw new ArgumentException("Image bytes are required.", nameof(bytes));
            return new ExcelTemplateImage(bytes.ToArray(), null, NormalizeContentType(contentType), widthPixels, heightPixels, offsetXPixels, offsetYPixels, name, altText, lockAspectRatio);
        }

        /// <summary>
        /// Creates a template image from a stream.
        /// </summary>
        public static ExcelTemplateImage FromStream(Stream stream, string contentType = "image/png", int widthPixels = 96, int heightPixels = 32, int offsetXPixels = 0, int offsetYPixels = 0, string? name = null, string? altText = null, bool lockAspectRatio = true) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            using var buffer = new MemoryStream();
            stream.CopyTo(buffer);
            return FromBytes(buffer.ToArray(), contentType, widthPixels, heightPixels, offsetXPixels, offsetYPixels, name, altText, lockAspectRatio);
        }

        /// <summary>
        /// Creates a template image from a remote URL. The image is downloaded when the template is applied.
        /// </summary>
        public static ExcelTemplateImage FromUrl(string url, int widthPixels = 96, int heightPixels = 32, int offsetXPixels = 0, int offsetYPixels = 0, string? name = null, string? altText = null, bool lockAspectRatio = true) {
            if (string.IsNullOrWhiteSpace(url)) throw new ArgumentNullException(nameof(url));
            return new ExcelTemplateImage(null, url.Trim(), OfficeImageInfo.GetMimeType(OfficeImageFormat.Png), widthPixels, heightPixels, offsetXPixels, offsetYPixels, name, altText, lockAspectRatio);
        }

        internal bool TryAddToSheet(ExcelSheet sheet, int row, int column) {
            if (Bytes != null) {
                sheet.AddImage(row, column, Bytes, ContentType, WidthPixels, HeightPixels, OffsetXPixels, OffsetYPixels, Name, AltText, LockAspectRatio);
                return true;
            }

            if (!string.IsNullOrWhiteSpace(Url)
                && ImageDownloader.TryFetch(Url!, timeoutSeconds: 5, maxBytes: 2_000_000, out var bytes, out var contentType)
                && bytes != null) {
                sheet.AddImage(row, column, bytes, string.IsNullOrWhiteSpace(contentType) ? ContentType : contentType!, WidthPixels, HeightPixels, OffsetXPixels, OffsetYPixels, Name, AltText, LockAspectRatio);
                return true;
            }

            return false;
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
