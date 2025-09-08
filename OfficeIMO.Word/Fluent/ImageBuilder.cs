using System;
using System.IO;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Linq;
using OfficeIMO;
using OfficeIMO.Word;

namespace OfficeIMO.Word.Fluent {
    /// <summary>
    /// Builder for images.
    /// </summary>
    public class ImageBuilder {
        private readonly WordFluentDocument _fluent;
        private WordImage? _image;
        private WordParagraph? _paragraph;

        internal ImageBuilder(WordFluentDocument fluent) {
            _fluent = fluent;
        }

        /// <summary>
        /// Gets the image created by the most recent operation.
        /// </summary>
        public WordImage? Image => _image;

        /// <summary>
        /// Adds an image from a file path.
        /// </summary>
        /// <param name="path">Path to the image file.</param>
        public ImageBuilder Add(string path) {
            var paragraph = _fluent.Document.AddParagraph();
            paragraph.AddImage(path);
            _image = paragraph.Image;
            _paragraph = paragraph;
            return this;
        }

        /// <summary>
        /// Adds an image from a stream.
        /// </summary>
        /// <param name="stream">Stream containing image data.</param>
        /// <param name="fileName">File name of the image.</param>
        public ImageBuilder Add(Stream stream, string fileName) {
            var paragraph = _fluent.Document.AddParagraph();
            paragraph.AddImage(stream, fileName, null, null);
            _image = paragraph.Image;
            _paragraph = paragraph;
            return this;
        }

        /// <summary>
        /// Adds an image from a byte array.
        /// </summary>
        /// <param name="bytes">Image bytes.</param>
        /// <param name="fileName">File name of the image.</param>
        public ImageBuilder Add(byte[] bytes, string fileName) {
            using var ms = new MemoryStream(bytes);
            return Add(ms, fileName);
        }

        private const int MaxImageBytes = 10 * 1024 * 1024;
        private static readonly HttpClient _httpClient = new HttpClient();

        /// <summary>
        /// Downloads and adds an image from a URL.
        /// </summary>
        /// <param name="url">Image URL.</param>
        /// <param name="cancellationToken">Token used to cancel the download operation.</param>
        /// <returns>The current <see cref="ImageBuilder"/>.</returns>
        public ImageBuilder AddFromUrl(string url, CancellationToken cancellationToken = default) {
            return AddFromUrlAsync(url, cancellationToken).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Asynchronously downloads and adds an image from a URL.
        /// </summary>
        /// <param name="url">Image URL.</param>
        /// <param name="cancellationToken">Token used to cancel the download operation.</param>
        /// <returns>The current <see cref="ImageBuilder"/>.</returns>
        public async Task<ImageBuilder> AddFromUrlAsync(string url, CancellationToken cancellationToken = default) {
            ValidateUrl(url);

            try {
                using var response = await _httpClient.GetAsync(url, HttpCompletionOption.ResponseHeadersRead, cancellationToken).ConfigureAwait(false);
                response.EnsureSuccessStatusCode();

                var mediaType = response.Content.Headers.ContentType?.MediaType;
                if (mediaType == null || !mediaType.StartsWith("image/", StringComparison.OrdinalIgnoreCase)) {
                    throw new InvalidOperationException("URL did not return an image.");
                }

                using var ms = new MemoryStream();
                using var stream = await response.Content.ReadAsStreamAsync().ConfigureAwait(false);
                var buffer = new byte[81920];
                long totalRead = 0;
                int read;
                while ((read = await stream.ReadAsync(buffer, 0, buffer.Length, cancellationToken).ConfigureAwait(false)) > 0) {
                    totalRead += read;
                    if (totalRead > MaxImageBytes) {
                        throw new InvalidOperationException($"Image exceeds maximum allowed size of {MaxImageBytes} bytes.");
                    }
                    await ms.WriteAsync(buffer, 0, read, cancellationToken).ConfigureAwait(false);
                }

                ms.Position = 0;
                string fileName = GetFileName(url);
                return Add(ms, fileName);
            } catch (OperationCanceledException) {
                throw;
            } catch (Exception ex) {
                throw new InvalidOperationException($"Failed to download image from '{url}'.", ex);
            }
        }

        private static void ValidateUrl(string url) {
            var uri = new Uri(url, UriKind.Absolute);
            if (uri.Scheme != Uri.UriSchemeHttp && uri.Scheme != Uri.UriSchemeHttps) {
                throw new ArgumentException("Only HTTP/HTTPS URLs are allowed.", nameof(url));
            }
        }

        /// <summary>
        /// Sets the size of the image in pixels.
        /// </summary>
        /// <param name="width">Image width in pixels.</param>
        /// <param name="height">Optional image height in pixels.</param>
        public ImageBuilder Size(double width, double? height = null) {
            if (_image != null) {
                _image.Width = width;
                if (height != null) {
                    _image.Height = height.Value;
                }
            }
            return this;
        }

        /// <summary>
        /// Sets the maximum allowed width of the image.
        /// </summary>
        /// <param name="width">Maximum width in pixels.</param>
        public ImageBuilder MaxWidth(double width) {
            if (_image != null) {
                if (_image.Width == null || _image.Width > width) {
                    _image.Width = width;
                }
            }
            return this;
        }

        /// <summary>
        /// Crops the image by the specified amounts in centimeters.
        /// </summary>
        /// <param name="left">Amount to crop from the left.</param>
        /// <param name="top">Amount to crop from the top.</param>
        /// <param name="right">Amount to crop from the right.</param>
        /// <param name="bottom">Amount to crop from the bottom.</param>
        public ImageBuilder Crop(double left, double top, double right, double bottom) {
            if (_image != null) {
                _image.CropLeftCentimeters = left;
                _image.CropTopCentimeters = top;
                _image.CropRightCentimeters = right;
                _image.CropBottomCentimeters = bottom;
            }
            return this;
        }

        /// <summary>
        /// Rotates the image by the specified number of degrees.
        /// </summary>
        /// <param name="degrees">Rotation angle in degrees.</param>
        public ImageBuilder Rotate(double degrees) {
            if (_image != null) {
                _image.Rotation = (int)Math.Round(degrees);
            }
            return this;
        }

        /// <summary>
        /// Applies a wrapping style to the image.
        /// </summary>
        /// <param name="wrapImage">Wrapping option.</param>
        public ImageBuilder Wrap(WrapTextImage wrapImage) {
            if (_image != null) {
                _image.WrapText = wrapImage;
            }
            return this;
        }

        /// <summary>
        /// Places the image behind text.
        /// </summary>
        public ImageBuilder BehindText() {
            return Wrap(WrapTextImage.BehindText);
        }

        /// <summary>
        /// Sets alternative text for the image.
        /// </summary>
        /// <param name="title">Optional title for the image.</param>
        /// <param name="description">Optional description for the image.</param>
        public ImageBuilder Alt(string? title = null, string? description = null) {
            if (_image != null) {
                if (title != null) {
                    _image.Title = title;
                }
                if (description != null) {
                    _image.Description = description;
                }
            }
            return this;
        }

        /// <summary>
        /// Makes the image a hyperlink pointing to the specified URL.
        /// </summary>
        /// <param name="url">Destination URL.</param>
        public ImageBuilder Link(string url) {
            if (_image != null && _paragraph != null) {
                var uri = new Uri(url);
                var paragraphElement = _paragraph._paragraph;

                HyperlinkRelationship rel;
                var headerPart = paragraphElement.Ancestors<Header>().FirstOrDefault()?.HeaderPart;
                var footerPart = paragraphElement.Ancestors<Footer>().FirstOrDefault()?.FooterPart;

                if (headerPart != null) {
                    rel = headerPart.AddHyperlinkRelationship(uri, true);
                } else if (footerPart != null) {
                    rel = footerPart.AddHyperlinkRelationship(uri, true);
                } else {
                    rel = _fluent.Document._wordprocessingDocument.MainDocumentPart!.AddHyperlinkRelationship(uri, true);
                }

                var imageRun = (Run)_image._Image.Parent!;
                imageRun.Remove();
                Hyperlink hyperlink = new Hyperlink() { Id = rel.Id };
                hyperlink.Append(imageRun);
                paragraphElement.Append(hyperlink);
            }

            return this;
        }

        /// <summary>
        /// Sets horizontal alignment for the image's paragraph.
        /// </summary>
        /// <param name="alignment">Desired horizontal alignment.</param>
        public ImageBuilder Align(HorizontalAlignment alignment) {
            var justification = alignment switch {
                HorizontalAlignment.Center => JustificationValues.Center,
                HorizontalAlignment.Right => JustificationValues.Right,
                _ => JustificationValues.Left,
            };
            _paragraph?.SetAlignment(justification);
            return this;
        }

        private static string GetFileName(string url) {
            try {
                var uri = new Uri(url);
                var fileName = Path.GetFileName(uri.LocalPath);
                return string.IsNullOrEmpty(fileName) ? "image" : fileName;
            } catch (UriFormatException) {
                return "image";
            }
        }
    }
}
