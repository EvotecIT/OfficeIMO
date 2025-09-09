using System;
using System.Collections.Generic;
using OfficeIMO.Excel;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Formats.Png;

namespace OfficeIMO.Excel.Fluent {
    /// <summary>
    /// Fluent builder for worksheet header and footer with optional images.
    /// </summary>
    public sealed class HeaderFooterBuilder {
        private string? _hl, _hc, _hr, _fl, _fc, _fr;
        private bool _diffFirst, _diffOddEven, _alignMargins = true, _scaleWithDoc = true;
        private sealed class ImageRequest {
            public bool Header { get; set; }
            public HeaderFooterPosition Position { get; set; }
            public byte[]? Bytes { get; set; }
            public string? Url { get; set; }
            public string ContentType { get; set; } = "image/png";
            public double? W { get; set; }
            public double? H { get; set; }
        }
        private readonly List<ImageRequest> _images = new();

        /// <summary>Sets left header text.</summary>
        public HeaderFooterBuilder Left(string text) { _hl = text; return this; }
        /// <summary>Sets center header text.</summary>
        public HeaderFooterBuilder Center(string text) { _hc = text; return this; }
        /// <summary>Sets right header text.</summary>
        public HeaderFooterBuilder Right(string text) { _hr = text; return this; }

        /// <summary>Sets left footer text.</summary>
        public HeaderFooterBuilder FooterLeft(string text) { _fl = text; return this; }
        /// <summary>Sets center footer text.</summary>
        public HeaderFooterBuilder FooterCenter(string text) { _fc = text; return this; }
        /// <summary>Sets right footer text.</summary>
        public HeaderFooterBuilder FooterRight(string text) { _fr = text; return this; }

        /// <summary>Uses a different header and footer on the first page.</summary>
        public HeaderFooterBuilder DifferentFirstPage(bool value = true) { _diffFirst = value; return this; }
        /// <summary>Uses different headers/footers for odd and even pages.</summary>
        public HeaderFooterBuilder DifferentOddEven(bool value = true) { _diffOddEven = value; return this; }
        /// <summary>Aligns header/footer with page margins.</summary>
        public HeaderFooterBuilder AlignWithMargins(bool value = true) { _alignMargins = value; return this; }
        /// <summary>Scales header/footer with document scaling.</summary>
        public HeaderFooterBuilder ScaleWithDocument(bool value = true) { _scaleWithDoc = value; return this; }

        /// <summary>Adds a centered header image.</summary>
        public HeaderFooterBuilder CenterImage(byte[] bytes, string contentType = "image/png", double? widthPoints = null, double? heightPoints = null)
            => Image(true, HeaderFooterPosition.Center, bytes, contentType, widthPoints, heightPoints);
        /// <summary>Adds a left header image.</summary>
        public HeaderFooterBuilder LeftImage(byte[] bytes, string contentType = "image/png", double? widthPoints = null, double? heightPoints = null)
            => Image(true, HeaderFooterPosition.Left, bytes, contentType, widthPoints, heightPoints);
        /// <summary>Adds a right header image.</summary>
        public HeaderFooterBuilder RightImage(byte[] bytes, string contentType = "image/png", double? widthPoints = null, double? heightPoints = null)
            => Image(true, HeaderFooterPosition.Right, bytes, contentType, widthPoints, heightPoints);

        /// <summary>Adds a centered footer image.</summary>
        public HeaderFooterBuilder FooterCenterImage(byte[] bytes, string contentType = "image/png", double? widthPoints = null, double? heightPoints = null)
            => Image(false, HeaderFooterPosition.Center, bytes, contentType, widthPoints, heightPoints);
        /// <summary>Adds a left footer image.</summary>
        public HeaderFooterBuilder FooterLeftImage(byte[] bytes, string contentType = "image/png", double? widthPoints = null, double? heightPoints = null)
            => Image(false, HeaderFooterPosition.Left, bytes, contentType, widthPoints, heightPoints);
        /// <summary>Adds a right footer image.</summary>
        public HeaderFooterBuilder FooterRightImage(byte[] bytes, string contentType = "image/png", double? widthPoints = null, double? heightPoints = null)
            => Image(false, HeaderFooterPosition.Right, bytes, contentType, widthPoints, heightPoints);

        /// <summary>Adds header/footer image from raw bytes.</summary>
        private HeaderFooterBuilder Image(bool header, HeaderFooterPosition pos, byte[] bytes, string contentType, double? w, double? h) {
            if (bytes == null || bytes.Length == 0) throw new ArgumentException("Image bytes are required.", nameof(bytes));
            _images.Add(new ImageRequest { Header = header, Position = pos, Bytes = bytes, ContentType = contentType, W = w, H = h });
            return this;
        }

        /// <summary>Adds a centered header image from URL (downloaded on apply).</summary>
        public HeaderFooterBuilder CenterImageUrl(string url, double? widthPoints = null, double? heightPoints = null)
            => ImageUrl(true, HeaderFooterPosition.Center, url, widthPoints, heightPoints);
        public HeaderFooterBuilder LeftImageUrl(string url, double? widthPoints = null, double? heightPoints = null)
            => ImageUrl(true, HeaderFooterPosition.Left, url, widthPoints, heightPoints);
        public HeaderFooterBuilder RightImageUrl(string url, double? widthPoints = null, double? heightPoints = null)
            => ImageUrl(true, HeaderFooterPosition.Right, url, widthPoints, heightPoints);
        public HeaderFooterBuilder FooterCenterImageUrl(string url, double? widthPoints = null, double? heightPoints = null)
            => ImageUrl(false, HeaderFooterPosition.Center, url, widthPoints, heightPoints);
        public HeaderFooterBuilder FooterLeftImageUrl(string url, double? widthPoints = null, double? heightPoints = null)
            => ImageUrl(false, HeaderFooterPosition.Left, url, widthPoints, heightPoints);
        public HeaderFooterBuilder FooterRightImageUrl(string url, double? widthPoints = null, double? heightPoints = null)
            => ImageUrl(false, HeaderFooterPosition.Right, url, widthPoints, heightPoints);

        private HeaderFooterBuilder ImageUrl(bool header, HeaderFooterPosition pos, string url, double? w, double? h) {
            if (string.IsNullOrWhiteSpace(url)) return this;
            _images.Add(new ImageRequest { Header = header, Position = pos, Url = url, W = w, H = h });
            return this;
        }

        internal void Apply(ExcelSheet sheet) {
            sheet.SetHeaderFooter(_hl, _hc, _hr, _fl, _fc, _fr, _diffFirst, _diffOddEven, _alignMargins, _scaleWithDoc);
            foreach (var img in _images) {
                byte[]? bytes = img.Bytes;
                string contentType = img.ContentType;
                if (bytes == null && !string.IsNullOrWhiteSpace(img.Url)) {
                    if (ImageDownloader.TryFetch(img.Url!, 5, 2_000_000, out var fetched, out var _)) { bytes = fetched; }
                }
                // Normalize to PNG to match working path used by local asset case
                if (bytes != null && contentType != "image/png")
                {
                    try
                    {
                        using var image = global::SixLabors.ImageSharp.Image.Load(bytes);
                        using var ms = new System.IO.MemoryStream();
                        image.Save(ms, new PngEncoder());
                        bytes = ms.ToArray();
                        contentType = "image/png";
                    }
                    catch { /* if decode fails, keep original bytes/contentType */ }
                }
                if (bytes == null) continue;
                if (img.Header) sheet.SetHeaderImage(img.Position, bytes, contentType, img.W, img.H);
                else sheet.SetFooterImage(img.Position, bytes, contentType, img.W, img.H);
            }
        }
    }
}
