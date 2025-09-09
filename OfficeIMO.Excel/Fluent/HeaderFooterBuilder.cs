using System;
using System.Collections.Generic;
using OfficeIMO.Excel.Enums;

namespace OfficeIMO.Excel.Fluent
{
    /// <summary>
    /// Fluent builder for worksheet header and footer with optional images.
    /// </summary>
    public sealed class HeaderFooterBuilder
    {
        private string? _hl, _hc, _hr, _fl, _fc, _fr;
        private bool _diffFirst, _diffOddEven, _alignMargins = true, _scaleWithDoc = true;
        private readonly List<(bool header, HeaderFooterPosition pos, byte[] bytes, string contentType, double? w, double? h)> _images = new();

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

        private HeaderFooterBuilder Image(bool header, HeaderFooterPosition pos, byte[] bytes, string contentType, double? w, double? h)
        {
            if (bytes == null || bytes.Length == 0) throw new ArgumentException("Image bytes are required.", nameof(bytes));
            _images.Add((header, pos, bytes, contentType, w, h));
            return this;
        }

        internal void Apply(ExcelSheet sheet)
        {
            sheet.SetHeaderFooter(_hl, _hc, _hr, _fl, _fc, _fr, _diffFirst, _diffOddEven, _alignMargins, _scaleWithDoc);
            foreach (var img in _images)
            {
                if (img.header)
                    sheet.SetHeaderImage(img.pos, img.bytes, img.contentType, img.w, img.h);
                else
                    sheet.SetFooterImage(img.pos, img.bytes, img.contentType, img.w, img.h);
            }
        }
    }
}
