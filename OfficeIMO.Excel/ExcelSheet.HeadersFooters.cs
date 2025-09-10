using System;
using System.Globalization;
using System.IO;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using SixLabors.ImageSharp;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Snapshot of header and footer text and related flags for this worksheet.
        /// </summary>
        public sealed class HeaderFooterSnapshot
        {
            /// <summary>Left section text of the header (odd pages).</summary>
            public string HeaderLeft { get; set; } = string.Empty;
            /// <summary>Center section text of the header (odd pages).</summary>
            public string HeaderCenter { get; set; } = string.Empty;
            /// <summary>Right section text of the header (odd pages).</summary>
            public string HeaderRight { get; set; } = string.Empty;
            /// <summary>Left section text of the footer (odd pages).</summary>
            public string FooterLeft { get; set; } = string.Empty;
            /// <summary>Center section text of the footer (odd pages).</summary>
            public string FooterCenter { get; set; } = string.Empty;
            /// <summary>Right section text of the footer (odd pages).</summary>
            public string FooterRight { get; set; } = string.Empty;
            /// <summary>First page has different header/footer.</summary>
            public bool DifferentFirstPage { get; set; }
            /// <summary>Odd and even pages have different headers/footers.</summary>
            public bool DifferentOddEven { get; set; }
            /// <summary>True if any header section contains the picture placeholder (&amp;G).</summary>
            public bool HeaderHasPicturePlaceholder { get; set; }
            /// <summary>True if any footer section contains the picture placeholder (&amp;G).</summary>
            public bool FooterHasPicturePlaceholder { get; set; }
        }

        /// <summary>
        /// Returns a snapshot of the current header and footer strings (odd pages) split into left/center/right sections,
        /// including flags and whether a picture placeholder (&amp;G) is present.
        /// </summary>
        public HeaderFooterSnapshot GetHeaderFooter()
        {
            var ws = _worksheetPart.Worksheet;
            var hf = ws.GetFirstChild<HeaderFooter>();
            string oddHeader = hf?.OddHeader?.Text ?? string.Empty;
            string oddFooter = hf?.OddFooter?.Text ?? string.Empty;

            (string L, string C, string R) Parse(string text)
            {
                string l = string.Empty, c = string.Empty, r = string.Empty;
                if (string.IsNullOrEmpty(text)) return (l, c, r);
                int i = 0;
                while (i < text.Length)
                {
                    char ch = text[i++];
                    if (ch == '&' && i < text.Length)
                    {
                        char sec = text[i++];
                        if (sec == 'L' || sec == 'C' || sec == 'R')
                        {
                            var sb = new StringBuilder();
                            while (i < text.Length)
                            {
                                if (text[i] == '&' && i + 1 < text.Length)
                                {
                                    char nxt = text[i + 1];
                                    if (nxt == 'L' || nxt == 'C' || nxt == 'R') break;
                                }
                                sb.Append(text[i++]);
                            }
                            string val = sb.ToString();
                            if (sec == 'L') l = val; else if (sec == 'C') c = val; else r = val;
                        }
                    }
                }
                return (l ?? string.Empty, c ?? string.Empty, r ?? string.Empty);
            }

            var (hl, hc, hr) = Parse(oddHeader);
            var (fl, fc, fr) = Parse(oddFooter);

            // If &G is missing from the text, but a LegacyDrawingHeaderFooter part exists,
            // treat it as picture-present (defensive for files where tokens were stripped).
            bool hasHeaderImageRel = false, hasFooterImageRel = false;
            try {
                var legacy = _worksheetPart.Worksheet.GetFirstChild<LegacyDrawingHeaderFooter>();
                if (legacy?.Id?.Value is string relId && !string.IsNullOrEmpty(relId)) {
                    var part = _worksheetPart.GetPartById(relId);
                    hasHeaderImageRel = part is VmlDrawingPart; // both header/footer share the same VML part
                    hasFooterImageRel = hasHeaderImageRel;
                }
            } catch { /* ignore */ }

            return new HeaderFooterSnapshot
            {
                HeaderLeft = hl,
                HeaderCenter = hc,
                HeaderRight = hr,
                FooterLeft = fl,
                FooterCenter = fc,
                FooterRight = fr,
                DifferentFirstPage = hf?.DifferentFirst?.Value ?? false,
                DifferentOddEven = hf?.DifferentOddEven?.Value ?? false,
                HeaderHasPicturePlaceholder = (hl.IndexOf("&G", StringComparison.Ordinal) >= 0) || (hc.IndexOf("&G", StringComparison.Ordinal) >= 0) || (hr.IndexOf("&G", StringComparison.Ordinal) >= 0) || hasHeaderImageRel,
                FooterHasPicturePlaceholder = (fl.IndexOf("&G", StringComparison.Ordinal) >= 0) || (fc.IndexOf("&G", StringComparison.Ordinal) >= 0) || (fr.IndexOf("&G", StringComparison.Ordinal) >= 0) || hasFooterImageRel
            };
        }
        /// <summary>
        /// Sets the header and/or footer text for this worksheet.
        /// </summary>
        /// <param name="headerLeft">Left header text (optional).</param>
        /// <param name="headerCenter">Center header text (optional).</param>
        /// <param name="headerRight">Right header text (optional).</param>
        /// <param name="footerLeft">Left footer text (optional).</param>
        /// <param name="footerCenter">Center footer text (optional).</param>
        /// <param name="footerRight">Right footer text (optional).</param>
        /// <param name="differentFirstPage">Use a different header/footer on the first page.</param>
        /// <param name="differentOddEven">Use different headers/footers for odd and even pages.</param>
        /// <param name="alignWithMargins">Align header/footer with page margins.</param>
        /// <param name="scaleWithDoc">Scale header/footer with document scaling.</param>
        public void SetHeaderFooter(
            string? headerLeft = null,
            string? headerCenter = null,
            string? headerRight = null,
            string? footerLeft = null,
            string? footerCenter = null,
            string? footerRight = null,
            bool differentFirstPage = false,
            bool differentOddEven = false,
            bool alignWithMargins = true,
            bool scaleWithDoc = true)
        {
            WriteLock(() =>
            {
                var ws = _worksheetPart.Worksheet;
                var hf = ws.GetFirstChild<HeaderFooter>();
                if (hf == null)
                {
                    hf = new HeaderFooter();
                    // Per Excel behavior, place HeaderFooter before any Drawing when possible
                    var drawing = ws.GetFirstChild<Drawing>();
                    if (drawing != null) ws.InsertBefore(hf, drawing);
                    else
                    {
                        var after = ws.GetFirstChild<PageSetup>();
                        if (after != null) ws.InsertAfter(hf, after); else ws.Append(hf);
                    }
                }

                hf.DifferentFirst = differentFirstPage ? true : (bool?)null;
                hf.DifferentOddEven = differentOddEven ? true : (bool?)null;
                hf.AlignWithMargins = alignWithMargins ? true : (bool?)null;
                hf.ScaleWithDoc = scaleWithDoc ? true : (bool?)null;

                string? Build(string? left, string? center, string? right)
                {
                    var sb = new StringBuilder();
                    if (!string.IsNullOrEmpty(left)) sb.Append("&L").Append(EscapeHeaderFooter(left));
                    if (!string.IsNullOrEmpty(center)) sb.Append("&C").Append(EscapeHeaderFooter(center));
                    if (!string.IsNullOrEmpty(right)) sb.Append("&R").Append(EscapeHeaderFooter(right));
                    return sb.Length == 0 ? null : sb.ToString();
                }

                var oddHeader = Build(headerLeft, headerCenter, headerRight);
                var oddFooter = Build(footerLeft, footerCenter, footerRight);

                if (oddHeader != null) hf.OddHeader = new OddHeader(oddHeader); else hf.OddHeader = null;
                if (oddFooter != null) hf.OddFooter = new OddFooter(oddFooter); else hf.OddFooter = null;

                if (differentOddEven)
                {
                    if (oddHeader != null) hf.EvenHeader = new EvenHeader(oddHeader);
                    if (oddFooter != null) hf.EvenFooter = new EvenFooter(oddFooter);
                }
                if (differentFirstPage)
                {
                    if (oddHeader != null) hf.FirstHeader = new FirstHeader(oddHeader);
                    if (oddFooter != null) hf.FirstFooter = new FirstFooter(oddFooter);
                }

                ws.Save();
            });
        }

        /// <summary>
        /// Adds an image to the worksheet header at the given position. This will also ensure the header text
        /// contains the picture placeholder (&amp;G) in the corresponding section. Subsequent calls replace any
        /// previously set header/footer images for this sheet.
        /// </summary>
        /// <param name="position">Left, Center, or Right header section.</param>
        /// <param name="imageBytes">Image bytes.</param>
        /// <param name="contentType">e.g. image/png, image/jpeg. Defaults to image/png.</param>
        /// <param name="widthPoints">Optional width in points. If omitted, inferred from image size at 96 DPI.</param>
        /// <param name="heightPoints">Optional height in points. If omitted, inferred proportionally.</param>
        public void SetHeaderImage(HeaderFooterPosition position, byte[] imageBytes, string contentType = "image/png", double? widthPoints = null, double? heightPoints = null)
        {
            if (imageBytes == null || imageBytes.Length == 0) throw new ArgumentException("Image bytes are required.", nameof(imageBytes));
            WriteLock(() =>
            {
                EnsureHeaderFooterPicture(position, isHeader: true, imageBytes, contentType, widthPoints, heightPoints);
            });
        }

        /// <summary>
        /// Downloads an image from URL and sets it in the header at the given position (convenience wrapper).
        /// </summary>
        public void SetHeaderImageUrl(HeaderFooterPosition position, string url, double? widthPoints = null, double? heightPoints = null)
        {
            if (string.IsNullOrWhiteSpace(url)) return;
            if (OfficeIMO.Excel.ImageDownloader.TryFetch(url, 5, 2_000_000, out var bytes, out var _ ) && bytes != null)
                SetHeaderImage(position, bytes, "image/png", widthPoints, heightPoints);
        }

        /// <summary>
        /// Adds an image to the worksheet footer at the given position. This will also ensure the footer text
        /// contains the picture placeholder (&amp;G) in the corresponding section. Subsequent calls replace any
        /// previously set header/footer images for this sheet.
        /// </summary>
        public void SetFooterImage(HeaderFooterPosition position, byte[] imageBytes, string contentType = "image/png", double? widthPoints = null, double? heightPoints = null)
        {
            if (imageBytes == null || imageBytes.Length == 0) throw new ArgumentException("Image bytes are required.", nameof(imageBytes));
            WriteLock(() =>
            {
                EnsureHeaderFooterPicture(position, isHeader: false, imageBytes, contentType, widthPoints, heightPoints);
            });
        }

        /// <summary>
        /// Downloads an image from URL and sets it in the footer at the given position (convenience wrapper).
        /// </summary>
        public void SetFooterImageUrl(HeaderFooterPosition position, string url, double? widthPoints = null, double? heightPoints = null)
        {
            if (string.IsNullOrWhiteSpace(url)) return;
            if (OfficeIMO.Excel.ImageDownloader.TryFetch(url, 5, 2_000_000, out var bytes, out var _ ) && bytes != null)
                SetFooterImage(position, bytes, "image/png", widthPoints, heightPoints);
        }

        private static string EscapeHeaderFooter(string? text)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;
            bool IsTokenStarter(char c)
            {
                // Recognize common Excel header/footer tokens following '&' to avoid escaping them
                switch (c)
                {
                    case 'L': case 'C': case 'R': // section markers
                    case 'P': case 'N': case 'D': case 'T': case 'A': case 'F': // page, pages, date, time, sheet, file
                    case 'G': // picture placeholder
                    case 'K': // color: &Krrggbb
                    case 'B': case 'I': case 'U': case 'S': // bold, italic, underline, strike
                        return true;
                }
                return false;
            }

            var t = text!;
            var sb = new StringBuilder(t.Length + 8);
            for (int i = 0; i < t.Length; i++)
            {
                char ch = t[i];
                if (ch == '&')
                {
                    if (i + 1 < t.Length)
                    {
                        char n = t[i + 1];
                        if (n == '&') { sb.Append("&&"); i++; continue; }
                        if (n == '"') { sb.Append('&'); continue; } // font name spec: &"Name,Style"
                        if (IsTokenStarter(n)) { sb.Append('&'); continue; }
                    }
                    // literal &
                    sb.Append("&&");
                }
                else sb.Append(ch);
            }
            return sb.ToString();
        }

        private void EnsureHeaderFooterPicture(HeaderFooterPosition position, bool isHeader, byte[] imageBytes, string contentType, double? widthPoints, double? heightPoints)
        {
            var ws = _worksheetPart.Worksheet;

            // 1) Ensure HeaderFooter element exists and contains &G in correct section
            var hf = ws.GetFirstChild<HeaderFooter>();
            if (hf == null)
            {
                hf = new HeaderFooter();
                var drawingBefore = ws.GetFirstChild<Drawing>();
                if (drawingBefore != null) ws.InsertBefore(hf, drawingBefore);
                else { var after = ws.GetFirstChild<PageSetup>(); if (after != null) ws.InsertAfter(hf, after); else ws.Append(hf); }
            }

            // Build/update section string to include &G placeholder
            void UpsertSection(bool header, HeaderFooterPosition pos)
            {
                string? current = null;
                if (header)
                    current = hf.OddHeader?.Text;
                else
                    current = hf.OddFooter?.Text;

                string l = string.Empty, c = string.Empty, r = string.Empty;
                if (!string.IsNullOrEmpty(current))
                {
                    // Attempt to parse existing sections to preserve other content
                    // The header/footer schema uses &L, &C, &R markers.
                    int i = 0;
                    var curr = current!;
                    while (i < curr.Length)
                    {
                        char ch = curr[i++];
                        if (ch == '&' && i < curr.Length)
                        {
                            char sec = curr[i++];
                            var sb = new StringBuilder();
                            while (i < curr.Length)
                            {
                                if (curr[i] == '&' && i + 1 < curr.Length && (curr[i + 1] == 'L' || curr[i + 1] == 'C' || curr[i + 1] == 'R')) break;
                                sb.Append(curr[i++]);
                            }
                            string val = sb.ToString();
                            if (sec == 'L') l = val; else if (sec == 'C') c = val; else if (sec == 'R') r = val;
                        }
                    }
                }

                // Ensure picture placeholder &G is present for the selected section
                string EnsureG(string s)
                {
                    // Use IndexOf for .NET Standard 2.0 compatibility
                    return (s.IndexOf("&G", StringComparison.Ordinal) >= 0) ? s : ("&G" + s);
                }

                switch (pos)
                {
                    case HeaderFooterPosition.Left: l = EnsureG(l); break;
                    case HeaderFooterPosition.Center: c = EnsureG(c); break;
                    case HeaderFooterPosition.Right: r = EnsureG(r); break;
                }

                string rebuilt = new StringBuilder()
                    .Append(string.IsNullOrEmpty(l) ? string.Empty : "&L" + l)
                    .Append(string.IsNullOrEmpty(c) ? string.Empty : "&C" + c)
                    .Append(string.IsNullOrEmpty(r) ? string.Empty : "&R" + r)
                    .ToString();

                if (header)
                    hf.OddHeader = string.IsNullOrEmpty(rebuilt) ? null : new OddHeader(rebuilt);
                else
                    hf.OddFooter = string.IsNullOrEmpty(rebuilt) ? null : new OddFooter(rebuilt);
            }

            UpsertSection(isHeader, position);

            // 2) Create or reuse VML drawing part for header/footer images
            VmlDrawingPart vmlPart;
            string relId;
            var legacy = ws.GetFirstChild<LegacyDrawingHeaderFooter>();
            if (legacy != null && legacy.Id != null)
            {
                vmlPart = (VmlDrawingPart)_worksheetPart.GetPartById(legacy.Id!);
                relId = _worksheetPart.GetIdOfPart(vmlPart);
            }
            else
            {
                vmlPart = _worksheetPart.AddNewPart<VmlDrawingPart>();
                relId = _worksheetPart.GetIdOfPart(vmlPart);
                legacy = new LegacyDrawingHeaderFooter { Id = relId };
                // Insert after HeaderFooter for proper order
                if (hf != null) ws.InsertAfter(legacy, hf); else ws.Append(legacy);
            }

            // 3) Add/replace image in the VML drawing part
            ImagePart imgPart;
            if (contentType.Equals("image/png", StringComparison.OrdinalIgnoreCase))
                imgPart = vmlPart.AddImagePart(ImagePartType.Png);
            else if (contentType.Equals("image/jpeg", StringComparison.OrdinalIgnoreCase) || contentType.Equals("image/jpg", StringComparison.OrdinalIgnoreCase))
                imgPart = vmlPart.AddImagePart(ImagePartType.Jpeg);
            else if (contentType.Equals("image/gif", StringComparison.OrdinalIgnoreCase))
                imgPart = vmlPart.AddImagePart(ImagePartType.Gif);
            else if (contentType.Equals("image/bmp", StringComparison.OrdinalIgnoreCase))
                imgPart = vmlPart.AddImagePart(ImagePartType.Bmp);
            else
                imgPart = vmlPart.AddImagePart(ImagePartType.Png);

            using (var ms = new MemoryStream(imageBytes))
            {
                imgPart.FeedData(ms);
            }
            string imgRelId = vmlPart.GetIdOfPart(imgPart);

            // Infer width/height if not provided
            double wPt = widthPoints ?? 0;
            double hPt = heightPoints ?? 0;
            if (wPt <= 0 || hPt <= 0)
            {
                try
                {
                    using var img = Image.Load(imageBytes);
                    double dpiX = 96.0, dpiY = 96.0; // ImageSharp stores resolution separately; defaults vary
                    var md = img.Metadata.ResolutionUnits;
                    if (img.Metadata.HorizontalResolution > 0) dpiX = img.Metadata.HorizontalResolution;
                    if (img.Metadata.VerticalResolution > 0) dpiY = img.Metadata.VerticalResolution;
                    if (wPt <= 0) wPt = img.Width * 72.0 / dpiX;
                    if (hPt <= 0) hPt = img.Height * 72.0 / dpiY;
                }
                catch { wPt = wPt <= 0 ? 144.0 : wPt; hPt = hPt <= 0 ? 48.0 : hPt; }
            }

            // 4) Write minimal VML markup with a shape for the selected section
            string shapeId = isHeader
                ? (position == HeaderFooterPosition.Left ? "LH" : position == HeaderFooterPosition.Center ? "CH" : "RH")
                : (position == HeaderFooterPosition.Left ? "LF" : position == HeaderFooterPosition.Center ? "CF" : "RF");

            string vml = $@"<?xml version=""1.0"" encoding=""UTF-8""?>
<xml xmlns:v=""urn:schemas-microsoft-com:vml"" xmlns:o=""urn:schemas-microsoft-com:office:office"" xmlns:x=""urn:schemas-microsoft-com:office:excel"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"">
  <o:shapelayout v:ext=""edit""><o:idmap v:ext=""edit"" data=""1""/></o:shapelayout>
  <v:shapetype id=""_x0000_t75"" coordsize=""21600,21600"" o:spt=""75"" o:preferrelative=""t"" path=""m@4@5l@4@11@9@11@9@5xe"" filled=""f"" stroked=""f"">
    <v:stroke joinstyle=""miter""/>
    <v:formulas>
      <v:f eqn=""if lineDrawn pixelLineWidth 0""/>
      <v:f eqn=""sum @0 1 0""/>
      <v:f eqn=""sum 0 0 @1""/>
      <v:f eqn=""prod @2 1 2""/>
      <v:f eqn=""prod @3 21600 pixelWidth""/>
      <v:f eqn=""prod @3 21600 pixelHeight""/>
      <v:f eqn=""sum @0 0 1""/>
      <v:f eqn=""prod @6 1 2""/>
      <v:f eqn=""prod @7 21600 pixelWidth""/>
      <v:f eqn=""sum @8 21600 0""/>
      <v:f eqn=""prod @7 21600 pixelHeight""/>
      <v:f eqn=""sum @10 21600 0""/>
    </v:formulas>
    <v:path o:extrusionok=""f"" gradientshapeok=""t"" o:connecttype=""rect""/>
    <o:lock v:ext=""edit"" aspectratio=""t""/>
  </v:shapetype>
  <v:shape id=""{shapeId}"" o:spid=""_x0000_s1025"" type=""#_x0000_t75"" style=""position:absolute;margin-left:0;margin-top:0;width:{wPt.ToString(CultureInfo.InvariantCulture)}pt;height:{hPt.ToString(CultureInfo.InvariantCulture)}pt;z-index:1"">
    <v:imagedata r:id=""{imgRelId}"" o:relid=""{imgRelId}"" o:title=""""/>
    </v:shape>
</xml>";

            using (var ms = new MemoryStream(Encoding.UTF8.GetBytes(vml)))
            {
                vmlPart.FeedData(ms);
            }

            ws.Save();
        }
    }
}
