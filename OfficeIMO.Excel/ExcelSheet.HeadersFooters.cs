using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Drawing;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Snapshot of header and footer text and related flags for this worksheet.
        /// </summary>
        public sealed class HeaderFooterSnapshot {
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
            /// <summary>Left section text of the header (first page).</summary>
            public string FirstHeaderLeft { get; set; } = string.Empty;
            /// <summary>Center section text of the header (first page).</summary>
            public string FirstHeaderCenter { get; set; } = string.Empty;
            /// <summary>Right section text of the header (first page).</summary>
            public string FirstHeaderRight { get; set; } = string.Empty;
            /// <summary>Left section text of the footer (first page).</summary>
            public string FirstFooterLeft { get; set; } = string.Empty;
            /// <summary>Center section text of the footer (first page).</summary>
            public string FirstFooterCenter { get; set; } = string.Empty;
            /// <summary>Right section text of the footer (first page).</summary>
            public string FirstFooterRight { get; set; } = string.Empty;
            /// <summary>Left section text of the header (even pages).</summary>
            public string EvenHeaderLeft { get; set; } = string.Empty;
            /// <summary>Center section text of the header (even pages).</summary>
            public string EvenHeaderCenter { get; set; } = string.Empty;
            /// <summary>Right section text of the header (even pages).</summary>
            public string EvenHeaderRight { get; set; } = string.Empty;
            /// <summary>Left section text of the footer (even pages).</summary>
            public string EvenFooterLeft { get; set; } = string.Empty;
            /// <summary>Center section text of the footer (even pages).</summary>
            public string EvenFooterCenter { get; set; } = string.Empty;
            /// <summary>Right section text of the footer (even pages).</summary>
            public string EvenFooterRight { get; set; } = string.Empty;
            /// <summary>First page has different header/footer.</summary>
            public bool DifferentFirstPage { get; set; }
            /// <summary>Odd and even pages have different headers/footers.</summary>
            public bool DifferentOddEven { get; set; }
            /// <summary>True if any header section contains the picture placeholder (&amp;G).</summary>
            public bool HeaderHasPicturePlaceholder { get; set; }
            /// <summary>True if any footer section contains the picture placeholder (&amp;G).</summary>
            public bool FooterHasPicturePlaceholder { get; set; }
            /// <summary>Left section image of the header (odd pages), when available.</summary>
            public HeaderFooterImageSnapshot? HeaderLeftImage { get; set; }
            /// <summary>Center section image of the header (odd pages), when available.</summary>
            public HeaderFooterImageSnapshot? HeaderCenterImage { get; set; }
            /// <summary>Right section image of the header (odd pages), when available.</summary>
            public HeaderFooterImageSnapshot? HeaderRightImage { get; set; }
            /// <summary>Left section image of the footer (odd pages), when available.</summary>
            public HeaderFooterImageSnapshot? FooterLeftImage { get; set; }
            /// <summary>Center section image of the footer (odd pages), when available.</summary>
            public HeaderFooterImageSnapshot? FooterCenterImage { get; set; }
            /// <summary>Right section image of the footer (odd pages), when available.</summary>
            public HeaderFooterImageSnapshot? FooterRightImage { get; set; }
        }

        /// <summary>
        /// Snapshot of an Excel header/footer image.
        /// </summary>
        public sealed class HeaderFooterImageSnapshot {
            private readonly byte[] _bytes;

            internal HeaderFooterImageSnapshot(HeaderFooterPosition position, byte[] bytes, string contentType, double widthPoints, double heightPoints) {
                Position = position;
                _bytes = (byte[])bytes.Clone();
                ContentType = contentType;
                WidthPoints = widthPoints;
                HeightPoints = heightPoints;
            }

            /// <summary>Header/footer section position.</summary>
            public HeaderFooterPosition Position { get; }
            /// <summary>Image bytes.</summary>
            public byte[] Bytes => (byte[])_bytes.Clone();
            /// <summary>Image content type, such as image/png or image/jpeg.</summary>
            public string ContentType { get; }
            /// <summary>Image width in points.</summary>
            public double WidthPoints { get; }
            /// <summary>Image height in points.</summary>
            public double HeightPoints { get; }
        }

        internal static string NormalizeImageContentType(string? contentType, string parameterName) {
            if (string.IsNullOrWhiteSpace(contentType)) return OfficeImageInfo.GetMimeType(OfficeImageFormat.Png);

            if (!OfficeImageInfo.TryNormalizeImageContentType(contentType, out var normalizedContentType)) {
                throw new ArgumentException("Content type must start with 'image/'", parameterName);
            }

            return normalizedContentType;
        }

        /// <summary>
        /// Returns a snapshot of the current header and footer strings (odd pages) split into left/center/right sections,
        /// including flags and whether a picture placeholder (&amp;G) is present.
        /// </summary>
        public HeaderFooterSnapshot GetHeaderFooter() {
            var ws = WorksheetRoot;
            var hf = ws.GetFirstChild<HeaderFooter>();
            string oddHeader = hf?.OddHeader?.Text ?? string.Empty;
            string oddFooter = hf?.OddFooter?.Text ?? string.Empty;
            string firstHeader = hf?.FirstHeader?.Text ?? string.Empty;
            string firstFooter = hf?.FirstFooter?.Text ?? string.Empty;
            string evenHeader = hf?.EvenHeader?.Text ?? string.Empty;
            string evenFooter = hf?.EvenFooter?.Text ?? string.Empty;

            var (hl, hc, hr) = ParseHeaderFooterSections(oddHeader);
            var (fl, fc, fr) = ParseHeaderFooterSections(oddFooter);
            var (fhl, fhc, fhr) = ParseHeaderFooterSections(firstHeader);
            var (ffl, ffc, ffr) = ParseHeaderFooterSections(firstFooter);
            var (ehl, ehc, ehr) = ParseHeaderFooterSections(evenHeader);
            var (efl, efc, efr) = ParseHeaderFooterSections(evenFooter);

            Dictionary<string, HeaderFooterImageSnapshot> imagesByShapeId = ReadHeaderFooterImages();
            bool hasHeaderImageRel = imagesByShapeId.ContainsKey("LH") || imagesByShapeId.ContainsKey("CH") || imagesByShapeId.ContainsKey("RH");
            bool hasFooterImageRel = imagesByShapeId.ContainsKey("LF") || imagesByShapeId.ContainsKey("CF") || imagesByShapeId.ContainsKey("RF");
            try {
                var legacy = WorksheetRoot.GetFirstChild<LegacyDrawingHeaderFooter>();
                if (legacy?.Id?.Value is string relId && !string.IsNullOrEmpty(relId)) {
                    var part = _worksheetPart.GetPartById(relId);
                    if (part is VmlDrawingPart && imagesByShapeId.Count == 0) {
                        hasHeaderImageRel = true; // defensive for files where VML exists but cannot be parsed.
                        hasFooterImageRel = true;
                    }
                }
            } catch { /* ignore */ }

            return new HeaderFooterSnapshot {
                HeaderLeft = hl,
                HeaderCenter = hc,
                HeaderRight = hr,
                FooterLeft = fl,
                FooterCenter = fc,
                FooterRight = fr,
                FirstHeaderLeft = fhl,
                FirstHeaderCenter = fhc,
                FirstHeaderRight = fhr,
                FirstFooterLeft = ffl,
                FirstFooterCenter = ffc,
                FirstFooterRight = ffr,
                EvenHeaderLeft = ehl,
                EvenHeaderCenter = ehc,
                EvenHeaderRight = ehr,
                EvenFooterLeft = efl,
                EvenFooterCenter = efc,
                EvenFooterRight = efr,
                DifferentFirstPage = hf?.DifferentFirst?.Value ?? false,
                DifferentOddEven = hf?.DifferentOddEven?.Value ?? false,
                HeaderHasPicturePlaceholder = HasPicturePlaceholder(hl, hc, hr, fhl, fhc, fhr, ehl, ehc, ehr) || hasHeaderImageRel,
                FooterHasPicturePlaceholder = HasPicturePlaceholder(fl, fc, fr, ffl, ffc, ffr, efl, efc, efr) || hasFooterImageRel,
                HeaderLeftImage = imagesByShapeId.TryGetValue("LH", out var headerLeftImage) ? headerLeftImage : null,
                HeaderCenterImage = imagesByShapeId.TryGetValue("CH", out var headerCenterImage) ? headerCenterImage : null,
                HeaderRightImage = imagesByShapeId.TryGetValue("RH", out var headerRightImage) ? headerRightImage : null,
                FooterLeftImage = imagesByShapeId.TryGetValue("LF", out var footerLeftImage) ? footerLeftImage : null,
                FooterCenterImage = imagesByShapeId.TryGetValue("CF", out var footerCenterImage) ? footerCenterImage : null,
                FooterRightImage = imagesByShapeId.TryGetValue("RF", out var footerRightImage) ? footerRightImage : null
            };
        }

        private static bool HasPicturePlaceholder(params string[] values) {
            foreach (string value in values) {
                if (value.IndexOf("&G", StringComparison.Ordinal) >= 0) {
                    return true;
                }
            }

            return false;
        }

        private Dictionary<string, HeaderFooterImageSnapshot> ReadHeaderFooterImages() {
            var images = new Dictionary<string, HeaderFooterImageSnapshot>(StringComparer.OrdinalIgnoreCase);
            VmlDrawingPart? vmlPart = null;
            try {
                var legacy = WorksheetRoot.GetFirstChild<LegacyDrawingHeaderFooter>();
                if (legacy?.Id?.Value is string relId && !string.IsNullOrWhiteSpace(relId)) {
                    vmlPart = _worksheetPart.GetPartById(relId) as VmlDrawingPart;
                }
            } catch {
                return images;
            }

            if (vmlPart == null) {
                return images;
            }

            long remainingSourceImageBytes = ExcelImageExportOptions.DefaultMaximumTotalSourceImageBytes;

            XDocument vmlDocument;
            try {
                using Stream stream = vmlPart.GetStream(FileMode.Open, FileAccess.Read);
                vmlDocument = LoadVmlXDocument(stream);
            } catch {
                return images;
            }

            foreach (XElement shape in vmlDocument.Descendants().Where(element => string.Equals(element.Name.LocalName, "shape", StringComparison.OrdinalIgnoreCase))) {
                string? shapeId = shape.Attribute("id")?.Value;
                if (string.IsNullOrWhiteSpace(shapeId) || !TryGetHeaderFooterPosition(shapeId!, out bool isHeader, out HeaderFooterPosition position)) {
                    continue;
                }

                XElement? imageData = shape.Descendants().FirstOrDefault(element => string.Equals(element.Name.LocalName, "imagedata", StringComparison.OrdinalIgnoreCase));
                if (imageData == null) {
                    continue;
                }

                string? relationshipId = imageData.Attributes().FirstOrDefault(attribute =>
                    string.Equals(attribute.Name.LocalName, "id", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(attribute.Name.LocalName, "relid", StringComparison.OrdinalIgnoreCase))?.Value;
                if (string.IsNullOrWhiteSpace(relationshipId)) {
                    continue;
                }

                if (TryReadHeaderFooterImage(vmlPart, relationshipId!, shape.Attribute("style")?.Value, position, ref remainingSourceImageBytes, out HeaderFooterImageSnapshot? image)) {
                    images[shapeId!] = image!;
                }
            }

            return images;
        }

        private static bool TryReadHeaderFooterImage(
            VmlDrawingPart vmlPart,
            string relationshipId,
            string? style,
            HeaderFooterPosition position,
            ref long remainingSourceImageBytes,
            out HeaderFooterImageSnapshot? image) {
            image = null;
            ImagePart imagePart;
            try {
                if (vmlPart.GetPartById(relationshipId) is not ImagePart part) {
                    return false;
                }

                imagePart = part;
            } catch {
                return false;
            }

            using Stream source = imagePart.GetStream(FileMode.Open, FileAccess.Read);
            if (!ExcelImageExportLimits.TryReadSourceImageBytes(source, remainingSourceImageBytes, out byte[] bytes)) return false;
            remainingSourceImageBytes -= bytes.LongLength;

            double widthPoints = TryReadStylePoints(style, "width") ?? 0D;
            double heightPoints = TryReadStylePoints(style, "height") ?? 0D;
            if (widthPoints <= 0D || heightPoints <= 0D) {
                try {
                    var info = OfficeImageReader.Identify(bytes);
                    if (widthPoints <= 0D) {
                        widthPoints = info.Width * 72D / info.DpiX;
                    }

                    if (heightPoints <= 0D) {
                        heightPoints = info.Height * 72D / info.DpiY;
                    }
                } catch {
                    widthPoints = widthPoints <= 0D ? 144D : widthPoints;
                    heightPoints = heightPoints <= 0D ? 48D : heightPoints;
                }
            }

            image = new HeaderFooterImageSnapshot(position, bytes, imagePart.ContentType, widthPoints, heightPoints);
            return true;
        }

        private static double? TryReadStylePoints(string? style, string propertyName) {
            if (string.IsNullOrWhiteSpace(style)) {
                return null;
            }

            string prefix = propertyName + ":";
            foreach (string segment in style!.Split(';')) {
                string trimmed = segment.Trim();
                if (!trimmed.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)) {
                    continue;
                }

                string value = trimmed.Substring(prefix.Length).Trim();
                if (value.EndsWith("pt", StringComparison.OrdinalIgnoreCase)) {
                    value = value.Substring(0, value.Length - 2).Trim();
                }

                if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double points) && points > 0D) {
                    return points;
                }
            }

            return null;
        }

        private static bool TryGetHeaderFooterPosition(string shapeId, out bool isHeader, out HeaderFooterPosition position) {
            isHeader = false;
            position = HeaderFooterPosition.Left;
            switch (shapeId.ToUpperInvariant()) {
                case "LH":
                    isHeader = true;
                    position = HeaderFooterPosition.Left;
                    return true;
                case "CH":
                    isHeader = true;
                    position = HeaderFooterPosition.Center;
                    return true;
                case "RH":
                    isHeader = true;
                    position = HeaderFooterPosition.Right;
                    return true;
                case "LF":
                    position = HeaderFooterPosition.Left;
                    return true;
                case "CF":
                    position = HeaderFooterPosition.Center;
                    return true;
                case "RF":
                    position = HeaderFooterPosition.Right;
                    return true;
                default:
                    return false;
            }
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
            bool scaleWithDoc = true) {
            WriteLock(() => {
                var ws = WorksheetRoot;
                var hf = ws.GetFirstChild<HeaderFooter>();
                if (hf == null) {
                    hf = new HeaderFooter();
                    // Per Excel behavior, place HeaderFooter before any Drawing when possible
                    var drawing = ws.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Drawing>();
                    if (drawing != null) ws.InsertBefore(hf, drawing);
                    else {
                        var after = ws.GetFirstChild<PageSetup>();
                        if (after != null) ws.InsertAfter(hf, after); else ws.Append(hf);
                    }
                }

                hf.DifferentFirst = differentFirstPage ? true : (bool?)null;
                hf.DifferentOddEven = differentOddEven ? true : (bool?)null;
                hf.AlignWithMargins = alignWithMargins;
                hf.ScaleWithDoc = scaleWithDoc;

                var oddHeader = BuildHeaderFooterSections(headerLeft, headerCenter, headerRight);
                var oddFooter = BuildHeaderFooterSections(footerLeft, footerCenter, footerRight);

                if (oddHeader != null) hf.OddHeader = new OddHeader(oddHeader); else hf.OddHeader = null;
                if (oddFooter != null) hf.OddFooter = new OddFooter(oddFooter); else hf.OddFooter = null;

                if (differentOddEven) {
                    if (oddHeader != null) hf.EvenHeader = new EvenHeader(oddHeader);
                    if (oddFooter != null) hf.EvenFooter = new EvenFooter(oddFooter);
                } else {
                    hf.EvenHeader = null;
                    hf.EvenFooter = null;
                }
                if (differentFirstPage) {
                    if (oddHeader != null) hf.FirstHeader = new FirstHeader(oddHeader);
                    if (oddFooter != null) hf.FirstFooter = new FirstFooter(oddFooter);
                } else {
                    hf.FirstHeader = null;
                    hf.FirstFooter = null;
                }

                CleanupHeaderFooterPictureArtifacts();
                ws.Save();
            });
        }

        /// <summary>
        /// Sets a first-page header and/or footer variant for this worksheet.
        /// </summary>
        public void SetFirstPageHeaderFooter(
            string? headerLeft = null,
            string? headerCenter = null,
            string? headerRight = null,
            string? footerLeft = null,
            string? footerCenter = null,
            string? footerRight = null,
            bool enabled = true) {
            WriteLock(() => {
                HeaderFooter hf = EnsureHeaderFooter();
                hf.DifferentFirst = enabled ? true : (bool?)null;
                hf.FirstHeader = enabled ? BuildHeaderFooterSections(headerLeft, headerCenter, headerRight) is string header ? new FirstHeader(header) : null : null;
                hf.FirstFooter = enabled ? BuildHeaderFooterSections(footerLeft, footerCenter, footerRight) is string footer ? new FirstFooter(footer) : null : null;
                CleanupHeaderFooterPictureArtifacts();
                WorksheetRoot.Save();
            });
        }

        /// <summary>
        /// Sets an even-page header and/or footer variant for this worksheet.
        /// </summary>
        public void SetEvenPageHeaderFooter(
            string? headerLeft = null,
            string? headerCenter = null,
            string? headerRight = null,
            string? footerLeft = null,
            string? footerCenter = null,
            string? footerRight = null,
            bool enabled = true) {
            WriteLock(() => {
                HeaderFooter hf = EnsureHeaderFooter();
                hf.DifferentOddEven = enabled ? true : (bool?)null;
                hf.EvenHeader = enabled ? BuildHeaderFooterSections(headerLeft, headerCenter, headerRight) is string header ? new EvenHeader(header) : null : null;
                hf.EvenFooter = enabled ? BuildHeaderFooterSections(footerLeft, footerCenter, footerRight) is string footer ? new EvenFooter(footer) : null : null;
                CleanupHeaderFooterPictureArtifacts();
                WorksheetRoot.Save();
            });
        }

        private HeaderFooter EnsureHeaderFooter() {
            var ws = WorksheetRoot;
            var hf = ws.GetFirstChild<HeaderFooter>();
            if (hf != null) {
                return hf;
            }

            hf = new HeaderFooter();
            var drawing = ws.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Drawing>();
            if (drawing != null) {
                ws.InsertBefore(hf, drawing);
            } else {
                var after = ws.GetFirstChild<PageSetup>();
                if (after != null) {
                    ws.InsertAfter(hf, after);
                } else {
                    ws.Append(hf);
                }
            }

            return hf;
        }

        private static (string L, string C, string R) ParseHeaderFooterSections(string text) {
            string left = string.Empty, center = string.Empty, right = string.Empty;
            if (string.IsNullOrEmpty(text)) {
                return (left, center, right);
            }

            int i = 0;
            while (i < text.Length) {
                char ch = text[i++];
                if (ch != '&' || i >= text.Length) {
                    continue;
                }

                char section = text[i++];
                if (section != 'L' && section != 'C' && section != 'R') {
                    continue;
                }

                var builder = new StringBuilder();
                while (i < text.Length) {
                    if (text[i] == '&' && i + 1 < text.Length) {
                        char next = text[i + 1];
                        if (next == 'L' || next == 'C' || next == 'R') {
                            break;
                        }
                    }

                    builder.Append(text[i++]);
                }

                string value = builder.ToString();
                if (section == 'L') {
                    left = value;
                } else if (section == 'C') {
                    center = value;
                } else {
                    right = value;
                }
            }

            return (left, center, right);
        }

        private static string? BuildHeaderFooterSections(string? left, string? center, string? right) {
            var builder = new StringBuilder();
            if (!string.IsNullOrEmpty(left)) builder.Append("&L").Append(EscapeHeaderFooter(left));
            if (!string.IsNullOrEmpty(center)) builder.Append("&C").Append(EscapeHeaderFooter(center));
            if (!string.IsNullOrEmpty(right)) builder.Append("&R").Append(EscapeHeaderFooter(right));
            return builder.Length == 0 ? null : builder.ToString();
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
        public void SetHeaderImage(HeaderFooterPosition position, byte[] imageBytes, string contentType = "image/png", double? widthPoints = null, double? heightPoints = null) {
            if (imageBytes == null || imageBytes.Length == 0) throw new ArgumentException("Image bytes are required.", nameof(imageBytes));
            WriteLock(() => {
                var normalizedContentType = NormalizeImageContentType(contentType, nameof(contentType));
                EnsureHeaderFooterPicture(position, isHeader: true, imageBytes, normalizedContentType, widthPoints, heightPoints);
            });
        }

        /// <summary>
        /// Asynchronously downloads an image from a URL and sets it in the header at the given position.
        /// </summary>
        public async Task SetHeaderImageFromUrlAsync(HeaderFooterPosition position, string url, double? widthPoints = null,
            double? heightPoints = null, CancellationToken cancellationToken = default) {
            OfficeRemoteImage remote = await OfficeRemoteImageLoader.LoadAsync(
                url,
                cancellationToken: cancellationToken).ConfigureAwait(false);
            SetHeaderImage(position, remote.ToBytes(), remote.ContentType, widthPoints, heightPoints);
        }

        /// <summary>
        /// Adds an image to the worksheet footer at the given position. This will also ensure the footer text
        /// contains the picture placeholder (&amp;G) in the corresponding section. Subsequent calls replace any
        /// previously set header/footer images for this sheet.
        /// </summary>
        public void SetFooterImage(HeaderFooterPosition position, byte[] imageBytes, string contentType = "image/png", double? widthPoints = null, double? heightPoints = null) {
            if (imageBytes == null || imageBytes.Length == 0) throw new ArgumentException("Image bytes are required.", nameof(imageBytes));
            WriteLock(() => {
                var normalizedContentType = NormalizeImageContentType(contentType, nameof(contentType));
                EnsureHeaderFooterPicture(position, isHeader: false, imageBytes, normalizedContentType, widthPoints, heightPoints);
            });
        }

        /// <summary>
        /// Asynchronously downloads an image from a URL and sets it in the footer at the given position.
        /// </summary>
        public async Task SetFooterImageFromUrlAsync(HeaderFooterPosition position, string url, double? widthPoints = null,
            double? heightPoints = null, CancellationToken cancellationToken = default) {
            OfficeRemoteImage remote = await OfficeRemoteImageLoader.LoadAsync(
                url,
                cancellationToken: cancellationToken).ConfigureAwait(false);
            SetFooterImage(position, remote.ToBytes(), remote.ContentType, widthPoints, heightPoints);
        }

        private static string EscapeHeaderFooter(string? text) {
            if (string.IsNullOrEmpty(text)) return string.Empty;
            bool IsTokenStarter(char c) {
                // Recognize common Excel header/footer tokens following '&' to avoid escaping them
                if (c >= '0' && c <= '9') {
                    return true;
                }

                switch (c) {
                    case 'L':
                    case 'C':
                    case 'R': // section markers
                    case 'P':
                    case 'N':
                    case 'D':
                    case 'T':
                    case 'A':
                    case 'F':
                    case 'Z': // page, pages, date, time, sheet, file, path
                    case 'G': // picture placeholder
                    case 'K': // color: &Krrggbb
                    case 'B':
                    case 'I':
                    case 'E':
                    case 'U':
                    case 'S': // bold, italic, underline, strike
                    case 'X':
                    case 'Y': // superscript, subscript
                    case '[': // bracketed fields such as &[Page], &[Pages], &[Tab]
                        return true;
                }
                return false;
            }

            var t = text!;
            var sb = new StringBuilder(t.Length + 8);
            for (int i = 0; i < t.Length; i++) {
                char ch = t[i];
                if (ch == '&') {
                    if (i + 1 < t.Length) {
                        char n = t[i + 1];
                        if (n == '&') { sb.Append("&&"); i++; continue; }
                        if (n == '"') { sb.Append('&'); continue; } // font name spec: &"Name,Style"
                        if (IsTokenStarter(n)) { sb.Append('&'); continue; }
                    }
                    // literal &
                    sb.Append("&&");
                } else sb.Append(ch);
            }
            return sb.ToString();
        }

        private void EnsureHeaderFooterPicture(HeaderFooterPosition position, bool isHeader, byte[] imageBytes, string contentType, double? widthPoints, double? heightPoints) {
            var ws = WorksheetRoot;

            // 1) Ensure HeaderFooter element exists and contains &G in correct section
            var hf = ws.GetFirstChild<HeaderFooter>();
            if (hf == null) {
                hf = new HeaderFooter();
                var drawingBefore = ws.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Drawing>();
                if (drawingBefore != null) ws.InsertBefore(hf, drawingBefore);
                else { var after = ws.GetFirstChild<PageSetup>(); if (after != null) ws.InsertAfter(hf, after); else ws.Append(hf); }
            }

            // Build/update section string to include &G placeholder
            void UpsertSection(bool header, HeaderFooterPosition pos) {
                string? current = null;
                if (header)
                    current = hf.OddHeader?.Text;
                else
                    current = hf.OddFooter?.Text;

                string l = string.Empty, c = string.Empty, r = string.Empty;
                if (!string.IsNullOrEmpty(current)) {
                    // Attempt to parse existing sections to preserve other content
                    // The header/footer schema uses &L, &C, &R markers.
                    int i = 0;
                    var curr = current!;
                    while (i < curr.Length) {
                        char ch = curr[i++];
                        if (ch == '&' && i < curr.Length) {
                            char sec = curr[i++];
                            var sb = new StringBuilder();
                            while (i < curr.Length) {
                                if (curr[i] == '&' && i + 1 < curr.Length && (curr[i + 1] == 'L' || curr[i + 1] == 'C' || curr[i + 1] == 'R')) break;
                                sb.Append(curr[i++]);
                            }
                            string val = sb.ToString();
                            if (sec == 'L') l = val; else if (sec == 'C') c = val; else if (sec == 'R') r = val;
                        }
                    }
                }

                // Ensure picture placeholder &G is present for the selected section
                string EnsureG(string s) {
                    // Use IndexOf for .NET Standard 2.0 compatibility
                    return (s.IndexOf("&G", StringComparison.Ordinal) >= 0) ? s : ("&G" + s);
                }

                switch (pos) {
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
            if (legacy != null && legacy.Id != null) {
                vmlPart = (VmlDrawingPart)_worksheetPart.GetPartById(legacy.Id!);
                relId = _worksheetPart.GetIdOfPart(vmlPart);
            } else {
                vmlPart = _worksheetPart.AddNewPart<VmlDrawingPart>();
                relId = _worksheetPart.GetIdOfPart(vmlPart);
                legacy = new LegacyDrawingHeaderFooter { Id = relId };
                // Insert after HeaderFooter for proper order
                if (hf != null) ws.InsertAfter(legacy, hf); else ws.Append(legacy);
            }

            // 3) Add/replace image in the VML drawing part
            ImagePart imgPart = vmlPart.AddImagePart(ToImagePartType(contentType));

            using (var ms = new MemoryStream(imageBytes)) {
                imgPart.FeedData(ms);
            }
            string imgRelId = vmlPart.GetIdOfPart(imgPart);

            // Infer width/height if not provided
            double wPt = widthPoints ?? 0;
            double hPt = heightPoints ?? 0;
            if (wPt <= 0 || hPt <= 0) {
                try {
                    var img = OfficeImageReader.Identify(imageBytes);
                    double dpiX = img.DpiX;
                    double dpiY = img.DpiY;
                    if (wPt <= 0) wPt = img.Width * 72.0 / dpiX;
                    if (hPt <= 0) hPt = img.Height * 72.0 / dpiY;
                } catch { wPt = wPt <= 0 ? 144.0 : wPt; hPt = hPt <= 0 ? 48.0 : hPt; }
            }

            // 4) Upsert the VML shape for the selected section while preserving other header/footer images.
            string shapeId = isHeader
                ? (position == HeaderFooterPosition.Left ? "LH" : position == HeaderFooterPosition.Center ? "CH" : "RH")
                : (position == HeaderFooterPosition.Left ? "LF" : position == HeaderFooterPosition.Center ? "CF" : "RF");

            XDocument vmlDocument = LoadOrCreateHeaderFooterVmlDocument(vmlPart);
            UpsertHeaderFooterVmlShape(vmlDocument, shapeId, imgRelId, wPt, hPt);
            using (var ms = new MemoryStream()) {
                vmlDocument.Save(ms, SaveOptions.DisableFormatting);
                ms.Position = 0;
                vmlPart.FeedData(ms);
            }

            ws.Save();
        }

        internal void CleanupHeaderFooterPictureArtifacts() {
            var ws = WorksheetRoot;
            var legacy = ws.GetFirstChild<LegacyDrawingHeaderFooter>();
            if (legacy?.Id?.Value is not string legacyRelId || string.IsNullOrWhiteSpace(legacyRelId)) {
                return;
            }

            OpenXmlPart? legacyPart = null;
            try {
                legacyPart = _worksheetPart.GetPartById(legacyRelId);
            } catch {
                ws.RemoveChild(legacy);
                return;
            }

            if (HeaderFooterContainsPicturePlaceholder()) {
                return;
            }

            if (legacyPart is VmlDrawingPart vmlPart) {
                _worksheetPart.DeletePart(vmlPart);
            }

            ws.RemoveChild(legacy);
        }

        private bool HeaderFooterContainsPicturePlaceholder() {
            var hf = WorksheetRoot.GetFirstChild<HeaderFooter>();
            if (hf == null) {
                return false;
            }

            static bool HasPicture(OpenXmlLeafTextElement? element)
                => element?.Text?.IndexOf("&G", StringComparison.Ordinal) >= 0;

            return HasPicture(hf.OddHeader)
                || HasPicture(hf.OddFooter)
                || HasPicture(hf.EvenHeader)
                || HasPicture(hf.EvenFooter)
                || HasPicture(hf.FirstHeader)
                || HasPicture(hf.FirstFooter);
        }
    }
}
