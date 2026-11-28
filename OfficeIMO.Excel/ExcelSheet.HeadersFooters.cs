using OfficeIMO.Excel.OpenXml;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Drawing;
using System.Globalization;
using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Snapshot of header and footer text and related flags for this worksheet.
        /// </summary>
        internal sealed class ExcelHeaderFooterInfo {
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
            public ExcelHeaderFooterImageInfo? HeaderLeftImage { get; set; }
            /// <summary>Center section image of the header (odd pages), when available.</summary>
            public ExcelHeaderFooterImageInfo? HeaderCenterImage { get; set; }
            /// <summary>Right section image of the header (odd pages), when available.</summary>
            public ExcelHeaderFooterImageInfo? HeaderRightImage { get; set; }
            /// <summary>Left section image of the footer (odd pages), when available.</summary>
            public ExcelHeaderFooterImageInfo? FooterLeftImage { get; set; }
            /// <summary>Center section image of the footer (odd pages), when available.</summary>
            public ExcelHeaderFooterImageInfo? FooterCenterImage { get; set; }
            /// <summary>Right section image of the footer (odd pages), when available.</summary>
            public ExcelHeaderFooterImageInfo? FooterRightImage { get; set; }

            internal long SourceImageByteCount =>
                (HeaderLeftImage?.SourceByteCount ?? 0L) +
                (HeaderCenterImage?.SourceByteCount ?? 0L) +
                (HeaderRightImage?.SourceByteCount ?? 0L) +
                (FooterLeftImage?.SourceByteCount ?? 0L) +
                (FooterCenterImage?.SourceByteCount ?? 0L) +
                (FooterRightImage?.SourceByteCount ?? 0L);
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
        internal ExcelHeaderFooterInfo GetHeaderFooter() =>
            GetHeaderFooter(ExcelImageExportOptions.DefaultMaximumTotalSourceImageBytes);

        internal ExcelHeaderFooterInfo GetHeaderFooter(long maximumTotalSourceImageBytes) {
            ArgumentOutOfRangeException.ThrowIfLessThanOrEqual(maximumTotalSourceImageBytes, 0L);

            var metadata = ReadWorksheetObject(worksheet => {
                OfficeXmlNode footer = worksheet.Child(WorksheetMarkup, "headerFooter");
                return (
                    OddHeader: GetHeaderFooterText(footer, ExcelHeaderFooterKind.OddHeader),
                    OddFooter: GetHeaderFooterText(footer, ExcelHeaderFooterKind.OddFooter),
                    FirstHeader: GetHeaderFooterText(footer, ExcelHeaderFooterKind.FirstHeader),
                    FirstFooter: GetHeaderFooterText(footer, ExcelHeaderFooterKind.FirstFooter),
                    EvenHeader: GetHeaderFooterText(footer, ExcelHeaderFooterKind.EvenHeader),
                    EvenFooter: GetHeaderFooterText(footer, ExcelHeaderFooterKind.EvenFooter),
                    First: footer.BooleanAttribute("differentFirst") ?? false,
                    OddEven: footer.BooleanAttribute("differentOddEven") ?? false,
                    HasPicture: !worksheet.Child(WorksheetMarkup, "legacyDrawingHF").IsNone);
            });
            var (hl, hc, hr) = ParseHeaderFooterSections(metadata.OddHeader ?? string.Empty);
            var (fl, fc, fr) = ParseHeaderFooterSections(metadata.OddFooter ?? string.Empty);
            var (fhl, fhc, fhr) = ParseHeaderFooterSections(metadata.FirstHeader ?? string.Empty);
            var (ffl, ffc, ffr) = ParseHeaderFooterSections(metadata.FirstFooter ?? string.Empty);
            var (ehl, ehc, ehr) = ParseHeaderFooterSections(metadata.EvenHeader ?? string.Empty);
            var (efl, efc, efr) = ParseHeaderFooterSections(metadata.EvenFooter ?? string.Empty);

            var (imagesByShapeId, hasPicturePart) = metadata.HasPicture
                ? ReadHeaderFooterImages(maximumTotalSourceImageBytes)
                : (new Dictionary<string, ExcelHeaderFooterImageInfo>(StringComparer.OrdinalIgnoreCase), false);
            bool hasHeaderImageRel = imagesByShapeId.ContainsKey("LH") || imagesByShapeId.ContainsKey("CH") || imagesByShapeId.ContainsKey("RH");
            bool hasFooterImageRel = imagesByShapeId.ContainsKey("LF") || imagesByShapeId.ContainsKey("CF") || imagesByShapeId.ContainsKey("RF");
            if (hasPicturePart && imagesByShapeId.Count == 0) {
                hasHeaderImageRel = true;
                hasFooterImageRel = true;
            }

            return new ExcelHeaderFooterInfo {
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
                DifferentFirstPage = metadata.First,
                DifferentOddEven = metadata.OddEven,
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
                if (HasHeaderFooterCode(value, 'G')) {
                    return true;
                }
            }

            return false;
        }

        private (Dictionary<string, ExcelHeaderFooterImageInfo> Images, bool HasDrawingPart) ReadHeaderFooterImages(long maximumTotalSourceImageBytes) =>
            _excelDocument.ExecuteReadWithoutMaterializing(() => {
                if (DeferredEntry is { } entry && _excelDocument.TryReadWorksheetPackage(entry,
                    (package, sheet) => ReadOwnedHeaderFooterImages(package, sheet, maximumTotalSourceImageBytes), out var owned)) return owned;

                CommitCommentVmlChanges();
                var empty = new Dictionary<string, ExcelHeaderFooterImageInfo>(StringComparer.OrdinalIgnoreCase);
                try {
                    string? id = ViewWorksheetRoot.Child(WorksheetMarkup, "legacyDrawingHF").Attribute(OfficeXmlNamespaces.Relationships, "id");
                    ExcelPartRef drawing = ExcelPartRef.WorksheetForMetadata(this).GetPartById(id);
                    if (drawing.ContentType != HeaderFooterVmlContentType) return (empty, false);
                    using Stream source = drawing.OpenRead();
                    return (ReadVmlHeaderFooterImages(source, maximumTotalSourceImageBytes, (relationship, limit) => {
                        ExcelPartRef image = drawing.GetPartById(relationship);
                        string type = image.ContentType;
                        if (!type.StartsWith("image/", StringComparison.OrdinalIgnoreCase)) return default;
                        using Stream data = image.OpenRead();
                        return ExcelImageExportLimits.TryReadSourceImageBytes(data, limit, out byte[] bytes)
                            ? (bytes, type) : default;
                    }), true);
                } catch (Exception exception) when (exception is IOException || exception is XmlException
                    || exception is ArgumentException || exception is InvalidOperationException) {
                    return (empty, false);
                }
            });

        private static Dictionary<string, ExcelHeaderFooterImageInfo> ReadVmlHeaderFooterImages(Stream source,
            long maximumTotalSourceImageBytes, Func<string, long, (byte[]? Bytes, string? ContentType)> readImage) {
            var images = new Dictionary<string, ExcelHeaderFooterImageInfo>(StringComparer.OrdinalIgnoreCase);
            long remainingSourceImageBytes = maximumTotalSourceImageBytes;
            XDocument vmlDocument;
            try {
                vmlDocument = LoadHeaderFooterVmlXDocument(source);
            } catch (Exception exception) when (exception is IOException || exception is XmlException) {
                return images;
            }

            foreach (XElement shape in vmlDocument.Descendants().Where(element => string.Equals(element.Name.LocalName, "shape", StringComparison.OrdinalIgnoreCase))) {
                string? shapeId = shape.Attribute("id")?.Value;
                if (string.IsNullOrWhiteSpace(shapeId) || !TryGetHeaderFooterPosition(shapeId!, out bool isHeader, out ExcelHeaderFooterPosition position)) {
                    continue;
                }

                XElement? imageData = shape.Descendants().FirstOrDefault(element => string.Equals(element.Name.LocalName, "imagedata", StringComparison.OrdinalIgnoreCase));
                if (imageData == null) {
                    continue;
                }

                // Imported VML can retain a stale r:id while o:relid points to the actual image.
                foreach (string relationshipId in new[] {
                    imageData.Attribute(HeaderFooterRelationshipNamespace + "id")?.Value,
                    imageData.Attribute(HeaderFooterOfficeNamespace + "relid")?.Value
                }.Where(value => !string.IsNullOrWhiteSpace(value)).Select(value => value!).Distinct(StringComparer.Ordinal)) {
                    if (TryReadHeaderFooterImage(readImage, relationshipId, shape.Attribute("style")?.Value, position, ref remainingSourceImageBytes, out ExcelHeaderFooterImageInfo? image)) {
                        images[shapeId!] = image!;
                        break;
                    }
                }
            }

            return images;
        }

        /// <summary>
        /// Reads header/footer VML with the standard relationship prefix available to legacy
        /// producers that omit its declaration. Explicit namespace declarations still take precedence.
        /// </summary>
        private static XDocument LoadHeaderFooterVmlXDocument(Stream stream) {
            if (stream.CanSeek && stream.Length > MaxVmlXmlPartBytes) {
                throw new InvalidDataException($"Excel VML XML part exceeds {MaxVmlXmlPartBytes} bytes.");
            }

            var nameTable = new NameTable();
            var namespaces = new XmlNamespaceManager(nameTable);
            namespaces.AddNamespace("r", HeaderFooterRelationshipNamespace.NamespaceName);
            var context = new XmlParserContext(nameTable, namespaces, null, XmlSpace.None);
            using XmlReader reader = XmlReader.Create(stream, VmlXmlReaderSettings, context);
            return XDocument.Load(reader, LoadOptions.None);
        }

        private static bool TryReadHeaderFooterImage(
            Func<string, long, (byte[]? Bytes, string? ContentType)> readImage,
            string relationshipId,
            string? style,
            ExcelHeaderFooterPosition position,
            ref long remainingSourceImageBytes,
            out ExcelHeaderFooterImageInfo? image) {
            image = null;
            (byte[]? bytes, string? type) = readImage(relationshipId, remainingSourceImageBytes);
            if (bytes == null || type == null) return false;
            remainingSourceImageBytes -= bytes.LongLength;

            double width = TryReadStylePoints(style, "width") ?? 0D;
            double height = TryReadStylePoints(style, "height") ?? 0D;
            if (width <= 0D || height <= 0D) {
                try {
                    var info = OfficeImageReader.Identify(bytes);
                    if (width <= 0D) {
                        width = info.Width * 72D / info.DpiX;
                    }

                    if (height <= 0D) {
                        height = info.Height * 72D / info.DpiY;
                    }
                } catch {
                    width = width <= 0D ? 144D : width;
                    height = height <= 0D ? 48D : height;
                }
            }

            try {
                image = new ExcelHeaderFooterImageInfo(position, bytes, type, OfficeLength.Points(width), OfficeLength.Points(height));
                return true;
            } catch (Exception exception) when (exception is ArgumentOutOfRangeException || exception is OverflowException) {
                image = null;
                return false;
            }
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

                if (OfficeIMO.Core.Internal.OfficeRoundTripText.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double points) && points > 0D) {
                    return points;
                }
            }

            return null;
        }

        private static bool TryGetHeaderFooterPosition(string shapeId, out bool isHeader, out ExcelHeaderFooterPosition position) {
            isHeader = false;
            position = ExcelHeaderFooterPosition.Left;
            switch (shapeId.ToUpperInvariant()) {
                case "LH":
                    isHeader = true;
                    position = ExcelHeaderFooterPosition.Left;
                    return true;
                case "CH":
                    isHeader = true;
                    position = ExcelHeaderFooterPosition.Center;
                    return true;
                case "RH":
                    isHeader = true;
                    position = ExcelHeaderFooterPosition.Right;
                    return true;
                case "LF":
                    position = ExcelHeaderFooterPosition.Left;
                    return true;
                case "CF":
                    position = ExcelHeaderFooterPosition.Center;
                    return true;
                case "RF":
                    position = ExcelHeaderFooterPosition.Right;
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
        internal void SetHeaderFooterCore(
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
            EditHeaderFooter(footer => {
                footer.SetBooleanAttribute("differentFirst", differentFirstPage ? true : (bool?)null);
                footer.SetBooleanAttribute("differentOddEven", differentOddEven ? true : (bool?)null);
                footer.SetBooleanAttribute("alignWithMargins", alignWithMargins);
                footer.SetBooleanAttribute("scaleWithDoc", scaleWithDoc);
                string? oddHeader = BuildHeaderFooterSections(headerLeft, headerCenter, headerRight);
                string? oddFooter = BuildHeaderFooterSections(footerLeft, footerCenter, footerRight);
                SetHeaderFooterText(footer, ExcelHeaderFooterKind.OddHeader, oddHeader);
                SetHeaderFooterText(footer, ExcelHeaderFooterKind.OddFooter, oddFooter);
                if (!differentOddEven || oddHeader != null)
                    SetHeaderFooterText(footer, ExcelHeaderFooterKind.EvenHeader, differentOddEven ? oddHeader : null);
                if (!differentOddEven || oddFooter != null)
                    SetHeaderFooterText(footer, ExcelHeaderFooterKind.EvenFooter, differentOddEven ? oddFooter : null);
                if (!differentFirstPage || oddHeader != null)
                    SetHeaderFooterText(footer, ExcelHeaderFooterKind.FirstHeader, differentFirstPage ? oddHeader : null);
                if (!differentFirstPage || oddFooter != null)
                    SetHeaderFooterText(footer, ExcelHeaderFooterKind.FirstFooter, differentFirstPage ? oddFooter : null);
            });
        }

        /// <summary>
        /// Sets a first-page header and/or footer variant for this worksheet.
        /// </summary>
        internal void SetFirstPageHeaderFooterCore(
            string? headerLeft = null,
            string? headerCenter = null,
            string? headerRight = null,
            string? footerLeft = null,
            string? footerCenter = null,
            string? footerRight = null,
            bool enabled = true) {
            EditHeaderFooter(footer => {
                footer.SetBooleanAttribute("differentFirst", enabled ? true : (bool?)null);
                SetHeaderFooterText(footer, ExcelHeaderFooterKind.FirstHeader, enabled ? BuildHeaderFooterSections(headerLeft, headerCenter, headerRight) : null);
                SetHeaderFooterText(footer, ExcelHeaderFooterKind.FirstFooter, enabled ? BuildHeaderFooterSections(footerLeft, footerCenter, footerRight) : null);
            });
        }

        /// <summary>
        /// Sets an even-page header and/or footer variant for this worksheet.
        /// </summary>
        internal void SetEvenPageHeaderFooterCore(
            string? headerLeft = null,
            string? headerCenter = null,
            string? headerRight = null,
            string? footerLeft = null,
            string? footerCenter = null,
            string? footerRight = null,
            bool enabled = true) {
            EditHeaderFooter(footer => {
                footer.SetBooleanAttribute("differentOddEven", enabled ? true : (bool?)null);
                SetHeaderFooterText(footer, ExcelHeaderFooterKind.EvenHeader, enabled ? BuildHeaderFooterSections(headerLeft, headerCenter, headerRight) : null);
                SetHeaderFooterText(footer, ExcelHeaderFooterKind.EvenFooter, enabled ? BuildHeaderFooterSections(footerLeft, footerCenter, footerRight) : null);
            });
        }

        private static OfficeXmlNode EnsureHeaderFooter(OfficeXmlNode worksheet) =>
            AddWorksheetObject(worksheet, "headerFooter", "rowBreaks", "colBreaks", "customProperties",
                "cellWatches", "ignoredErrors", "smartTags", "drawing", "legacyDrawing", "legacyDrawingHF",
                "picture", "oleObjects", "controls", "webPublishItems", "tableParts");

        private void EditHeaderFooter(Action<OfficeXmlNode> edit) =>
            EditWorksheetObject(worksheet => {
                edit(EnsureHeaderFooter(worksheet));
                if (worksheet.IsSdk && !worksheet.Child(WorksheetMarkup, "legacyDrawingHF").IsNone) {
                    CleanupHeaderFooterPictureArtifacts();
                }
                return true;
            }, afterOwnedEdit: CleanupOwnedHeaderFooterPicture);

        /// <summary>
        /// Adds an image to the worksheet header at the given position. This will also ensure the header text
        /// contains the picture placeholder (&amp;G) in the corresponding section. Subsequent calls replace any
        /// previously set header/footer images for this sheet.
        /// </summary>
        /// <param name="position">Left, Center, or Right header section.</param>
        /// <param name="imageBytes">Image bytes.</param>
        /// <param name="contentType">e.g. image/png, image/jpeg. Defaults to image/png.</param>
        /// <param name="width">Optional physical width. If omitted, inferred from image size at 96 DPI.</param>
        /// <param name="height">Optional physical height. If omitted, inferred proportionally.</param>
        internal void SetHeaderImageCore(ExcelHeaderFooterPosition position, byte[] imageBytes, string contentType = "image/png", OfficeLength? width = null, OfficeLength? height = null) {
            if (imageBytes == null || imageBytes.Length == 0) throw new ArgumentException("Image bytes are required.", nameof(imageBytes));
            WriteLock(() => {
                var normalizedContentType = NormalizeImageContentType(contentType, nameof(contentType));
                EnsureHeaderFooterPicture(position, isHeader: true, imageBytes, normalizedContentType, width, height);
            });
        }

        internal async Task SetHeaderImageFromUrlCoreAsync(ExcelHeaderFooterPosition position, string url,
            OfficeRemoteImageLoadOptions? remoteImageOptions, OfficeLength? width, OfficeLength? height,
            CancellationToken cancellationToken) {
            if (width.HasValue && width.Value <= OfficeLength.Zero) throw new ArgumentOutOfRangeException(nameof(width));
            if (height.HasValue && height.Value <= OfficeLength.Zero) throw new ArgumentOutOfRangeException(nameof(height));
            OfficeRemoteImage remote = await OfficeRemoteImageLoader.LoadAsync(
                url,
                remoteImageOptions,
                cancellationToken).ConfigureAwait(false);
            SetHeaderImageCore(position, remote.ToBytes(), remote.ContentType, width, height);
        }

        /// <summary>
        /// Adds an image to the worksheet footer at the given position. This will also ensure the footer text
        /// contains the picture placeholder (&amp;G) in the corresponding section. Subsequent calls replace any
        /// previously set header/footer images for this sheet.
        /// </summary>
        internal void SetFooterImageCore(ExcelHeaderFooterPosition position, byte[] imageBytes, string contentType = "image/png", OfficeLength? width = null, OfficeLength? height = null) {
            if (imageBytes == null || imageBytes.Length == 0) throw new ArgumentException("Image bytes are required.", nameof(imageBytes));
            WriteLock(() => {
                var normalizedContentType = NormalizeImageContentType(contentType, nameof(contentType));
                EnsureHeaderFooterPicture(position, isHeader: false, imageBytes, normalizedContentType, width, height);
            });
        }

        internal async Task SetFooterImageFromUrlCoreAsync(ExcelHeaderFooterPosition position, string url,
            OfficeRemoteImageLoadOptions? remoteImageOptions, OfficeLength? width, OfficeLength? height,
            CancellationToken cancellationToken) {
            if (width.HasValue && width.Value <= OfficeLength.Zero) throw new ArgumentOutOfRangeException(nameof(width));
            if (height.HasValue && height.Value <= OfficeLength.Zero) throw new ArgumentOutOfRangeException(nameof(height));
            OfficeRemoteImage remote = await OfficeRemoteImageLoader.LoadAsync(
                url,
                remoteImageOptions,
                cancellationToken).ConfigureAwait(false);
            SetFooterImageCore(position, remote.ToBytes(), remote.ContentType, width, height);
        }

        private void EnsureHeaderFooterPicture(ExcelHeaderFooterPosition position, bool isHeader, byte[] imageBytes, string contentType, OfficeLength? width, OfficeLength? height) {
            if (width.HasValue && width.Value <= OfficeLength.Zero) throw new ArgumentOutOfRangeException(nameof(width));
            if (height.HasValue && height.Value <= OfficeLength.Zero) throw new ArgumentOutOfRangeException(nameof(height));
            (double widthPoints, double heightPoints) = ResolveHeaderFooterImageSize(imageBytes, width, height);
            string shape = (position == ExcelHeaderFooterPosition.Left ? "L" : position == ExcelHeaderFooterPosition.Center ? "C" : "R")
                + (isHeader ? "H" : "F");
            if (DeferredEntry is { } entry && _excelDocument.TryEditWorksheetObject(entry,
                worksheet => {
                    SetHeaderFooterPicturePlaceholder(worksheet, position, isHeader);
                    return true;
                },
                (package, worksheet, root) => WriteOwnedHeaderFooterPicture(package, worksheet, root, shape,
                    imageBytes, contentType, widthPoints, heightPoints))) return;

            OfficeXmlNode worksheetRoot = ViewWorksheetRoot;
            SetHeaderFooterPicturePlaceholder(worksheetRoot, position, isHeader);
            OfficeXmlNode legacy = worksheetRoot.Child(WorksheetMarkup, "legacyDrawingHF");
            string? id = legacy.Attribute(OfficeXmlNamespaces.Relationships, "id");
            VmlDrawingPart vmlPart;
            if (!string.IsNullOrEmpty(id)) {
                vmlPart = (VmlDrawingPart)_worksheetPart.GetPartById(id!);
            } else {
                vmlPart = _worksheetPart.AddNewPart<VmlDrawingPart>();
                SetHeaderFooterDrawingRelationship(worksheetRoot, _worksheetPart.GetIdOfPart(vmlPart));
            }

            ImagePart imagePart = vmlPart.AddImagePart(ToImagePartType(contentType));
            using (var source = new MemoryStream(imageBytes, writable: false)) imagePart.FeedData(source);
            CommitCommentVmlChanges();
            XDocument vml;
            using (Stream drawing = vmlPart.GetStream(FileMode.Open, FileAccess.Read)) vml = LoadOrCreateHeaderFooterVmlDocument(drawing);
            UpsertHeaderFooterVmlShape(vml, shape, vmlPart.GetIdOfPart(imagePart), widthPoints, heightPoints);
            using (var output = new MemoryStream()) {
                vml.Save(output, SaveOptions.DisableFormatting);
                output.Position = 0;
                vmlPart.FeedData(output);
            }
            ExcelPartRef.WorksheetForMetadata(this).SaveRoot();
        }

        private static void SetHeaderFooterPicturePlaceholder(OfficeXmlNode worksheet, ExcelHeaderFooterPosition position, bool isHeader) {
            OfficeXmlNode footer = EnsureHeaderFooter(worksheet);
            ExcelHeaderFooterKind kind = isHeader ? ExcelHeaderFooterKind.OddHeader : ExcelHeaderFooterKind.OddFooter;
            var (left, center, right) = ParseHeaderFooterSections(GetHeaderFooterText(footer, kind) ?? string.Empty);
            static string WithPicture(string text) => HasHeaderFooterCode(text, 'G') ? text : "&G" + text;
            switch (position) {
                case ExcelHeaderFooterPosition.Left: left = WithPicture(left); break;
                case ExcelHeaderFooterPosition.Center: center = WithPicture(center); break;
                case ExcelHeaderFooterPosition.Right: right = WithPicture(right); break;
            }
            string text = (left.Length == 0 ? string.Empty : "&L" + left)
                + (center.Length == 0 ? string.Empty : "&C" + center)
                + (right.Length == 0 ? string.Empty : "&R" + right);
            SetHeaderFooterText(footer, kind, text.Length == 0 ? null : text);
        }

        private static void SetHeaderFooterDrawingRelationship(OfficeXmlNode worksheet, string id) =>
            AddWorksheetObject(worksheet, "legacyDrawingHF", "picture", "oleObjects", "controls", "webPublishItems", "tableParts")
                .SetAttribute(OfficeXmlNamespaces.Relationships, "id", id);

        private static (double Width, double Height) ResolveHeaderFooterImageSize(byte[] image, OfficeLength? width, OfficeLength? height) {
            double pointsWide = width?.ToPoints() ?? 0;
            double pointsHigh = height?.ToPoints() ?? 0;
            if (pointsWide > 0 && pointsHigh > 0) return (pointsWide, pointsHigh);
            try {
                var dimensions = OfficeImageReader.Identify(image);
                if (pointsWide <= 0) pointsWide = dimensions.Width * 72D / dimensions.DpiX;
                if (pointsHigh <= 0) pointsHigh = dimensions.Height * 72D / dimensions.DpiY;
            } catch {
                if (pointsWide <= 0) pointsWide = 144D;
                if (pointsHigh <= 0) pointsHigh = 48D;
            }
            return (pointsWide, pointsHigh);
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

        private bool HeaderFooterContainsPicturePlaceholder() =>
            WorksheetHeaderContainsPicture(ViewWorksheetRoot);
    }
}
