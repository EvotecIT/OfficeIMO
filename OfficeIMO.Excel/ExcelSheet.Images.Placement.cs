using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Adds an image from disk anchored to a worksheet cell. When <paramref name="scalePercent"/> is provided,
        /// the image is sized from its original pixel dimensions.
        /// </summary>
        /// <param name="row">1-based row index where the top edge of the image will be anchored.</param>
        /// <param name="column">1-based column index where the left edge of the image will be anchored.</param>
        /// <param name="path">Image file path.</param>
        /// <param name="widthPixels">Optional exact image width in pixels. Defaults to the image's original width when known.</param>
        /// <param name="heightPixels">Optional exact image height in pixels. Defaults to the image's original height when known.</param>
        /// <param name="scalePercent">Optional percentage of the original image dimensions, for example 20 for 20%.</param>
        /// <param name="offsetXPixels">Optional horizontal offset in pixels from the cell's left boundary.</param>
        /// <param name="offsetYPixels">Optional vertical offset in pixels from the cell's top boundary.</param>
        /// <param name="name">Optional drawing name used by Excel's selection pane.</param>
        /// <param name="altText">Optional alternative text description for accessibility.</param>
        /// <param name="title">Optional alternative text title.</param>
        /// <param name="lockAspectRatio">Whether Excel should keep the picture aspect ratio locked.</param>
        /// <param name="rotationDegrees">Clockwise image rotation in degrees.</param>
        /// <example>
        /// <code>
        /// using var workbook = ExcelDocument.Create("report.xlsx");
        /// var sheet = workbook.Sheets[0];
        /// sheet.AddImageFromFile(2, 2, "logo.png", scalePercent: 20, name: "Logo", altText: "Company logo");
        /// workbook.Save();
        /// </code>
        /// </example>
        public ExcelImage AddImageFromFile(int row, int column, string path, int? widthPixels = null, int? heightPixels = null,
            double? scalePercent = null, int offsetXPixels = 0, int offsetYPixels = 0, string? name = null, string? altText = null,
            string? title = null, bool lockAspectRatio = true, double rotationDegrees = 0) {
            if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("Image path is required.", nameof(path));
            if (!File.Exists(path)) throw new FileNotFoundException($"Image file '{path}' was not found.", path);

            byte[] bytes = File.ReadAllBytes(path);
            OfficeImageReader.TryIdentify(bytes, path, out OfficeImageInfo info);
            var (resolvedWidth, resolvedHeight) = ResolveImageSize(info, widthPixels, heightPixels, scalePercent);
            string contentType = info.Format == OfficeImageFormat.Unknown ? ContentTypeFromExtension(path) : info.MimeType;

            ExcelImage image = AddImage(row, column, bytes, contentType, resolvedWidth, resolvedHeight, offsetXPixels, offsetYPixels, name, altText, lockAspectRatio);
            ApplyImageMetadata(image, title, rotationDegrees);
            return image;
        }

        /// <summary>
        /// Downloads an image from URL, anchors it to a worksheet cell, and optionally sizes it from its original dimensions.
        /// Returns null when the image cannot be fetched.
        /// </summary>
        /// <param name="row">1-based row index where the top edge of the image will be anchored.</param>
        /// <param name="column">1-based column index where the left edge of the image will be anchored.</param>
        /// <param name="url">Remote image URL to download. Requests timeout after 5 seconds and must be smaller than 2 MB.</param>
        /// <param name="widthPixels">Optional exact image width in pixels. Defaults to the image's original width when known.</param>
        /// <param name="heightPixels">Optional exact image height in pixels. Defaults to the image's original height when known.</param>
        /// <param name="scalePercent">Optional percentage of the original image dimensions, for example 20 for 20%.</param>
        /// <param name="offsetXPixels">Optional horizontal offset in pixels from the cell's left boundary.</param>
        /// <param name="offsetYPixels">Optional vertical offset in pixels from the cell's top boundary.</param>
        /// <param name="name">Optional drawing name used by Excel's selection pane.</param>
        /// <param name="altText">Optional alternative text description for accessibility.</param>
        /// <param name="title">Optional alternative text title.</param>
        /// <param name="lockAspectRatio">Whether Excel should keep the picture aspect ratio locked.</param>
        /// <param name="rotationDegrees">Clockwise image rotation in degrees.</param>
        public ExcelImage? AddImageFromUrl(int row, int column, string url, int? widthPixels, int? heightPixels,
            double? scalePercent = null, int offsetXPixels = 0, int offsetYPixels = 0, string? name = null, string? altText = null,
            string? title = null, bool lockAspectRatio = true, double rotationDegrees = 0) {
            if (string.IsNullOrWhiteSpace(url)) return null;
            if (!ImageDownloader.TryFetch(url, timeoutSeconds: 5, maxBytes: 2_000_000, out var bytes, out var contentType) || bytes == null) {
                return null;
            }

            OfficeImageReader.TryIdentify(bytes, null, out OfficeImageInfo info);
            var (resolvedWidth, resolvedHeight) = ResolveImageSize(info, widthPixels, heightPixels, scalePercent);
            ExcelImage image = AddImage(row, column, bytes, ResolveImageContentType(contentType, info),
                resolvedWidth, resolvedHeight, offsetXPixels, offsetYPixels, name, altText, lockAspectRatio);
            ApplyImageMetadata(image, title, rotationDegrees);
            return image;
        }

        /// <summary>
        /// Downloads an image from URL, scales it from its detected dimensions, and anchors it to a worksheet cell.
        /// Returns null when the image cannot be fetched.
        /// </summary>
        /// <param name="row">1-based row index where the top edge of the image will be anchored.</param>
        /// <param name="column">1-based column index where the left edge of the image will be anchored.</param>
        /// <param name="url">Remote image URL to download. Requests timeout after 5 seconds and must be smaller than 2 MB.</param>
        /// <param name="scalePercent">Percentage of the original image dimensions, for example 20 for 20%.</param>
        /// <param name="offsetXPixels">Optional horizontal offset in pixels from the cell's left boundary.</param>
        /// <param name="offsetYPixels">Optional vertical offset in pixels from the cell's top boundary.</param>
        /// <param name="name">Optional drawing name used by Excel's selection pane.</param>
        /// <param name="altText">Optional alternative text description for accessibility.</param>
        /// <param name="title">Optional alternative text title.</param>
        /// <param name="lockAspectRatio">Whether Excel should keep the picture aspect ratio locked.</param>
        /// <param name="rotationDegrees">Clockwise image rotation in degrees.</param>
        public ExcelImage? AddImageFromUrl(int row, int column, string url, double scalePercent,
            int offsetXPixels = 0, int offsetYPixels = 0, string? name = null, string? altText = null,
            string? title = null, bool lockAspectRatio = true, double rotationDegrees = 0)
            => AddImageFromUrl(row, column, url, null, null, scalePercent, offsetXPixels, offsetYPixels, name, altText, title, lockAspectRatio, rotationDegrees);

        /// <summary>
        /// Adds an image anchored to an A1 range using a two-cell anchor. The image moves and sizes with the
        /// referenced cells by default, matching Excel's "Move and size with cells" behavior.
        /// </summary>
        /// <param name="range">A1 range such as A1:C15. The image is anchored from the top-left of the first cell to the bottom-right boundary of the last cell.</param>
        /// <param name="imageBytes">Image bytes.</param>
        /// <param name="contentType">Content type, for example image/png or image/jpeg.</param>
        /// <param name="offsetXPixels">Optional horizontal offset from the range start.</param>
        /// <param name="offsetYPixels">Optional vertical offset from the range start.</param>
        /// <param name="endOffsetXPixels">Optional horizontal offset applied to the range end marker.</param>
        /// <param name="endOffsetYPixels">Optional vertical offset applied to the range end marker.</param>
        /// <param name="name">Optional drawing name used by Excel's selection pane.</param>
        /// <param name="altText">Optional alternative text description for accessibility.</param>
        /// <param name="title">Optional alternative text title.</param>
        /// <param name="lockAspectRatio">Whether Excel should keep the picture aspect ratio locked.</param>
        /// <param name="placement">How the image behaves when cells move or resize.</param>
        /// <param name="rotationDegrees">Clockwise image rotation in degrees.</param>
        /// <example>
        /// <code>
        /// using var workbook = ExcelDocument.Create("report.xlsx");
        /// var sheet = workbook.Sheets[0];
        /// byte[] logo = File.ReadAllBytes("logo.png");
        /// sheet.AddImageToRange("A1:C15", logo, "image/png", name: "PinnedLogo",
        ///     altText: "Company logo", placement: ExcelImagePlacement.MoveAndSize);
        /// workbook.Save();
        /// </code>
        /// </example>
        public ExcelImage AddImageToRange(string range, byte[] imageBytes, string contentType = "image/png",
            int offsetXPixels = 0, int offsetYPixels = 0, int endOffsetXPixels = 0, int endOffsetYPixels = 0,
            string? name = null, string? altText = null, string? title = null, bool lockAspectRatio = true,
            ExcelImagePlacement placement = ExcelImagePlacement.MoveAndSize, double rotationDegrees = 0) {
            if (imageBytes == null || imageBytes.Length == 0) throw new ArgumentException("Image bytes are required.", nameof(imageBytes));
            var (startRow, startColumn, endRow, endColumn) = ParseImageRange(range);

            ExcelImage? image = null;
            WriteLock(() => {
                DrawingsPart drawingPart = GetOrCreateDrawingsPart();
                ImagePart imagePart = drawingPart.AddImagePart(ToImagePartType(contentType));
                using (var stream = new MemoryStream(imageBytes)) imagePart.FeedData(stream);
                string imageRelationshipId = drawingPart.GetIdOfPart(imagePart);

                var drawingId = NextDrawingId(drawingPart);
                string resolvedName = string.IsNullOrWhiteSpace(name) ? $"Picture {drawingId}" : name!.Trim();
                var (widthPixels, heightPixels) = CalculateRangeAnchorSizePixels(
                    startRow,
                    startColumn,
                    endRow,
                    endColumn,
                    offsetXPixels,
                    offsetYPixels,
                    endOffsetXPixels,
                    endOffsetYPixels);
                var anchor = new Xdr.TwoCellAnchor(
                    new Xdr.FromMarker(
                        new Xdr.ColumnId((startColumn - 1).ToString(System.Globalization.CultureInfo.InvariantCulture)),
                        new Xdr.ColumnOffset(PxToEmu(offsetXPixels).ToString(System.Globalization.CultureInfo.InvariantCulture)),
                        new Xdr.RowId((startRow - 1).ToString(System.Globalization.CultureInfo.InvariantCulture)),
                        new Xdr.RowOffset(PxToEmu(offsetYPixels).ToString(System.Globalization.CultureInfo.InvariantCulture))
                    ),
                    new Xdr.ToMarker(
                        new Xdr.ColumnId(endColumn.ToString(System.Globalization.CultureInfo.InvariantCulture)),
                        new Xdr.ColumnOffset(PxToEmu(endOffsetXPixels).ToString(System.Globalization.CultureInfo.InvariantCulture)),
                        new Xdr.RowId(endRow.ToString(System.Globalization.CultureInfo.InvariantCulture)),
                        new Xdr.RowOffset(PxToEmu(endOffsetYPixels).ToString(System.Globalization.CultureInfo.InvariantCulture))
                    ),
                    CreatePicture(drawingId, resolvedName, imageRelationshipId, altText, title, lockAspectRatio,
                        PxToEmu(widthPixels), PxToEmu(heightPixels), rotationDegrees),
                    new Xdr.ClientData()) {
                    EditAs = ToEditAsValue(placement)
                };

                Xdr.WorksheetDrawing worksheetDrawing = drawingPart.WorksheetDrawing!;
                worksheetDrawing.Append(anchor);
                worksheetDrawing.Save();
                WorksheetRoot.Save();
                _excelDocument.MarkPackageDirty();
                image = new ExcelImage(anchor.GetFirstChild<Xdr.Picture>()!, anchor, drawingPart, _excelDocument);
            });

            return image!;
        }

        /// <summary>
        /// Adds an image from disk anchored to an A1 range using a two-cell anchor.
        /// </summary>
        /// <param name="range">A1 range such as A1:C15. The image is anchored from the top-left of the first cell to the bottom-right boundary of the last cell.</param>
        /// <param name="path">Image file path.</param>
        /// <param name="offsetXPixels">Optional horizontal offset from the range start.</param>
        /// <param name="offsetYPixels">Optional vertical offset from the range start.</param>
        /// <param name="endOffsetXPixels">Optional horizontal offset applied to the range end marker.</param>
        /// <param name="endOffsetYPixels">Optional vertical offset applied to the range end marker.</param>
        /// <param name="name">Optional drawing name used by Excel's selection pane.</param>
        /// <param name="altText">Optional alternative text description for accessibility.</param>
        /// <param name="title">Optional alternative text title.</param>
        /// <param name="lockAspectRatio">Whether Excel should keep the picture aspect ratio locked.</param>
        /// <param name="placement">How the image behaves when cells move or resize.</param>
        /// <param name="rotationDegrees">Clockwise image rotation in degrees.</param>
        public ExcelImage AddImageFromFileToRange(string range, string path, int offsetXPixels = 0, int offsetYPixels = 0,
            int endOffsetXPixels = 0, int endOffsetYPixels = 0, string? name = null, string? altText = null, string? title = null,
            bool lockAspectRatio = true, ExcelImagePlacement placement = ExcelImagePlacement.MoveAndSize, double rotationDegrees = 0) {
            if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("Image path is required.", nameof(path));
            if (!File.Exists(path)) throw new FileNotFoundException($"Image file '{path}' was not found.", path);

            byte[] bytes = File.ReadAllBytes(path);
            OfficeImageReader.TryIdentify(bytes, path, out OfficeImageInfo info);
            string contentType = info.Format == OfficeImageFormat.Unknown ? ContentTypeFromExtension(path) : info.MimeType;
            return AddImageToRange(range, bytes, contentType, offsetXPixels, offsetYPixels, endOffsetXPixels, endOffsetYPixels,
                name, altText, title, lockAspectRatio, placement, rotationDegrees);
        }

        /// <summary>
        /// Downloads an image from URL and anchors it to an A1 range using a two-cell anchor. Returns null when the image cannot be fetched.
        /// </summary>
        public ExcelImage? AddImageFromUrlToRange(string range, string url, int offsetXPixels = 0, int offsetYPixels = 0,
            int endOffsetXPixels = 0, int endOffsetYPixels = 0, string? name = null, string? altText = null, string? title = null,
            bool lockAspectRatio = true, ExcelImagePlacement placement = ExcelImagePlacement.MoveAndSize, double rotationDegrees = 0) {
            if (string.IsNullOrWhiteSpace(url)) return null;
            if (!ImageDownloader.TryFetch(url, timeoutSeconds: 5, maxBytes: 2_000_000, out var bytes, out var contentType) || bytes == null) {
                return null;
            }

            return AddImageToRange(range, bytes, string.IsNullOrEmpty(contentType) ? "image/png" : contentType!, offsetXPixels,
                offsetYPixels, endOffsetXPixels, endOffsetYPixels, name, altText, title, lockAspectRatio, placement, rotationDegrees);
        }

        private DrawingsPart GetOrCreateDrawingsPart() {
            var drawing = WorksheetRoot.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Drawing>();
            if (drawing == null) {
                DrawingsPart drawingPart = _worksheetPart.AddNewPart<DrawingsPart>();
                drawingPart.WorksheetDrawing = new Xdr.WorksheetDrawing();
                string relationshipId = _worksheetPart.GetIdOfPart(drawingPart);
                WorksheetRoot.Append(new DocumentFormat.OpenXml.Spreadsheet.Drawing { Id = relationshipId });
                return drawingPart;
            }

            var existing = (DrawingsPart)_worksheetPart.GetPartById(drawing.Id!);
            existing.WorksheetDrawing ??= new Xdr.WorksheetDrawing();
            return existing;
        }

        private static Xdr.Picture CreatePicture(UInt32Value drawingId, string name, string imageRelationshipId, string? altText,
            string? title, bool lockAspectRatio, long widthEmu, long heightEmu, double rotationDegrees) {
            var drawingProperties = new Xdr.NonVisualDrawingProperties {
                Id = drawingId,
                Name = name,
                Description = altText ?? string.Empty
            };
            if (!string.IsNullOrWhiteSpace(title)) {
                drawingProperties.Title = title;
            }

            var transform = new A.Transform2D(
                new A.Offset { X = 0, Y = 0 },
                new A.Extents { Cx = widthEmu, Cy = heightEmu });
            if (Math.Abs(rotationDegrees) > double.Epsilon) {
                transform.Rotation = (int)Math.Round(rotationDegrees * 60000.0);
            }

            return new Xdr.Picture(
                new Xdr.NonVisualPictureProperties(
                    drawingProperties,
                    new Xdr.NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = lockAspectRatio })
                ),
                new Xdr.BlipFill(
                    new A.Blip { Embed = imageRelationshipId },
                    new A.Stretch(new A.FillRectangle())
                ),
                new Xdr.ShapeProperties(
                    transform,
                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
                )
            );
        }

        private static (int Width, int Height) ResolveImageSize(OfficeImageInfo info, int? widthPixels, int? heightPixels, double? scalePercent) {
            if (scalePercent.HasValue) {
                if (widthPixels.HasValue || heightPixels.HasValue) {
                    throw new ArgumentException("Scale percentage cannot be combined with explicit width or height.");
                }

                if (double.IsNaN(scalePercent.Value) || double.IsInfinity(scalePercent.Value) || scalePercent.Value <= 0) {
                    throw new ArgumentOutOfRangeException(nameof(scalePercent), "Scale percentage must be a positive finite number.");
                }

                if (info.Width <= 0 || info.Height <= 0) {
                    throw new NotSupportedException("Image dimensions could not be detected, so scale percentage cannot be applied.");
                }

                return (Math.Max(1, (int)Math.Round(info.Width * scalePercent.Value / 100.0)),
                    Math.Max(1, (int)Math.Round(info.Height * scalePercent.Value / 100.0)));
            }

            int width = widthPixels ?? (info.Width > 0 ? info.Width : 96);
            int height = heightPixels ?? (info.Height > 0 ? info.Height : 32);
            if (width <= 0) throw new ArgumentOutOfRangeException(nameof(widthPixels));
            if (height <= 0) throw new ArgumentOutOfRangeException(nameof(heightPixels));
            return (width, height);
        }

        private static (int StartRow, int StartColumn, int EndRow, int EndColumn) ParseImageRange(string range) {
            if (!A1.TryParseRange(range, out int startRow, out int startColumn, out int endRow, out int endColumn)) {
                throw new ArgumentException($"Invalid A1 range '{range}'. Use a range such as A1:C15.", nameof(range));
            }

            if (endColumn >= A1.MaxColumns) {
                throw new ArgumentOutOfRangeException(nameof(range), "Image range must end before the final Excel column so the two-cell end marker can use the boundary after the last cell.");
            }

            if (endRow >= A1.MaxRows) {
                throw new ArgumentOutOfRangeException(nameof(range), "Image range must end before the final Excel row so the two-cell end marker can use the boundary after the last cell.");
            }

            return (startRow, startColumn, endRow, endColumn);
        }

        private (int WidthPixels, int HeightPixels) CalculateRangeAnchorSizePixels(
            int startRow,
            int startColumn,
            int endRow,
            int endColumn,
            int offsetXPixels,
            int offsetYPixels,
            int endOffsetXPixels,
            int endOffsetYPixels) {
            int width = Math.Max(1, CalculateColumnSpanPixels(startColumn, endColumn) + endOffsetXPixels - offsetXPixels);
            int height = Math.Max(1, CalculateRowSpanPixels(startRow, endRow) + endOffsetYPixels - offsetYPixels);
            return (width, height);
        }

        private int CalculateColumnSpanPixels(int startColumn, int endColumn) {
            ExcelTextMeasurer textMeasurer = ExcelTextMeasurer.Create(GetWorkbookDefaultFontInfo());
            float mdw = textMeasurer.DefaultStyle.MaximumDigitWidth;
            if (mdw <= 0.0001f) {
                mdw = 7f;
            }

            double total = 0;
            for (int column = startColumn; column <= endColumn; column++) {
                if (IsColumnHidden(column)) {
                    continue;
                }

                total += GetColumnWidthPixels(column, mdw);
            }

            return Math.Max(1, (int)Math.Round(total));
        }

        private int CalculateRowSpanPixels(int startRow, int endRow) {
            double total = 0;
            for (int rowIndex = startRow; rowIndex <= endRow; rowIndex++) {
                total += GetRowHeightPixels(rowIndex);
            }

            return Math.Max(1, (int)Math.Round(total));
        }

        private bool IsColumnHidden(int columnIndex) {
            var columns = WorksheetRoot.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Columns>();
            var column = columns?.Elements<DocumentFormat.OpenXml.Spreadsheet.Column>()
                .FirstOrDefault(c => c.Min != null && c.Max != null && c.Min.Value <= (uint)columnIndex && c.Max.Value >= (uint)columnIndex);
            return column?.Hidden?.Value == true;
        }

        private double GetRowHeightPixels(int rowIndex) {
            var sheetData = WorksheetRoot.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetData>();
            var row = sheetData?.Elements<DocumentFormat.OpenXml.Spreadsheet.Row>()
                .FirstOrDefault(r => r.RowIndex != null && r.RowIndex.Value == (uint)rowIndex);
            if (row?.Hidden?.Value == true) {
                return 0;
            }

            double heightPoints = row?.Height?.Value > 0 && row.CustomHeight?.Value == true
                ? row.Height.Value
                : GetDefaultRowHeightPoints();
            return heightPoints * 96D / 72D;
        }

        private static void ApplyImageMetadata(ExcelImage image, string? title, double rotationDegrees) {
            if (!string.IsNullOrWhiteSpace(title)) {
                image.Title = title!;
            }

            if (Math.Abs(rotationDegrees) > double.Epsilon) {
                image.SetRotation(rotationDegrees);
            }
        }

        private static PartTypeInfo ToImagePartType(string? contentType) {
            return (contentType ?? string.Empty).ToLowerInvariant() switch {
                "image/png" => ImagePartType.Png,
                "image/jpeg" or "image/jpg" => ImagePartType.Jpeg,
                "image/gif" => ImagePartType.Gif,
                "image/bmp" => ImagePartType.Bmp,
                "image/tiff" or "image/tif" => ImagePartType.Tiff,
                "image/svg+xml" or "image/svg" => ImagePartType.Svg,
                "image/x-emf" or "image/emf" => ImagePartType.Emf,
                "image/x-wmf" or "image/wmf" => ImagePartType.Wmf,
                "image/x-icon" or "image/vnd.microsoft.icon" or "image/ico" => ImagePartType.Icon,
                "image/x-pcx" or "image/pcx" => ImagePartType.Pcx,
                _ => throw new NotSupportedException($"Image content type '{contentType}' is not supported by Excel image parts.")
            };
        }

        private static string ContentTypeFromExtension(string path) {
            return OfficeImageReader.FromExtension(path) switch {
                OfficeImageFormat.Png => "image/png",
                OfficeImageFormat.Jpeg => "image/jpeg",
                OfficeImageFormat.Gif => "image/gif",
                OfficeImageFormat.Bmp => "image/bmp",
                OfficeImageFormat.Tiff => "image/tiff",
                OfficeImageFormat.Svg => "image/svg+xml",
                OfficeImageFormat.Emf => "image/x-emf",
                OfficeImageFormat.Wmf => "image/x-wmf",
                OfficeImageFormat.Icon => "image/x-icon",
                OfficeImageFormat.Pcx => "image/x-pcx",
                _ => "application/octet-stream"
            };
        }

        private static string ResolveImageContentType(string? declaredContentType, OfficeImageInfo detectedInfo) {
            if (detectedInfo.Format != OfficeImageFormat.Unknown) {
                return detectedInfo.MimeType;
            }

            return string.IsNullOrWhiteSpace(declaredContentType) ? "application/octet-stream" : declaredContentType!;
        }

        private static EnumValue<Xdr.EditAsValues> ToEditAsValue(ExcelImagePlacement placement) {
            return placement switch {
                ExcelImagePlacement.MoveAndSize => Xdr.EditAsValues.TwoCell,
                ExcelImagePlacement.MoveOnly => Xdr.EditAsValues.OneCell,
                ExcelImagePlacement.FreeFloating => Xdr.EditAsValues.Absolute,
                _ => Xdr.EditAsValues.TwoCell
            };
        }
    }
}
