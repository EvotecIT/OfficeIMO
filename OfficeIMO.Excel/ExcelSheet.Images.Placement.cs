using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        // Writes native range and absolute anchors without changing their marker or extent arithmetic.
        private ExcelImage AddRangeImageCore(string range, byte[] imageBytes, string contentType = "image/png",
            OfficeLength offsetX = default, OfficeLength offsetY = default, OfficeLength endOffsetX = default, OfficeLength endOffsetY = default,
            string? name = null, string? altText = null, string? title = null, bool lockAspectRatio = true,
            ExcelImagePlacement placement = ExcelImagePlacement.MoveAndSize, double rotationDegrees = 0, bool deduplicate = false) {
            if (imageBytes == null || imageBytes.Length == 0) throw new ArgumentException("Image bytes are required.", nameof(imageBytes));
            var (startRow, startColumn, endRow, endColumn) = ParseImageRange(range);
            ExcelImage? image = null;
            WriteLock(() => {
                DrawingsPart drawingPart = GetOrCreateDrawingsPart();
                ImagePart imagePart = GetImagePart(drawingPart, imageBytes, contentType, deduplicate);
                string imageRelationshipId = drawingPart.GetIdOfPart(imagePart);

                var drawingId = NextDrawingId(drawingPart);
                string resolvedName = string.IsNullOrWhiteSpace(name) ? $"Picture {drawingId}" : name!.Trim();
                var (width, height) = CalculateRangeAnchorSize(
                    startRow,
                    startColumn,
                    endRow,
                    endColumn,
                    offsetX,
                    offsetY,
                    endOffsetX,
                    endOffsetY);
                var anchor = new Xdr.TwoCellAnchor(
                    new Xdr.FromMarker(
                        new Xdr.ColumnId((startColumn - 1).ToString(System.Globalization.CultureInfo.InvariantCulture)),
                        new Xdr.ColumnOffset(offsetX.ToEmus().ToString(System.Globalization.CultureInfo.InvariantCulture)),
                        new Xdr.RowId((startRow - 1).ToString(System.Globalization.CultureInfo.InvariantCulture)),
                        new Xdr.RowOffset(offsetY.ToEmus().ToString(System.Globalization.CultureInfo.InvariantCulture))
                    ),
                    new Xdr.ToMarker(
                        new Xdr.ColumnId(endColumn.ToString(System.Globalization.CultureInfo.InvariantCulture)),
                        new Xdr.ColumnOffset(endOffsetX.ToEmus().ToString(System.Globalization.CultureInfo.InvariantCulture)),
                        new Xdr.RowId(endRow.ToString(System.Globalization.CultureInfo.InvariantCulture)),
                        new Xdr.RowOffset(endOffsetY.ToEmus().ToString(System.Globalization.CultureInfo.InvariantCulture))
                    ),
                    CreatePicture(drawingId, resolvedName, imageRelationshipId, altText, title, lockAspectRatio,
                        width.ToEmus(), height.ToEmus(), rotationDegrees),
                    new Xdr.ClientData()) {
                    EditAs = ToEditAsValue(placement)
                };

                Xdr.WorksheetDrawing worksheetDrawing = drawingPart.WorksheetDrawing!;
                worksheetDrawing.Append(anchor);
                worksheetDrawing.Save();
                WorksheetRoot.Save();
                _excelDocument.MarkPackageDirty();
                image = new ExcelImage(anchor.GetFirstChild<Xdr.Picture>()!, anchor, drawingPart, this);
            });

            return image!;
        }

        /// <summary>Restores a native absolute image, including signed positions bounded by the importing adapter.</summary>
        internal ExcelImage AddNativeAbsoluteImage(OfficeLength x, OfficeLength y, byte[] imageBytes, string contentType,
            OfficeLength? width, OfficeLength? height, string? name = null, string? altText = null, string? title = null,
            bool lockAspectRatio = true, double rotationDegrees = 0, bool deduplicate = false) {
            if (imageBytes == null || imageBytes.Length == 0) throw new ArgumentException("Image bytes are required.", nameof(imageBytes));
            OfficeLength resolvedWidth = width ?? OfficeLength.Pixels(96);
            OfficeLength resolvedHeight = height ?? OfficeLength.Pixels(32);
            if (resolvedWidth <= OfficeLength.Zero) throw new ArgumentOutOfRangeException(nameof(width));
            if (resolvedHeight <= OfficeLength.Zero) throw new ArgumentOutOfRangeException(nameof(height));
            ExcelImage? image = null;
            WriteLock(() => {
                DrawingsPart drawingPart = GetOrCreateDrawingsPart();
                ImagePart imagePart = GetImagePart(drawingPart, imageBytes, contentType, deduplicate);
                string imageRelationshipId = drawingPart.GetIdOfPart(imagePart);

                UInt32Value drawingId = NextDrawingId(drawingPart);
                string resolvedName = string.IsNullOrWhiteSpace(name) ? $"Picture {drawingId}" : name!.Trim();
                long widthEmu = resolvedWidth.ToEmus();
                long heightEmu = resolvedHeight.ToEmus();
                var anchor = new Xdr.AbsoluteAnchor(
                    new Xdr.Position {
                        X = x.ToEmus(),
                        Y = y.ToEmus()
                    },
                    new Xdr.Extent { Cx = widthEmu, Cy = heightEmu },
                    CreatePicture(drawingId, resolvedName, imageRelationshipId, altText, title, lockAspectRatio, widthEmu, heightEmu, rotationDegrees),
                    new Xdr.ClientData());

                Xdr.WorksheetDrawing worksheetDrawing = drawingPart.WorksheetDrawing!;
                worksheetDrawing.Append(anchor);
                worksheetDrawing.Save();
                WorksheetRoot.Save();
                _excelDocument.MarkPackageDirty();
                image = new ExcelImage(anchor.GetFirstChild<Xdr.Picture>()!, anchor, drawingPart, this);
            });

            return image!;
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

        private static (OfficeLength Width, OfficeLength Height) ResolveImageSize(OfficeImageInfo info, OfficeLength? width, OfficeLength? height, double? scalePercent) {
            if (scalePercent.HasValue) {
                if (width.HasValue || height.HasValue)
                    throw new ArgumentException("Scale percentage cannot be combined with explicit width or height.");
                if (double.IsNaN(scalePercent.Value) || double.IsInfinity(scalePercent.Value) || scalePercent.Value <= 0)
                    throw new ArgumentOutOfRangeException(nameof(scalePercent), "Scale percentage must be a positive finite number.");
                if (info.Width <= 0 || info.Height <= 0)
                    throw new NotSupportedException("Image dimensions could not be detected, so scale percentage cannot be applied.");
                OfficeLength scaledWidth = OfficeLength.Pixels(info.Width) * (scalePercent.Value / 100D);
                OfficeLength scaledHeight = OfficeLength.Pixels(info.Height) * (scalePercent.Value / 100D);
                return (OfficeLength.Emus(Math.Max(1L, scaledWidth.ToEmus())), OfficeLength.Emus(Math.Max(1L, scaledHeight.ToEmus())));
            }
            OfficeLength resolvedWidth = width ?? OfficeLength.Pixels(info.Width > 0 ? info.Width : 96);
            OfficeLength resolvedHeight = height ?? OfficeLength.Pixels(info.Height > 0 ? info.Height : 32);
            if (resolvedWidth <= OfficeLength.Zero) throw new ArgumentOutOfRangeException(nameof(width));
            if (resolvedHeight <= OfficeLength.Zero) throw new ArgumentOutOfRangeException(nameof(height));
            return (resolvedWidth, resolvedHeight);
        }

        private static (int StartRow, int StartColumn, int EndRow, int EndColumn) ParseImageRange(string range) {
            if (!ExcelA1.TryParseRange(range, out int startRow, out int startColumn, out int endRow, out int endColumn)) {
                throw new ArgumentException($"Invalid A1 range '{range}'. Use a range such as A1:C15.", nameof(range));
            }

            if (endColumn >= ExcelA1.MaxColumns) {
                throw new ArgumentOutOfRangeException(nameof(range), "Image range must end before the final Excel column so the two-cell end marker can use the boundary after the last cell.");
            }

            if (endRow >= ExcelA1.MaxRows) {
                throw new ArgumentOutOfRangeException(nameof(range), "Image range must end before the final Excel row so the two-cell end marker can use the boundary after the last cell.");
            }

            return (startRow, startColumn, endRow, endColumn);
        }

        private (OfficeLength Width, OfficeLength Height) CalculateRangeAnchorSize(
            int startRow, int startColumn, int endRow, int endColumn,
            OfficeLength offsetX, OfficeLength offsetY, OfficeLength endOffsetX, OfficeLength endOffsetY) {
            OfficeLength width = OfficeLength.Zero;
            OfficeLength height = OfficeLength.Zero;
            var geometry = Utilities.ExcelWorksheetGeometryIndex.Create(WorksheetPart);
            for (int column = startColumn; column <= endColumn; column++)
                width += geometry.GetColumnLength(column);
            for (int row = startRow; row <= endRow; row++)
                height += geometry.GetRowLength(row);
            width += endOffsetX - offsetX;
            height += endOffsetY - offsetY;
            // Hidden-only or inverted spans retain a positive native extent without pixel quantization.
            return (width > OfficeLength.Zero ? width : OfficeLength.Emus(1),
                height > OfficeLength.Zero ? height : OfficeLength.Emus(1));
        }

        private static void ApplyImageMetadata(ExcelImage image, string? title, double rotationDegrees) {
            if (!string.IsNullOrWhiteSpace(title)) {
                image.Title = title!;
            }

            if (Math.Abs(rotationDegrees) > double.Epsilon) {
                image.SetRotation(rotationDegrees);
            }
        }

        /// <summary>Determines whether an image content type can be stored in an Excel image part.</summary>
        /// <param name="contentType">The MIME content type to inspect.</param>
        /// <returns><see langword="true"/> when the content type maps to an Excel image part; otherwise <see langword="false"/>.</returns>
        public static bool IsSupportedImageContentType(string? contentType) =>
            TryGetImagePartType(contentType, out _);

        private static PartTypeInfo ToImagePartType(string? contentType) {
            if (TryGetImagePartType(contentType, out PartTypeInfo imagePartType)) return imagePartType;
            throw new NotSupportedException($"Image content type '{contentType}' is not supported by Excel image parts.");
        }

        private static bool TryGetImagePartType(string? contentType, out PartTypeInfo imagePartType) {
            switch ((contentType ?? string.Empty).Trim().ToLowerInvariant()) {
                case "image/png": imagePartType = ImagePartType.Png; return true;
                case "image/jpeg":
                case "image/jpg": imagePartType = ImagePartType.Jpeg; return true;
                case "image/gif": imagePartType = ImagePartType.Gif; return true;
                case "image/bmp": imagePartType = ImagePartType.Bmp; return true;
                case "image/tiff":
                case "image/tif": imagePartType = ImagePartType.Tiff; return true;
                case "image/svg+xml":
                case "image/svg": imagePartType = ImagePartType.Svg; return true;
                case "image/x-emf":
                case "image/emf": imagePartType = ImagePartType.Emf; return true;
                case "image/x-wmf":
                case "image/wmf": imagePartType = ImagePartType.Wmf; return true;
                case "image/x-icon":
                case "image/vnd.microsoft.icon":
                case "image/ico": imagePartType = ImagePartType.Icon; return true;
                case "image/x-pcx":
                case "image/pcx": imagePartType = ImagePartType.Pcx; return true;
                default:
                    imagePartType = default!;
                    return false;
            }
        }

        private static string ResolveImageContentType(string? declaredContentType, OfficeImageInfo detectedInfo) {
            if (detectedInfo.Format != OfficeImageFormat.Unknown) {
                return detectedInfo.MimeType;
            }

            return string.IsNullOrWhiteSpace(declaredContentType) ? "application/octet-stream" : declaredContentType!;
        }

        private static OfficeImageInfo RequireImageContent(byte[] bytes, string? fileName, string parameterName) {
            if (!OfficeImageReader.TryValidateContent(bytes, fileName, out OfficeImageInfo info)) {
                throw new ArgumentException("Image bytes do not contain a complete supported image.", parameterName);
            }
            return info;
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
