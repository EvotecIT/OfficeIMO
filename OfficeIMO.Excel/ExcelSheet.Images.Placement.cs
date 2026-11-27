using OfficeIMO.Drawing;
using OfficeIMO.Excel.OpenXml;
using OfficeIMO.OpenXml.Internal;
using static OfficeIMO.Excel.OpenXml.ExcelChartXml;

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
                ExcelPartRef drawingPart = GetOrCreateDrawingsPart();
                OfficeXmlNode worksheetDrawing = drawingPart.Root;
                OfficeXmlNode anchor = worksheetDrawing.New(Xdr, "twoCellAnchor");
                OfficeXmlNode picture = anchor.New(Xdr, "pic");
                OfficeXmlNode blipFill = picture.New(Xdr, "blipFill");
                OfficeXmlNode blip = CreatePictureBlip(drawingPart, blipFill, imageBytes, contentType, deduplicate);

                uint drawingId = NextDrawingId(drawingPart);
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
                anchor.SetAttribute("editAs", ToEditAsValue(placement));
                anchor.With(
                    NewMarker(anchor, "from", startColumn - 1, offsetX.ToEmus(), startRow - 1, offsetY.ToEmus()),
                    NewMarker(anchor, "to", endColumn, endOffsetX.ToEmus(), endRow, endOffsetY.ToEmus()),
                    FillPicture(picture, blipFill, drawingId, resolvedName, blip, altText, title, lockAspectRatio,
                        width.ToEmus(), height.ToEmus(), rotationDegrees),
                    anchor.New(Xdr, "clientData"));

                anchor = worksheetDrawing.Append(anchor);
                drawingPart.SaveRoot();
                ExcelPartRef.SaveWorksheetElement(this);
                _excelDocument.MarkPackageDirty();
                image = new ExcelImage(anchor.Child(Xdr, "pic"), anchor, drawingPart, this);
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
                ExcelPartRef drawingPart = GetOrCreateDrawingsPart();
                OfficeXmlNode worksheetDrawing = drawingPart.Root;
                OfficeXmlNode anchor = worksheetDrawing.New(Xdr, "absoluteAnchor");
                OfficeXmlNode picture = anchor.New(Xdr, "pic");
                OfficeXmlNode blipFill = picture.New(Xdr, "blipFill");
                OfficeXmlNode blip = CreatePictureBlip(drawingPart, blipFill, imageBytes, contentType, deduplicate);

                uint drawingId = NextDrawingId(drawingPart);
                string resolvedName = string.IsNullOrWhiteSpace(name) ? $"Picture {drawingId}" : name!.Trim();
                long widthEmu = resolvedWidth.ToEmus();
                long heightEmu = resolvedHeight.ToEmus();
                anchor.With(
                    anchor.New(Xdr, "pos").Attr("x", x.ToEmus()).Attr("y", y.ToEmus()),
                    anchor.New(Xdr, "ext").Attr("cx", widthEmu).Attr("cy", heightEmu),
                    FillPicture(picture, blipFill, drawingId, resolvedName, blip, altText, title, lockAspectRatio, widthEmu, heightEmu, rotationDegrees),
                    anchor.New(Xdr, "clientData"));

                anchor = worksheetDrawing.Append(anchor);
                drawingPart.SaveRoot();
                ExcelPartRef.SaveWorksheetElement(this);
                _excelDocument.MarkPackageDirty();
                image = new ExcelImage(anchor.Child(Xdr, "pic"), anchor, drawingPart, this);
            });

            return image!;
        }

        /// <summary>
        /// The worksheet's drawing part, created and linked from the worksheet when the worksheet has none, with a
        /// worksheet drawing root.
        /// </summary>
        private ExcelPartRef GetOrCreateDrawingsPart() {
            ExcelPartRef worksheetPart = ExcelPartRef.WorksheetForMetadata(this);
            OfficeXmlNode drawing = ExcelPartRef.WorksheetDrawingElement(this);
            if (drawing.IsNone) {
                ExcelPartRef drawingPart = worksheetPart.AddNewPart(ExcelPartKind.Drawings);
                drawingPart.SetRootMarkup(WorksheetDrawingRootMarkup);
                string relationshipId = worksheetPart.GetIdOfPart(drawingPart)!;
                ExcelPartRef.LinkWorksheetDrawing(this, relationshipId);
                return drawingPart;
            }

            ExcelPartRef existing = RequireDrawingsPart(worksheetPart, drawing.RelationshipAttribute("id"));
            if (existing.Root.IsNone) existing.SetRootMarkup(WorksheetDrawingRootMarkup);
            return existing;
        }

        /// <summary>A new worksheet drawing root, as <c>new WorksheetDrawing()</c> writes it.</summary>
        private static byte[] WorksheetDrawingRootMarkup => System.Text.Encoding.UTF8.GetBytes(
            "<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" />");

        // The drawing part the worksheet's drawing relationship targets; fails as the typed part lookup does.
        private static ExcelPartRef RequireDrawingsPart(ExcelPartRef worksheetPart, string? relationshipId) {
            if (relationshipId == null) throw new ArgumentNullException("id");
            ExcelPartRef part = worksheetPart.GetPartById(relationshipId);
            if (part.IsNone) throw new ArgumentOutOfRangeException("id");
            if (part.Kind != ExcelPartKind.Drawings) throw new InvalidCastException("The worksheet drawing relationship does not target a drawing part.");
            return part;
        }

        // Fills a detached xdr:pic made for its anchor and returns it.
        private static OfficeXmlNode FillPicture(OfficeXmlNode picture, OfficeXmlNode blipFill, uint drawingId, string name, OfficeXmlNode blip,
            string? altText, string? title, bool lockAspectRatio, long widthEmu, long heightEmu, double rotationDegrees) {
            OfficeXmlNode nonVisual = picture.New(Xdr, "nvPicPr");
            OfficeXmlNode drawingProperties = nonVisual.New(Xdr, "cNvPr")
                .Attr("id", drawingId)
                .Attr("name", name)
                .Attr("descr", altText ?? string.Empty);
            if (!string.IsNullOrWhiteSpace(title)) {
                drawingProperties.SetAttribute("title", title);
            }

            OfficeXmlNode shapeProperties = picture.New(Xdr, "spPr");
            OfficeXmlNode transform = shapeProperties.New(A, "xfrm");
            transform.With(
                transform.New(A, "off").Attr("x", 0L).Attr("y", 0L),
                transform.New(A, "ext").Attr("cx", widthEmu).Attr("cy", heightEmu));
            if (Math.Abs(rotationDegrees) > double.Epsilon) {
                transform.SetInt32Attribute("rot", (int)Math.Round(rotationDegrees * 60000.0));
            }

            OfficeXmlNode pictureDrawing = nonVisual.New(Xdr, "cNvPicPr");
            OfficeXmlNode stretch = blipFill.New(A, "stretch");
            OfficeXmlNode geometry = shapeProperties.New(A, "prstGeom").Attr("prst", "rect");
            return picture.With(
                nonVisual.With(
                    drawingProperties,
                    pictureDrawing.With(pictureDrawing.New(A, "picLocks").Attr("noChangeAspect", lockAspectRatio))
                ),
                blipFill.With(
                    blip,
                    stretch.With(stretch.New(A, "fillRect"))
                ),
                shapeProperties.With(
                    transform,
                    geometry.With(geometry.New(A, "avLst"))
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
            var geometry = ExcelPartRef.Worksheet(this).CreateWorksheetGeometry();
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

        private static (string ContentType, string Extension) ToImagePartContentType(string? contentType) {
            if (TryGetImagePartType(contentType, out var imagePartType)) return imagePartType;
            throw new NotSupportedException($"Image content type '{contentType}' is not supported by Excel image parts.");
        }

        // The image part content type and file extension Excel image parts use for each accepted content type.
        private static bool TryGetImagePartType(string? contentType, out (string ContentType, string Extension) imagePartType) {
            switch ((contentType ?? string.Empty).Trim().ToLowerInvariant()) {
                case "image/png": imagePartType = ("image/png", ".png"); return true;
                case "image/jpeg":
                case "image/jpg": imagePartType = ("image/jpeg", ".jpg"); return true;
                case "image/gif": imagePartType = ("image/gif", ".gif"); return true;
                case "image/bmp": imagePartType = ("image/bmp", ".bmp"); return true;
                case "image/tiff":
                case "image/tif": imagePartType = ("image/tiff", ".tiff"); return true;
                case "image/svg+xml":
                case "image/svg": imagePartType = ("image/svg+xml", ".svg"); return true;
                case "image/x-emf":
                case "image/emf": imagePartType = ("image/x-emf", ".emf"); return true;
                case "image/x-wmf":
                case "image/wmf": imagePartType = ("image/x-wmf", ".wmf"); return true;
                case "image/x-icon":
                case "image/vnd.microsoft.icon":
                case "image/ico": imagePartType = ("image/x-icon", ".ico"); return true;
                case "image/x-pcx":
                case "image/pcx": imagePartType = ("image/x-pcx", ".pcx"); return true;
                default:
                    imagePartType = default;
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

        private static string ToEditAsValue(ExcelImagePlacement placement) {
            return placement switch {
                ExcelImagePlacement.MoveAndSize => "twoCell",
                ExcelImagePlacement.MoveOnly => "oneCell",
                ExcelImagePlacement.FreeFloating => "absolute",
                _ => "twoCell"
            };
        }
    }
}
