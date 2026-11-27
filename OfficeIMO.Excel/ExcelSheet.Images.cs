using System.Globalization;
using OfficeIMO.Drawing;
using OfficeIMO.Excel.OpenXml;
using OfficeIMO.OpenXml.Internal;
using static OfficeIMO.Excel.OpenXml.ExcelChartXml;

namespace OfficeIMO.Excel {
    // Helpers for inserting images anchored to worksheet cells.
    public partial class ExcelSheet {
        /// <summary>
        /// Enumerates worksheet images anchored in the drawing layer.
        /// </summary>
        public IEnumerable<ExcelImage> Images {
            get {
                ExcelPartRef drawingPart = ExcelPartRef.WorksheetForMetadata(this).Drawings;
                OfficeXmlNode worksheetDrawing = drawingPart.IsNone ? default : drawingPart.Root;
                if (worksheetDrawing.IsNone) return Enumerable.Empty<ExcelImage>();

                var images = new List<ExcelImage>();
                foreach (OfficeXmlNode anchor in worksheetDrawing.Children()) {
                    if (!IsSupportedImageAnchor(anchor)) continue;
                    foreach (OfficeXmlNode picture in anchor.Descendants(Xdr, "pic")) {
                        images.Add(new ExcelImage(picture, anchor, drawingPart, this));
                    }
                }
                return images;
            }
        }

        /// <summary>
        /// Returns an image by non-visual drawing name, or null if it was not found.
        /// </summary>
        public ExcelImage? GetImage(string name) {
            if (string.IsNullOrWhiteSpace(name)) return null;
            return Images.FirstOrDefault(image => string.Equals(image.Name, name, StringComparison.OrdinalIgnoreCase));
        }

        // Writes a native one-cell anchor; source loading and options live in the authoring partials.
        private ExcelImage AddCellImageCore(int row, int column, byte[] imageBytes, string contentType = "image/png",
            OfficeLength? width = null, OfficeLength? height = null, OfficeLength offsetX = default, OfficeLength offsetY = default,
            string? name = null, string? altText = null, bool lockAspectRatio = true, bool deduplicate = false) {
            if (imageBytes == null || imageBytes.Length == 0) throw new ArgumentException("Image bytes are required.", nameof(imageBytes));
            if (row <= 0 || column <= 0) throw new ArgumentOutOfRangeException("Row and column are 1-based and must be positive.");
            OfficeLength resolvedWidth = width ?? OfficeLength.Pixels(96);
            OfficeLength resolvedHeight = height ?? OfficeLength.Pixels(32);
            if (resolvedWidth <= OfficeLength.Zero) throw new ArgumentOutOfRangeException(nameof(width));
            if (resolvedHeight <= OfficeLength.Zero) throw new ArgumentOutOfRangeException(nameof(height));
            ExcelImage? image = null;
            WriteLock(() => {
                // Ensure a drawing part exists and is referenced by the worksheet
                ExcelPartRef drawingPart = GetOrCreateDrawingsPart();
                OfficeXmlNode worksheetDrawing = drawingPart.Root;

                // Add the image part
                OfficeXmlNode anchor = worksheetDrawing.New(Xdr, "oneCellAnchor");
                OfficeXmlNode picture = anchor.New(Xdr, "pic");
                OfficeXmlNode blipFill = picture.New(Xdr, "blipFill");
                OfficeXmlNode blip = CreatePictureBlip(drawingPart, blipFill, imageBytes, contentType, deduplicate);

                // Build a OneCellAnchor: FromMarker + Extent + Picture + ClientData
                long cx = resolvedWidth.ToEmus();
                long cy = resolvedHeight.ToEmus();
                long dx = offsetX.ToEmus();
                long dy = offsetY.ToEmus();

                uint nvId = NextDrawingId(drawingPart);
                string resolvedName = string.IsNullOrWhiteSpace(name) ? $"Picture {nvId}" : name!.Trim();
                OfficeXmlNode nonVisual = picture.New(Xdr, "nvPicPr");
                OfficeXmlNode pictureDrawing = nonVisual.New(Xdr, "cNvPicPr");
                OfficeXmlNode stretch = blipFill.New(A, "stretch");
                OfficeXmlNode shapeProperties = picture.New(Xdr, "spPr");
                OfficeXmlNode transform = shapeProperties.New(A, "xfrm");
                OfficeXmlNode geometry = shapeProperties.New(A, "prstGeom").Attr("prst", "rect");
                anchor.With(
                    NewMarker(anchor, "from", column - 1, dx, row - 1, dy),
                    anchor.New(Xdr, "ext").Attr("cx", cx).Attr("cy", cy),
                    picture.With(
                        nonVisual.With(
                            nonVisual.New(Xdr, "cNvPr").Attr("id", nvId).Attr("name", resolvedName).Attr("descr", altText ?? string.Empty),
                            pictureDrawing.With(pictureDrawing.New(A, "picLocks").Attr("noChangeAspect", lockAspectRatio))
                        ),
                        blipFill.With(
                            blip,
                            stretch.With(stretch.New(A, "fillRect"))
                        ),
                        shapeProperties.With(
                            transform.With(
                                transform.New(A, "off").Attr("x", 0L).Attr("y", 0L),
                                transform.New(A, "ext").Attr("cx", cx).Attr("cy", cy)
                            ),
                            geometry.With(geometry.New(A, "avLst"))
                        )
                    ),
                    anchor.New(Xdr, "clientData")
                );

                anchor = worksheetDrawing.Append(anchor);
                drawingPart.SaveRoot();
                ExcelPartRef.SaveWorksheetElement(this);
                _excelDocument.MarkPackageDirty();
                image = new ExcelImage(anchor.Child(Xdr, "pic"), anchor, drawingPart, this);
            });

            return image!;
        }

        private static long PxToEmu(int px) => (long)Math.Round(px * 9525.0);

        private static bool IsSupportedImageAnchor(OfficeXmlNode anchor)
            => IsSupportedDrawingAnchor(anchor) && anchor.FirstDescendant(Xdr, "pic").Exists;

        /// <summary>An <c>xdr:from</c> or <c>xdr:to</c> marker with its column, column offset, row and row offset.</summary>
        private static OfficeXmlNode NewMarker(OfficeXmlNode context, string localName, long column, long columnOffset, long row, long rowOffset) {
            OfficeXmlNode marker = context.New(Xdr, localName);
            return marker.With(
                marker.NewText(Xdr, "col", column.ToString(CultureInfo.InvariantCulture)),
                marker.NewText(Xdr, "colOff", columnOffset.ToString(CultureInfo.InvariantCulture)),
                marker.NewText(Xdr, "row", row.ToString(CultureInfo.InvariantCulture)),
                marker.NewText(Xdr, "rowOff", rowOffset.ToString(CultureInfo.InvariantCulture)));
        }

        private static uint NextDrawingId(ExcelPartRef dp) {
            OfficeXmlNode worksheetDrawing = dp.Root;
            if (worksheetDrawing.IsNone) return 1;
            // An SDK-held drawing is walked through the SDK's typed elements, which each chart and image added scans.
            if (worksheetDrawing.IsSdk) return dp.MaxSdkDrawingId() + 1;
            uint max = 0;
            MaxDrawingId(worksheetDrawing, ref max);
            return max + 1;
        }

        private static void MaxDrawingId(OfficeXmlNode element, ref uint max) {
            for (OfficeXmlNode child = element.FirstChild; child.Exists; child = child.NextSibling) {
                if (child.Is(Xdr, "cNvPr")) {
                    uint? id = child.UInt32Attribute("id");
                    if (id.HasValue && id.Value > max) max = id.Value;
                }
                MaxDrawingId(child, ref max);
            }
        }
    }
}
