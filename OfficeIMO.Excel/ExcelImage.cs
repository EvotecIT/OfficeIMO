using OfficeIMO.Excel.OpenXml;
using OfficeIMO.Excel.Utilities;
using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;
using System.Globalization;
using static OfficeIMO.Excel.OpenXml.ExcelChartXml;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Represents a worksheet image anchored in the drawing layer.
    /// </summary>
    public sealed partial class ExcelImage {
        private readonly ExcelDrawingHandle _drawingHandle;
        /// <summary>The image's <c>xdr:pic</c> in the worksheet drawing.</summary>
        private OfficeXmlNode PictureNode => _drawingHandle.Element;
        /// <summary>The anchor (one-cell, two-cell or absolute) that holds the picture.</summary>
        private OfficeXmlNode AnchorNode {
            get {
                for (var node = PictureNode.Parent; !node.IsNone; node = node.Parent) {
                    if (node.Is(Xdr, "oneCellAnchor") || node.Is(Xdr, "twoCellAnchor") || node.Is(Xdr, "absoluteAnchor")) return node;
                }
                throw new InvalidOperationException("The picture is not in a drawing anchor.");
            }
        }
        /// <summary>The worksheet drawing part that holds the picture.</summary>
        private ExcelPartRef DrawingsPartRef => _drawingHandle.Part;
        private readonly ExcelDocument _document;

        internal ExcelImage(OfficeXmlNode picture, OfficeXmlNode anchor, ExcelPartRef drawingsPart, ExcelSheet sheet) {
            if (anchor.IsNone) throw new ArgumentNullException(nameof(anchor));
            _drawingHandle = new ExcelDrawingHandle(picture, drawingsPart, sheet);
            _document = sheet.Document;
        }

        /// <summary>
        /// Gets or sets the image drawing name.
        /// </summary>
        public string Name {
            get => DrawingProperties.Attribute("name") ?? string.Empty;
            set {
                OfficeXmlNode properties = DrawingProperties;
                if (properties.Exists) {
                    properties.SetAttribute("name", value ?? string.Empty);
                    Save();
                }
            }
        }

        /// <summary>
        /// Gets or sets the image title metadata.
        /// </summary>
        public string Title {
            get => DrawingProperties.Attribute("title") ?? string.Empty;
            set {
                OfficeXmlNode properties = DrawingProperties;
                if (properties.Exists) {
                    properties.SetAttribute("title", value ?? string.Empty);
                    Save();
                }
            }
        }

        /// <summary>
        /// Gets or sets DrawingML picture transparency from 0 through 100 percent.
        /// Null removes the explicit alpha transform.
        /// </summary>
        public int? TransparencyPercent {
            get {
                int? amount = Blip.Child(A, "alphaModFix").Int32Attribute("amt");
                return amount.HasValue
                    ? (int)Math.Round((100000 - amount.Value) / 1000D)
                    : null;
            }
            set {
                if (value is < 0 or > 100) {
                    throw new ArgumentOutOfRangeException(nameof(value), "Transparency must be between 0 and 100.");
                }
                OfficeXmlNode blip = Blip;
                if (blip.IsNone) return;
                OfficeXmlNode alpha = blip.Child(A, "alphaModFix");
                if (!value.HasValue) {
                    if (alpha.Exists) alpha.Remove();
                    Save();
                    return;
                }
                foreach (OfficeXmlNode transform in blip.Children().ToList()) {
                    if (transform.LocalName.StartsWith("alpha", StringComparison.Ordinal)) transform.Remove();
                }
                alpha = blip.New(A, "alphaModFix").Attr("amt", checked((int)Math.Round((100D - value.Value) * 1000D,
                    MidpointRounding.AwayFromZero)));
                OfficeXmlNode extensions = blip.Child(A, "extLst");
                if (extensions.Exists) blip.InsertBefore(alpha, extensions);
                else blip.Append(alpha);
                Save();
            }
        }

        /// <summary>
        /// Gets or sets the image alternative text description.
        /// </summary>
        public string Description {
            get => DrawingProperties.Attribute("descr") ?? string.Empty;
            set {
                OfficeXmlNode properties = DrawingProperties;
                if (properties.Exists) {
                    properties.SetAttribute("descr", value ?? string.Empty);
                    Save();
                }
            }
        }

        /// <summary>
        /// Gets whether Excel should keep the picture aspect ratio locked.
        /// </summary>
        public bool IsAspectRatioLocked =>
            PictureDrawingProperties.Child(A, "picLocks").BooleanAttribute("noChangeAspect") ?? false;

        /// <summary>
        /// Gets the 1-based row index where the image is anchored, when available.
        /// </summary>
        public int RowIndex => GetMarkerRow() + 1;

        /// <summary>
        /// Gets the 1-based column index where the image is anchored, when available.
        /// </summary>
        public int ColumnIndex => GetMarkerColumn() + 1;

        /// <summary>
        /// Gets the image width as a document length from the drawing extent.
        /// </summary>
        public OfficeLength Width => OfficeLength.Emus(GetExtentCx());

        /// <summary>
        /// Gets the image height as a document length from the drawing extent.
        /// </summary>
        public OfficeLength Height => OfficeLength.Emus(GetExtentCy());

        /// <summary>
        /// Gets the image horizontal offset from the anchor cell origin as a document length.
        /// </summary>
        public OfficeLength OffsetX => OfficeLength.Emus(GetMarkerColumnOffset());

        /// <summary>
        /// Gets the image vertical offset from the anchor cell origin as a document length.
        /// </summary>
        public OfficeLength OffsetY => OfficeLength.Emus(GetMarkerRowOffset());

        /// <summary>
        /// Gets whether the image is positioned with a two-cell drawing anchor.
        /// </summary>
        public bool HasTwoCellAnchor => AnchorNode.Is(Xdr, "twoCellAnchor");

        /// <summary>
        /// Gets whether the image is positioned with an absolute drawing anchor.
        /// </summary>
        public bool HasAbsoluteAnchor => AnchorNode.Is(Xdr, "absoluteAnchor");

        internal ExcelImagePlacement Placement {
            get {
                OfficeXmlNode anchor = AnchorNode;
                if (!anchor.Is(Xdr, "twoCellAnchor")) return ExcelImagePlacement.FreeFloating;
                string? editAs = anchor.Attribute("editAs");
                if (editAs == "oneCell") return ExcelImagePlacement.MoveOnly;
                if (editAs == "absolute") return ExcelImagePlacement.FreeFloating;
                return ExcelImagePlacement.MoveAndSize;
            }
        }

        /// <summary>
        /// Gets the 1-based ending row index for two-cell anchored images, when available.
        /// </summary>
        public int? ToRowIndex => TryGetToMarkerRow(out int row) ? row + 1 : null;

        /// <summary>
        /// Gets the 1-based ending column index for two-cell anchored images, when available.
        /// </summary>
        public int? ToColumnIndex => TryGetToMarkerColumn(out int column) ? column + 1 : null;

        /// <summary>
        /// Gets the horizontal offset from the two-cell ending marker column as a document length.
        /// </summary>
        public OfficeLength ToOffsetX => OfficeLength.Emus(GetToMarkerColumnOffset());

        /// <summary>
        /// Gets the vertical offset from the two-cell ending marker row as a document length.
        /// </summary>
        public OfficeLength ToOffsetY => OfficeLength.Emus(GetToMarkerRowOffset());

        /// <summary>
        /// Tries to read the image bounds when it is positioned by an absolute worksheet drawing anchor.
        /// </summary>
        public bool TryGetAbsoluteAnchorBounds(out OfficeLength x, out OfficeLength y, out OfficeLength width, out OfficeLength height) {
            x = y = width = height = OfficeLength.Zero;
            OfficeXmlNode absoluteAnchor = AnchorNode;
            if (!absoluteAnchor.Is(Xdr, "absoluteAnchor")) return false;
            OfficeXmlNode position = absoluteAnchor.Child(Xdr, "pos");
            OfficeXmlNode extent = absoluteAnchor.Child(Xdr, "ext");
            x = OfficeLength.Emus(position.Int64Attribute("x") ?? 0L);
            y = OfficeLength.Emus(position.Int64Attribute("y") ?? 0L);
            width = OfficeLength.Emus(extent.Int64Attribute("cx") ?? 0L);
            height = OfficeLength.Emus(extent.Int64Attribute("cy") ?? 0L);
            if (width <= OfficeLength.Zero || height <= OfficeLength.Zero) {
                width = Width;
                height = Height;
            }
            return width > OfficeLength.Zero && height > OfficeLength.Zero;
        }

        /// <summary>
        /// Sets the ending marker for a two-cell anchored image using 1-based worksheet marker coordinates.
        /// </summary>
        /// <param name="toRow">1-based ending row marker.</param>
        /// <param name="toColumn">1-based ending column marker.</param>
        /// <param name="toOffsetX">Horizontal offset from the ending marker column.</param>
        /// <param name="toOffsetY">Vertical offset from the ending marker row.</param>
        public ExcelImage SetTwoCellEndingMarker(int toRow, int toColumn, OfficeLength toOffsetX = default, OfficeLength toOffsetY = default) {
            ArgumentOutOfRangeException.ThrowIfLessThan(toRow, 1);
            ArgumentOutOfRangeException.ThrowIfLessThan(toColumn, 1);
            if (toOffsetX.ToEmus() < 0) throw new ArgumentOutOfRangeException(nameof(toOffsetX));
            if (toOffsetY.ToEmus() < 0) throw new ArgumentOutOfRangeException(nameof(toOffsetY));
            return SetNativeTwoCellEndingMarker(toRow, toColumn, toOffsetX, toOffsetY);
        }

        /// <summary>Restores signed native marker offsets after an importing adapter has bounded its geometry.</summary>
        internal ExcelImage SetNativeTwoCellEndingMarker(int toRow, int toColumn, OfficeLength toOffsetX, OfficeLength toOffsetY) {
            ArgumentOutOfRangeException.ThrowIfLessThan(toRow, 1);
            ArgumentOutOfRangeException.ThrowIfLessThan(toColumn, 1);
            OfficeXmlNode anchor = AnchorNode;
            OfficeXmlNode toMarker = anchor.Is(Xdr, "twoCellAnchor") ? anchor.Child(Xdr, "to") : default;
            if (toMarker.IsNone) {
                throw new InvalidOperationException("Only two-cell anchored images expose an ending marker.");
            }

            SetMarkerText(toMarker, "row", (toRow - 1).ToString(CultureInfo.InvariantCulture));
            SetMarkerText(toMarker, "col", (toColumn - 1).ToString(CultureInfo.InvariantCulture));
            SetMarkerText(toMarker, "rowOff", toOffsetY.ToEmus().ToString(CultureInfo.InvariantCulture));
            SetMarkerText(toMarker, "colOff", toOffsetX.ToEmus().ToString(CultureInfo.InvariantCulture));
            Save();
            return this;
        }

        /// <summary>
        /// Gets the left crop ratio authored on the image, where 0 means no crop and 1 means the full width.
        /// </summary>
        public double CropLeftRatio => GetCropRatio(CropRectangle.Int32Attribute("l"));

        /// <summary>
        /// Gets the top crop ratio authored on the image, where 0 means no crop and 1 means the full height.
        /// </summary>
        public double CropTopRatio => GetCropRatio(CropRectangle.Int32Attribute("t"));

        /// <summary>
        /// Gets the right crop ratio authored on the image, where 0 means no crop and 1 means the full width.
        /// </summary>
        public double CropRightRatio => GetCropRatio(CropRectangle.Int32Attribute("r"));

        /// <summary>
        /// Gets the bottom crop ratio authored on the image, where 0 means no crop and 1 means the full height.
        /// </summary>
        public double CropBottomRatio => GetCropRatio(CropRectangle.Int32Attribute("b"));

        /// <summary>
        /// Gets or sets clockwise image rotation in degrees.
        /// </summary>
        public double RotationDegrees {
            get => GetRotationDegrees(Transform.Int32Attribute("rot"));
            set => SetRotation(value);
        }

        /// <summary>
        /// Gets whether the image has an authored horizontal flip transform.
        /// </summary>
        public bool FlipHorizontal => Transform.BooleanAttribute("flipH") ?? false;

        /// <summary>
        /// Gets whether the image has an authored vertical flip transform.
        /// </summary>
        public bool FlipVertical => Transform.BooleanAttribute("flipV") ?? false;

        /// <summary>
        /// Gets this image anchor's zero-based order in the worksheet drawing layer.
        /// </summary>
        public int DrawingOrder => GetDrawingOrder(AnchorNode, DrawingsPartRef);

        /// <summary>
        /// Gets the image content type, such as image/png or image/jpeg. An SVG picture reports its SVG, not the PNG fallback.
        /// </summary>
        public string ContentType {
            get {
                ExcelPartRef imagePart = ImagePart;
                return imagePart.IsNone ? string.Empty : imagePart.ContentType;
            }
        }

        /// <summary>
        /// Returns a copy of the image bytes from the worksheet drawing relationship. An SVG picture returns its SVG.
        /// </summary>
        public byte[] ToBytes() {
            ExcelPartRef imagePart = ImagePart;
            if (imagePart.IsNone) {
                return Array.Empty<byte>();
            }

            return imagePart.ReadAllBytes();
        }

        internal bool TryReadBytes(out byte[] bytes) =>
            TryReadBytes(ExcelImageExportOptions.DefaultMaximumTotalSourceImageBytes, out bytes);

        internal bool TryReadBytes(long remainingAggregateBytes, out byte[] bytes) {
            ExcelPartRef imagePart = ImagePart;
            if (imagePart.IsNone) {
                bytes = Array.Empty<byte>();
                return false;
            }

            using Stream source = imagePart.OpenRead();
            return ExcelImageExportLimits.TryReadSourceImageBytes(source, remainingAggregateBytes, out bytes);
        }

        /// <summary>
        /// Sets image title and description metadata.
        /// </summary>
        public ExcelImage SetAltText(string description, string? title = null) {
            OfficeXmlNode properties = DrawingProperties;
            if (properties.Exists) {
                properties.SetAttribute("descr", description ?? string.Empty);
                ExcelDrawingAccessibility.ClearDecorative(properties);
                if (title != null) {
                    properties.SetAttribute("title", title);
                }
                Save();
            }

            return this;
        }

        /// <summary>Gets whether the image is marked decorative, so screen readers and accessibility checks skip it.</summary>
        public bool IsDecorative => ExcelDrawingAccessibility.IsDecorative(DrawingProperties);

        /// <summary>
        /// Marks the image as decorative: clears its alt text metadata and sets the Office decorative flag,
        /// so Excel and accessibility checkers skip it.
        /// </summary>
        public ExcelImage Decorative() {
            OfficeXmlNode properties = DrawingProperties;
            if (properties.Exists) {
                ExcelDrawingAccessibility.MarkDecorative(properties);
                Save();
            }

            return this;
        }

        /// <summary>
        /// Controls whether Excel should keep the picture aspect ratio locked.
        /// </summary>
        public ExcelImage LockAspectRatio(bool locked = true) {
            OfficeXmlNode drawingProps = PictureDrawingProperties;
            if (drawingProps.IsNone) {
                return this;
            }

            OfficeXmlNode locks = drawingProps.Child(A, "picLocks");
            if (locks.IsNone) {
                locks = drawingProps.Append(drawingProps.New(A, "picLocks"));
            }

            locks.SetBooleanAttribute("noChangeAspect", locked);
            Save();
            return this;
        }

        /// <summary>
        /// Sets clockwise image rotation in degrees.
        /// </summary>
        /// <param name="degrees">Clockwise rotation in degrees. Fractional values are supported by DrawingML.</param>
        public ExcelImage SetRotation(double degrees) {
            if (double.IsNaN(degrees) || double.IsInfinity(degrees)) {
                throw new ArgumentOutOfRangeException(nameof(degrees), "Rotation must be a finite number of degrees.");
            }

            OfficeXmlNode transform = EnsureTransform();
            if (transform.IsNone) return this;
            transform.SetInt32Attribute("rot", (int)Math.Round(degrees * 60000.0));
            Save();
            return this;
        }

        /// <summary>
        /// Sets image crop ratios, where 0 means no crop and 1 means the full source dimension.
        /// </summary>
        public ExcelImage SetCropRatio(double left, double top, double right, double bottom) {
            ValidateCropRatio(left, nameof(left));
            ValidateCropRatio(top, nameof(top));
            ValidateCropRatio(right, nameof(right));
            ValidateCropRatio(bottom, nameof(bottom));

            OfficeXmlNode blipFill = PictureNode.Child(Xdr, "blipFill");
            if (blipFill.IsNone) {
                return this;
            }

            if (left == 0D && top == 0D && right == 0D && bottom == 0D) {
                OfficeXmlNode existing = blipFill.Child(A, "srcRect");
                if (existing.Exists) existing.Remove();
                Save();
                return this;
            }

            OfficeXmlNode rectangle = blipFill.Child(A, "srcRect");
            if (rectangle.IsNone) {
                rectangle = blipFill.New(A, "srcRect");
                OfficeXmlNode blip = blipFill.Child(A, "blip");
                rectangle = blip.Exists ? blip.InsertAfterSelf(rectangle) : blipFill.InsertBefore(rectangle, blipFill.FirstChild);
            }

            rectangle.SetInt32Attribute("l", ToCropPercentage(left));
            rectangle.SetInt32Attribute("t", ToCropPercentage(top));
            rectangle.SetInt32Attribute("r", ToCropPercentage(right));
            rectangle.SetInt32Attribute("b", ToCropPercentage(bottom));
            Save();
            return this;
        }

        /// <summary>
        /// Sets image flip transforms.
        /// </summary>
        public ExcelImage SetFlip(bool horizontal, bool vertical) {
            OfficeXmlNode transform = EnsureTransform();
            if (transform.IsNone) return this;
            transform.SetBooleanAttribute("flipH", horizontal);
            transform.SetBooleanAttribute("flipV", vertical);
            Save();
            return this;
        }

        /// <summary>
        /// Sets the image size as a document length.
        /// </summary>
        public ExcelImage SetSize(OfficeLength width, OfficeLength height) {
            if (width <= OfficeLength.Zero) throw new ArgumentOutOfRangeException(nameof(width));
            if (height <= OfficeLength.Zero) throw new ArgumentOutOfRangeException(nameof(height));

            long cx = width.ToEmus();
            long cy = height.ToEmus();

            OfficeXmlNode anchor = AnchorNode;
            OfficeXmlNode extent = anchor.Child(Xdr, "ext");
            if (extent.Exists) {
                extent.SetInt64Attribute("cx", cx);
                extent.SetInt64Attribute("cy", cy);
            } else if (anchor.Is(Xdr, "twoCellAnchor")) {
                ResizeTwoCellAnchor(anchor, width, height);
            }

            OfficeXmlNode transformExtents = Transform.Child(A, "ext");
            if (transformExtents.Exists) {
                transformExtents.SetInt64Attribute("cx", cx);
                transformExtents.SetInt64Attribute("cy", cy);
            }

            Save();
            return this;
        }

        private void ResizeTwoCellAnchor(OfficeXmlNode anchor, OfficeLength width, OfficeLength height) {
            OfficeXmlNode fromMarker = anchor.Child(Xdr, "from");
            OfficeXmlNode toMarker = anchor.Child(Xdr, "to");
            if (fromMarker.IsNone || toMarker.IsNone) {
                return;
            }

            ExcelWorksheetGeometryIndex geometry = DrawingsPartRef.ParentWorksheet.CreateWorksheetGeometry();
            int startColumn = ParseMarkerIndex(fromMarker.ChildText(Xdr, "col"));
            int startRow = ParseMarkerIndex(fromMarker.ChildText(Xdr, "row"));
            var columnMarker = ResolveEndMarker(
                startColumn,
                ParseMarkerOffset(fromMarker.ChildText(Xdr, "colOff")),
                width.ToEmus(),
                ExcelA1.MaxColumns,
                index => geometry.GetColumnLength(index + 1).ToEmus());
            var rowMarker = ResolveEndMarker(
                startRow,
                ParseMarkerOffset(fromMarker.ChildText(Xdr, "rowOff")),
                height.ToEmus(),
                ExcelA1.MaxRows,
                index => geometry.GetRowLength(index + 1).ToEmus());

            SetMarkerText(toMarker, "col", columnMarker.Index.ToString(CultureInfo.InvariantCulture));
            SetMarkerText(toMarker, "colOff", columnMarker.OffsetEmu.ToString(CultureInfo.InvariantCulture));
            SetMarkerText(toMarker, "row", rowMarker.Index.ToString(CultureInfo.InvariantCulture));
            SetMarkerText(toMarker, "rowOff", rowMarker.OffsetEmu.ToString(CultureInfo.InvariantCulture));
        }

        /// <summary>
        /// Scales the image from its current size by the specified percentage.
        /// </summary>
        /// <param name="percent">Percentage of the current image size. For example, 20 means 20%.</param>
        public ExcelImage SetSizePercent(double percent) {
            if (double.IsNaN(percent) || double.IsInfinity(percent) || percent <= 0) {
                throw new ArgumentOutOfRangeException(nameof(percent), "Scale percentage must be a positive finite number.");
            }

            OfficeLength width = Width * (percent / 100D);
            OfficeLength height = Height * (percent / 100D);
            if (width <= OfficeLength.Zero) width = OfficeLength.Emus(1);
            if (height <= OfficeLength.Zero) height = OfficeLength.Emus(1);
            return SetSize(width, height);
        }

        /// <summary>
        /// Moves the image to a one-cell anchor position and optionally applies length offsets.
        /// </summary>
        /// <param name="row">1-based target row.</param>
        /// <param name="column">1-based target column.</param>
        /// <param name="offsetX">Horizontal offset from the cell edge, as a document length.</param>
        /// <param name="offsetY">Vertical offset from the cell edge, as a document length.</param>
        public ExcelImage MoveTo(int row, int column, OfficeLength offsetX = default, OfficeLength offsetY = default) {
            if (row <= 0 || row > ExcelA1.MaxRows) throw new ArgumentOutOfRangeException(nameof(row), "Row must be between 1 and the Excel row limit.");
            if (column <= 0 || column > ExcelA1.MaxColumns) throw new ArgumentOutOfRangeException(nameof(column), "Column must be between 1 and the Excel column limit.");

            OfficeXmlNode marker = AnchorNode.Child(Xdr, "from");
            if (marker.IsNone) {
                throw new NotSupportedException("Only images anchored to worksheet cells can be moved. Absolute image anchors do not have a cell marker.");
            }

            OfficeLength width = Width, height = Height;
            int previousColumn = ParseMarkerIndex(marker.ChildText(Xdr, "col"));
            int previousRow = ParseMarkerIndex(marker.ChildText(Xdr, "row"));
            long previousColumnOffset = ParseMarkerOffset(marker.ChildText(Xdr, "colOff"));
            long previousRowOffset = ParseMarkerOffset(marker.ChildText(Xdr, "rowOff"));
            int targetColumn = column - 1;
            int targetRow = row - 1;
            long targetColumnOffset = offsetX.ToEmus();
            long targetRowOffset = offsetY.ToEmus();

            MoveTwoCellEndMarkerIfNeeded(
                targetColumn - previousColumn,
                targetRow - previousRow,
                (decimal)targetColumnOffset - previousColumnOffset,
                (decimal)targetRowOffset - previousRowOffset);

            SetMarkerText(marker, "col", (column - 1).ToString(CultureInfo.InvariantCulture));
            SetMarkerText(marker, "colOff", targetColumnOffset.ToString(CultureInfo.InvariantCulture));
            SetMarkerText(marker, "row", (row - 1).ToString(CultureInfo.InvariantCulture));
            SetMarkerText(marker, "rowOff", targetRowOffset.ToString(CultureInfo.InvariantCulture));
            // Columns and rows differ in size, so the end marker is placed from the image's size at the new origin.
            OfficeXmlNode anchor = AnchorNode;
            if (anchor.Is(Xdr, "twoCellAnchor") && width > OfficeLength.Zero && height > OfficeLength.Zero) {
                ResizeTwoCellAnchor(anchor, width, height);
            }
            Save();
            return this;
        }

        private void MoveTwoCellEndMarkerIfNeeded(int columnDelta, int rowDelta, decimal columnOffsetDelta, decimal rowOffsetDelta) {
            OfficeXmlNode toMarker = AnchorNode.Child(Xdr, "to");
            if (toMarker.IsNone) {
                return;
            }

            int targetColumn = ParseMarkerIndex(toMarker.ChildText(Xdr, "col")) + columnDelta;
            int targetRow = ParseMarkerIndex(toMarker.ChildText(Xdr, "row")) + rowDelta;
            if (targetColumn < 0 || targetColumn >= ExcelA1.MaxColumns) {
                throw new ArgumentOutOfRangeException("column", "Moving this two-cell image would place its end marker outside the Excel column limit.");
            }

            if (targetRow < 0 || targetRow >= ExcelA1.MaxRows) {
                throw new ArgumentOutOfRangeException("row", "Moving this two-cell image would place its end marker outside the Excel row limit.");
            }

            long targetOffsetX = checked((long)(ParseMarkerOffset(toMarker.ChildText(Xdr, "colOff")) + columnOffsetDelta));
            long targetOffsetY = checked((long)(ParseMarkerOffset(toMarker.ChildText(Xdr, "rowOff")) + rowOffsetDelta));
            SetMarkerText(toMarker, "col", targetColumn.ToString(CultureInfo.InvariantCulture));
            SetMarkerText(toMarker, "colOff", targetOffsetX.ToString(CultureInfo.InvariantCulture));
            SetMarkerText(toMarker, "row", targetRow.ToString(CultureInfo.InvariantCulture));
            SetMarkerText(toMarker, "rowOff", targetOffsetY.ToString(CultureInfo.InvariantCulture));
        }

        // A typed marker property setter: replaces the child, or adds it in schema order.
        private static void SetMarkerText(OfficeXmlNode marker, string localName, string text) =>
            marker.SetOrdered(marker.NewText(Xdr, localName, text), ExcelDrawingXml.MarkerOrder);

        private OfficeXmlNode DrawingProperties
            => PictureNode.Child(Xdr, "nvPicPr").Child(Xdr, "cNvPr");

        private OfficeXmlNode PictureDrawingProperties
            => PictureNode.Child(Xdr, "nvPicPr").Child(Xdr, "cNvPicPr");

        private OfficeXmlNode Blip
            => PictureNode.Child(Xdr, "blipFill").Child(A, "blip");

        private OfficeXmlNode CropRectangle
            => PictureNode.Child(Xdr, "blipFill").Child(A, "srcRect");

        private OfficeXmlNode Transform
            => PictureNode.Child(Xdr, "spPr").Child(A, "xfrm");

        private OfficeXmlNode EnsureTransform() {
            OfficeXmlNode shapeProperties = PictureNode.Child(Xdr, "spPr");
            if (shapeProperties.IsNone) {
                return default;
            }

            OfficeXmlNode transform = shapeProperties.Child(A, "xfrm");
            if (transform.IsNone) {
                transform = shapeProperties.New(A, "xfrm");
                transform.With(
                    transform.New(A, "off").Attr("x", 0L).Attr("y", 0L),
                    transform.New(A, "ext").Attr("cx", GetExtentCx()).Attr("cy", GetExtentCy()));
                transform = shapeProperties.InsertBefore(transform, shapeProperties.FirstChild);
            }

            return transform;
        }

        // A picture with an SVG extension is that SVG; its ordinary blip is only the raster fallback.
        private ExcelPartRef ImagePart {
            get {
                ExcelPartRef svg = ResolveImagePart(GetSvgRelationshipId());
                return svg.Exists ? svg : FallbackImagePart;
            }
        }

        private ExcelPartRef FallbackImagePart => ResolveImagePart(Blip.RelationshipAttribute("embed"));

        private string? GetSvgRelationshipId() {
            foreach (OfficeXmlNode vector in ExcelDrawingXml.Descendants(Blip, SvgUri, "svgBlip")) {
                string? id = vector.RelationshipAttribute("embed");
                if (!string.IsNullOrEmpty(id)) return id;
            }
            return null;
        }

        private const string SvgUri = "http://schemas.microsoft.com/office/drawing/2016/SVG/main";

        /// <summary>Returns the raster fallback of an SVG picture, for renderers that cannot draw its SVG.</summary>
        internal bool TryGetRasterFallback(out byte[] bytes, out string contentType) {
            bytes = Array.Empty<byte>();
            contentType = string.Empty;
            if (GetSvgRelationshipId() == null) return false;
            ExcelPartRef fallback = FallbackImagePart;
            if (fallback.IsNone) return false;
            bytes = fallback.ReadAllBytes();
            contentType = fallback.ContentType;
            return bytes.Length > 0;
        }

        private ExcelPartRef ResolveImagePart(string? relationshipId) {
            if (string.IsNullOrWhiteSpace(relationshipId)) {
                return default;
            }

            ExcelPartRef part = DrawingsPartRef.GetPartById(relationshipId);
            return part.Kind == ExcelPartKind.Image ? part : default;
        }

        private void Save() {
            DrawingsPartRef.SaveRoot();
            _document.MarkPackageDirty();
        }

        private static int GetDrawingOrder(OfficeXmlNode anchor, ExcelPartRef drawingsPart) {
            OfficeXmlNode worksheetDrawing = drawingsPart.Root;
            if (worksheetDrawing.IsNone) {
                return 0;
            }

            int index = 0;
            foreach (OfficeXmlNode child in worksheetDrawing.Children()) {
                if (child == anchor) {
                    return index;
                }
                index++;
            }

            return 0;
        }

        private static int ParseMarkerIndex(string? text) {
            return int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out int value) && value >= 0 ? value : 0;
        }

        private static long ParseMarkerOffset(string? text) {
            return long.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out long value) ? value : 0L;
        }

        private static double GetCropRatio(int? percentage) {
            if (!percentage.HasValue || percentage.Value <= 0) {
                return 0D;
            }

            return Math.Min(0.999D, percentage.Value / 100000D);
        }

        private static void ValidateCropRatio(double ratio, string parameterName) {
            if (double.IsNaN(ratio) || double.IsInfinity(ratio) || ratio < 0D || ratio >= 1D) {
                throw new ArgumentOutOfRangeException(parameterName, "Crop ratios must be finite values greater than or equal to 0 and less than 1.");
            }
        }

        private static int ToCropPercentage(double ratio) => (int)Math.Round(ratio * 100000D);

        private static double GetRotationDegrees(int? angle) {
            if (!angle.HasValue || angle.Value == 0) {
                return 0D;
            }

            double degrees = angle.Value / 60000D;
            degrees %= 360D;
            if (degrees < 0D) {
                degrees += 360D;
            }

            return degrees;
        }

        private static (int Index, long OffsetEmu) ResolveEndMarker(
            int startIndex, long startOffsetEmu, long sizeEmu, int limit, Func<int, long> segmentSizeEmu) {
            int index = Math.Max(0, startIndex);
            decimal remaining = (decimal)sizeEmu + startOffsetEmu;
            int maximumIndex = Math.Min(limit - 1, index + ExcelImageExportLimits.MaximumAnchorSpanCells);
            while (index < maximumIndex) {
                long segment = segmentSizeEmu(index);
                if (remaining < segment) break;
                remaining -= segment;
                index++;
            }
            return (index, checked((long)remaining));
        }

        private int GetMarkerRow() {
            string? row = AnchorNode.Child(Xdr, "from").ChildText(Xdr, "row");
            return int.TryParse(row, out int value) && value >= 0 ? value : 0;
        }

        private int GetMarkerColumn() {
            string? column = AnchorNode.Child(Xdr, "from").ChildText(Xdr, "col");
            return int.TryParse(column, out int value) && value >= 0 ? value : 0;
        }

        private long GetMarkerColumnOffset() {
            string? offset = AnchorNode.Child(Xdr, "from").ChildText(Xdr, "colOff");
            return long.TryParse(offset, out long value) ? value : 0L;
        }

        private long GetMarkerRowOffset() {
            string? offset = AnchorNode.Child(Xdr, "from").ChildText(Xdr, "rowOff");
            return long.TryParse(offset, out long value) ? value : 0L;
        }

        private bool TryGetToMarkerRow(out int row) {
            string? value = AnchorNode.Child(Xdr, "to").ChildText(Xdr, "row");
            return int.TryParse(value, out row) && row >= 0;
        }

        private bool TryGetToMarkerColumn(out int column) {
            string? value = AnchorNode.Child(Xdr, "to").ChildText(Xdr, "col");
            return int.TryParse(value, out column) && column >= 0;
        }

        private long GetToMarkerColumnOffset() {
            string? offset = AnchorNode.Child(Xdr, "to").ChildText(Xdr, "colOff");
            return long.TryParse(offset, out long value) ? value : 0L;
        }

        private long GetToMarkerRowOffset() {
            string? offset = AnchorNode.Child(Xdr, "to").ChildText(Xdr, "rowOff");
            return long.TryParse(offset, out long value) ? value : 0L;
        }

        private long GetExtentCx() {
            if (TryGetTwoCellAnchorSizeEmu(horizontal: true, out long twoCellEmu)) {
                return twoCellEmu;
            }

            long? anchorExtent = AnchorNode.Child(Xdr, "ext").Int64Attribute("cx");
            if (anchorExtent.HasValue && anchorExtent.Value > 0) {
                return anchorExtent.Value;
            }

            long? shapeExtent = Transform.Child(A, "ext").Int64Attribute("cx");
            if (shapeExtent.HasValue && shapeExtent.Value > 0) {
                return shapeExtent.Value;
            }

            return 0L;
        }

        private long GetExtentCy() {
            if (TryGetTwoCellAnchorSizeEmu(horizontal: false, out long twoCellEmu)) {
                return twoCellEmu;
            }

            long? anchorExtent = AnchorNode.Child(Xdr, "ext").Int64Attribute("cy");
            if (anchorExtent.HasValue && anchorExtent.Value > 0) {
                return anchorExtent.Value;
            }

            long? shapeExtent = Transform.Child(A, "ext").Int64Attribute("cy");
            if (shapeExtent.HasValue && shapeExtent.Value > 0) {
                return shapeExtent.Value;
            }

            return 0L;
        }

        private bool TryGetTwoCellAnchorSizeEmu(bool horizontal, out long emu) {
            emu = 0;
            OfficeXmlNode twoCellAnchor = AnchorNode;
            if (!twoCellAnchor.Is(Xdr, "twoCellAnchor")) {
                return false;
            }
            OfficeXmlNode fromMarker = twoCellAnchor.Child(Xdr, "from");
            OfficeXmlNode toMarker = twoCellAnchor.Child(Xdr, "to");
            if (fromMarker.IsNone || toMarker.IsNone) {
                return false;
            }

            string? editAs = twoCellAnchor.Attribute("editAs");
            if (editAs != null && (editAs == "oneCell" || editAs == "absolute")) {
                return false;
            }

            int from = horizontal
                ? ParseMarkerIndex(fromMarker.ChildText(Xdr, "col"))
                : ParseMarkerIndex(fromMarker.ChildText(Xdr, "row"));
            int to = horizontal
                ? ParseMarkerIndex(toMarker.ChildText(Xdr, "col"))
                : ParseMarkerIndex(toMarker.ChildText(Xdr, "row"));
            int maximumMarker = horizontal ? 16384 : 1048576;
            if (!ExcelImageExportLimits.IsValidMarkerSpan(from, to, maximumMarker)) {
                return false;
            }

            long fromOffset = ParseMarkerOffset(horizontal
                ? fromMarker.ChildText(Xdr, "colOff")
                : fromMarker.ChildText(Xdr, "rowOff"));
            long toOffset = ParseMarkerOffset(horizontal
                ? toMarker.ChildText(Xdr, "colOff")
                : toMarker.ChildText(Xdr, "rowOff"));

            if (to < from) {
                return false;
            }

            ExcelWorksheetGeometryIndex geometry = DrawingsPartRef.ParentWorksheet.CreateWorksheetGeometry();
            decimal sizeEmu = 0M;
            for (int index = from; index < to; index++) {
                sizeEmu += horizontal
                    ? geometry.GetColumnLength(index + 1).ToEmus()
                    : geometry.GetRowLength(index + 1).ToEmus();
            }
            sizeEmu += (decimal)toOffset - fromOffset;
            emu = checked((long)Math.Max(1M, sizeEmu));
            return true;
        }
    }
}
