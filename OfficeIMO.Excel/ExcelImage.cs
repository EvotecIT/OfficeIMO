using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel.Utilities;
using OfficeIMO.Drawing;
using System.Globalization;
using A = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Represents a worksheet image anchored in the drawing layer.
    /// </summary>
    public sealed partial class ExcelImage {
        private readonly ExcelDrawingHandle _drawingHandle;
        private Xdr.Picture _picture => (Xdr.Picture)_drawingHandle.Element;
        private OpenXmlElement _anchor => _picture.Ancestors().First(element =>
            element is Xdr.OneCellAnchor || element is Xdr.TwoCellAnchor || element is Xdr.AbsoluteAnchor);
        private DrawingsPart _drawingsPart => _drawingHandle.Part;
        private readonly ExcelDocument _document;

        internal ExcelImage(Xdr.Picture picture, OpenXmlElement anchor, DrawingsPart drawingsPart, ExcelSheet sheet) {
            if (anchor == null) throw new ArgumentNullException(nameof(anchor));
            _drawingHandle = new ExcelDrawingHandle(picture, drawingsPart, sheet);
            _document = sheet.Document;
        }

        /// <summary>
        /// Gets or sets the image drawing name.
        /// </summary>
        public string Name {
            get => DrawingProperties?.Name?.Value ?? string.Empty;
            set {
                if (DrawingProperties != null) {
                    DrawingProperties.Name = value ?? string.Empty;
                    Save();
                }
            }
        }

        /// <summary>
        /// Gets or sets the image title metadata.
        /// </summary>
        public string Title {
            get => DrawingProperties?.Title?.Value ?? string.Empty;
            set {
                if (DrawingProperties != null) {
                    DrawingProperties.Title = value ?? string.Empty;
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
                int? amount = _picture.BlipFill?.Blip?
                    .GetFirstChild<A.AlphaModulationFixed>()?.Amount?.Value;
                return amount.HasValue
                    ? (int)Math.Round((100000 - amount.Value) / 1000D)
                    : null;
            }
            set {
                if (value is < 0 or > 100) {
                    throw new ArgumentOutOfRangeException(nameof(value), "Transparency must be between 0 and 100.");
                }
                A.Blip? blip = _picture.BlipFill?.Blip;
                if (blip == null) return;
                A.AlphaModulationFixed? alpha = blip.GetFirstChild<A.AlphaModulationFixed>();
                if (!value.HasValue) {
                    alpha?.Remove();
                    Save();
                    return;
                }
                foreach (OpenXmlElement transform in blip.ChildElements
                             .Where(child => child.LocalName.StartsWith("alpha", StringComparison.Ordinal)).ToArray()) {
                    transform.Remove();
                }
                alpha = new A.AlphaModulationFixed {
                    Amount = checked((int)Math.Round((100D - value.Value) * 1000D,
                        MidpointRounding.AwayFromZero))
                };
                A.BlipExtensionList? extensions = blip.GetFirstChild<A.BlipExtensionList>();
                if (extensions != null) blip.InsertBefore(alpha, extensions);
                else blip.Append(alpha);
                Save();
            }
        }

        /// <summary>
        /// Gets or sets the image alternative text description.
        /// </summary>
        public string Description {
            get => DrawingProperties?.Description?.Value ?? string.Empty;
            set {
                if (DrawingProperties != null) {
                    DrawingProperties.Description = value ?? string.Empty;
                    Save();
                }
            }
        }

        /// <summary>
        /// Gets whether Excel should keep the picture aspect ratio locked.
        /// </summary>
        public bool IsAspectRatioLocked {
            get {
                var locks = _picture.NonVisualPictureProperties?.NonVisualPictureDrawingProperties?.GetFirstChild<A.PictureLocks>();
                return locks?.NoChangeAspect?.Value ?? false;
            }
        }

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
        public bool HasTwoCellAnchor => _anchor is Xdr.TwoCellAnchor;

        /// <summary>
        /// Gets whether the image is positioned with an absolute drawing anchor.
        /// </summary>
        public bool HasAbsoluteAnchor => _anchor is Xdr.AbsoluteAnchor;

        internal ExcelImagePlacement Placement {
            get {
                if (_anchor is not Xdr.TwoCellAnchor twoCell) return ExcelImagePlacement.FreeFloating;
                if (twoCell.EditAs?.Value == Xdr.EditAsValues.OneCell) return ExcelImagePlacement.MoveOnly;
                if (twoCell.EditAs?.Value == Xdr.EditAsValues.Absolute) return ExcelImagePlacement.FreeFloating;
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
            if (_anchor is not Xdr.AbsoluteAnchor absoluteAnchor) return false;
            x = OfficeLength.Emus(absoluteAnchor.Position?.X?.Value ?? 0L);
            y = OfficeLength.Emus(absoluteAnchor.Position?.Y?.Value ?? 0L);
            width = OfficeLength.Emus(absoluteAnchor.Extent?.Cx?.Value ?? 0L);
            height = OfficeLength.Emus(absoluteAnchor.Extent?.Cy?.Value ?? 0L);
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
            if (toRow < 1) throw new ArgumentOutOfRangeException(nameof(toRow));
            if (toColumn < 1) throw new ArgumentOutOfRangeException(nameof(toColumn));
            if (toOffsetX.ToEmus() < 0) throw new ArgumentOutOfRangeException(nameof(toOffsetX));
            if (toOffsetY.ToEmus() < 0) throw new ArgumentOutOfRangeException(nameof(toOffsetY));
            return SetNativeTwoCellEndingMarker(toRow, toColumn, toOffsetX, toOffsetY);
        }

        /// <summary>Restores signed native marker offsets after an importing adapter has bounded its geometry.</summary>
        internal ExcelImage SetNativeTwoCellEndingMarker(int toRow, int toColumn, OfficeLength toOffsetX, OfficeLength toOffsetY) {
            if (toRow < 1) throw new ArgumentOutOfRangeException(nameof(toRow));
            if (toColumn < 1) throw new ArgumentOutOfRangeException(nameof(toColumn));
            if (_anchor is not Xdr.TwoCellAnchor anchor || anchor.ToMarker == null) {
                throw new InvalidOperationException("Only two-cell anchored images expose an ending marker.");
            }

            anchor.ToMarker.RowId = new Xdr.RowId((toRow - 1).ToString(CultureInfo.InvariantCulture));
            anchor.ToMarker.ColumnId = new Xdr.ColumnId((toColumn - 1).ToString(CultureInfo.InvariantCulture));
            anchor.ToMarker.RowOffset = new Xdr.RowOffset(toOffsetY.ToEmus().ToString(CultureInfo.InvariantCulture));
            anchor.ToMarker.ColumnOffset = new Xdr.ColumnOffset(toOffsetX.ToEmus().ToString(CultureInfo.InvariantCulture));
            Save();
            return this;
        }

        /// <summary>
        /// Gets the left crop ratio authored on the image, where 0 means no crop and 1 means the full width.
        /// </summary>
        public double CropLeftRatio => GetCropRatio(CropRectangle?.Left?.Value);

        /// <summary>
        /// Gets the top crop ratio authored on the image, where 0 means no crop and 1 means the full height.
        /// </summary>
        public double CropTopRatio => GetCropRatio(CropRectangle?.Top?.Value);

        /// <summary>
        /// Gets the right crop ratio authored on the image, where 0 means no crop and 1 means the full width.
        /// </summary>
        public double CropRightRatio => GetCropRatio(CropRectangle?.Right?.Value);

        /// <summary>
        /// Gets the bottom crop ratio authored on the image, where 0 means no crop and 1 means the full height.
        /// </summary>
        public double CropBottomRatio => GetCropRatio(CropRectangle?.Bottom?.Value);

        /// <summary>
        /// Gets or sets clockwise image rotation in degrees.
        /// </summary>
        public double RotationDegrees {
            get => GetRotationDegrees(Transform?.Rotation?.Value);
            set => SetRotation(value);
        }

        /// <summary>
        /// Gets whether the image has an authored horizontal flip transform.
        /// </summary>
        public bool FlipHorizontal => Transform?.HorizontalFlip?.Value ?? false;

        /// <summary>
        /// Gets whether the image has an authored vertical flip transform.
        /// </summary>
        public bool FlipVertical => Transform?.VerticalFlip?.Value ?? false;

        /// <summary>
        /// Gets this image anchor's zero-based order in the worksheet drawing layer.
        /// </summary>
        public int DrawingOrder => GetDrawingOrder(_anchor, _drawingsPart);

        /// <summary>
        /// Gets the image content type, such as image/png or image/jpeg.
        /// </summary>
        public string ContentType => ImagePart?.ContentType ?? string.Empty;

        /// <summary>
        /// Returns a copy of the image bytes from the worksheet drawing relationship.
        /// </summary>
        public byte[] ToBytes() {
            ImagePart? imagePart = ImagePart;
            if (imagePart == null) {
                return Array.Empty<byte>();
            }

            using Stream source = imagePart.GetStream();
            using var destination = new MemoryStream();
            source.CopyTo(destination);
            return destination.ToArray();
        }

        internal bool TryReadBytes(out byte[] bytes) =>
            TryReadBytes(ExcelImageExportOptions.DefaultMaximumTotalSourceImageBytes, out bytes);

        internal bool TryReadBytes(long remainingAggregateBytes, out byte[] bytes) {
            ImagePart? imagePart = ImagePart;
            if (imagePart == null) {
                bytes = Array.Empty<byte>();
                return false;
            }

            using Stream source = imagePart.GetStream(FileMode.Open, FileAccess.Read);
            return ExcelImageExportLimits.TryReadSourceImageBytes(source, remainingAggregateBytes, out bytes);
        }

        /// <summary>
        /// Sets image title and description metadata.
        /// </summary>
        public ExcelImage SetAltText(string description, string? title = null) {
            if (DrawingProperties != null) {
                DrawingProperties.Description = description ?? string.Empty;
                if (title != null) {
                    DrawingProperties.Title = title;
                }
                Save();
            }

            return this;
        }

        /// <summary>
        /// Clears image alt text metadata for decorative images.
        /// </summary>
        public ExcelImage Decorative() {
            if (DrawingProperties != null) {
                DrawingProperties.Description = string.Empty;
                DrawingProperties.Title = string.Empty;
                Save();
            }

            return this;
        }

        /// <summary>
        /// Controls whether Excel should keep the picture aspect ratio locked.
        /// </summary>
        public ExcelImage LockAspectRatio(bool locked = true) {
            var drawingProps = _picture.NonVisualPictureProperties?.NonVisualPictureDrawingProperties;
            if (drawingProps == null) {
                return this;
            }

            var locks = drawingProps.GetFirstChild<A.PictureLocks>();
            if (locks == null) {
                locks = new A.PictureLocks();
                drawingProps.Append(locks);
            }

            locks.NoChangeAspect = locked;
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

            A.Transform2D? transform = EnsureTransform();
            if (transform == null) return this;
            transform.Rotation = (int)Math.Round(degrees * 60000.0);
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

            Xdr.BlipFill? blipFill = _picture.BlipFill;
            if (blipFill == null) {
                return this;
            }

            if (left == 0D && top == 0D && right == 0D && bottom == 0D) {
                blipFill.SourceRectangle?.Remove();
                Save();
                return this;
            }

            A.SourceRectangle? rectangle = blipFill.SourceRectangle;
            if (rectangle == null) {
                rectangle = new A.SourceRectangle();
                A.Blip? blip = blipFill.GetFirstChild<A.Blip>();
                if (blip != null) {
                    blipFill.InsertAfter(rectangle, blip);
                } else {
                    blipFill.PrependChild(rectangle);
                }
            }

            rectangle.Left = ToCropPercentage(left);
            rectangle.Top = ToCropPercentage(top);
            rectangle.Right = ToCropPercentage(right);
            rectangle.Bottom = ToCropPercentage(bottom);
            Save();
            return this;
        }

        /// <summary>
        /// Sets image flip transforms.
        /// </summary>
        public ExcelImage SetFlip(bool horizontal, bool vertical) {
            A.Transform2D? transform = EnsureTransform();
            if (transform == null) return this;
            transform.HorizontalFlip = horizontal;
            transform.VerticalFlip = vertical;
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

            var extent = _anchor.GetFirstChild<Xdr.Extent>();
            if (extent != null) {
                extent.Cx = cx;
                extent.Cy = cy;
            } else if (_anchor is Xdr.TwoCellAnchor twoCellAnchor) {
                ResizeTwoCellAnchor(twoCellAnchor, width, height);
            }

            var transform = _picture.ShapeProperties?.GetFirstChild<A.Transform2D>();
            var transformExtents = transform?.GetFirstChild<A.Extents>();
            if (transformExtents != null) {
                transformExtents.Cx = cx;
                transformExtents.Cy = cy;
            }

            Save();
            return this;
        }

        private void ResizeTwoCellAnchor(Xdr.TwoCellAnchor anchor, OfficeLength width, OfficeLength height) {
            Xdr.FromMarker? fromMarker = anchor.FromMarker;
            Xdr.ToMarker? toMarker = anchor.ToMarker;
            if (fromMarker == null || toMarker == null) {
                return;
            }

            WorksheetPart? worksheetPart = _drawingsPart.GetParentParts().OfType<WorksheetPart>().FirstOrDefault();
            ExcelWorksheetGeometryIndex geometry = ExcelWorksheetGeometryIndex.Create(worksheetPart);
            int startColumn = ParseMarkerIndex(fromMarker.ColumnId?.Text);
            int startRow = ParseMarkerIndex(fromMarker.RowId?.Text);
            var columnMarker = ResolveEndMarker(
                startColumn,
                ParseMarkerOffset(fromMarker.ColumnOffset?.Text),
                width.ToEmus(),
                ExcelA1.MaxColumns,
                index => geometry.GetColumnLength(index + 1).ToEmus());
            var rowMarker = ResolveEndMarker(
                startRow,
                ParseMarkerOffset(fromMarker.RowOffset?.Text),
                height.ToEmus(),
                ExcelA1.MaxRows,
                index => geometry.GetRowLength(index + 1).ToEmus());

            toMarker.ColumnId = new Xdr.ColumnId(columnMarker.Index.ToString(CultureInfo.InvariantCulture));
            toMarker.ColumnOffset = new Xdr.ColumnOffset(columnMarker.OffsetEmu.ToString(CultureInfo.InvariantCulture));
            toMarker.RowId = new Xdr.RowId(rowMarker.Index.ToString(CultureInfo.InvariantCulture));
            toMarker.RowOffset = new Xdr.RowOffset(rowMarker.OffsetEmu.ToString(CultureInfo.InvariantCulture));
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

            var marker = _anchor.GetFirstChild<Xdr.FromMarker>();
            if (marker == null) {
                throw new NotSupportedException("Only images anchored to worksheet cells can be moved. Absolute image anchors do not have a cell marker.");
            }

            int previousColumn = ParseMarkerIndex(marker.ColumnId?.Text);
            int previousRow = ParseMarkerIndex(marker.RowId?.Text);
            long previousColumnOffset = ParseMarkerOffset(marker.ColumnOffset?.Text);
            long previousRowOffset = ParseMarkerOffset(marker.RowOffset?.Text);
            int targetColumn = column - 1;
            int targetRow = row - 1;
            long targetColumnOffset = offsetX.ToEmus();
            long targetRowOffset = offsetY.ToEmus();

            MoveTwoCellEndMarkerIfNeeded(
                targetColumn - previousColumn,
                targetRow - previousRow,
                (decimal)targetColumnOffset - previousColumnOffset,
                (decimal)targetRowOffset - previousRowOffset);

            marker.ColumnId = new Xdr.ColumnId((column - 1).ToString(CultureInfo.InvariantCulture));
            marker.ColumnOffset = new Xdr.ColumnOffset(targetColumnOffset.ToString(CultureInfo.InvariantCulture));
            marker.RowId = new Xdr.RowId((row - 1).ToString(CultureInfo.InvariantCulture));
            marker.RowOffset = new Xdr.RowOffset(targetRowOffset.ToString(CultureInfo.InvariantCulture));
            Save();
            return this;
        }

        private void MoveTwoCellEndMarkerIfNeeded(int columnDelta, int rowDelta, decimal columnOffsetDelta, decimal rowOffsetDelta) {
            Xdr.ToMarker? toMarker = _anchor.GetFirstChild<Xdr.ToMarker>();
            if (toMarker == null) {
                return;
            }

            int targetColumn = ParseMarkerIndex(toMarker.ColumnId?.Text) + columnDelta;
            int targetRow = ParseMarkerIndex(toMarker.RowId?.Text) + rowDelta;
            if (targetColumn < 0 || targetColumn >= ExcelA1.MaxColumns) {
                throw new ArgumentOutOfRangeException("column", "Moving this two-cell image would place its end marker outside the Excel column limit.");
            }

            if (targetRow < 0 || targetRow >= ExcelA1.MaxRows) {
                throw new ArgumentOutOfRangeException("row", "Moving this two-cell image would place its end marker outside the Excel row limit.");
            }

            long targetOffsetX = checked((long)(ParseMarkerOffset(toMarker.ColumnOffset?.Text) + columnOffsetDelta));
            long targetOffsetY = checked((long)(ParseMarkerOffset(toMarker.RowOffset?.Text) + rowOffsetDelta));
            toMarker.ColumnId = new Xdr.ColumnId(targetColumn.ToString(CultureInfo.InvariantCulture));
            toMarker.ColumnOffset = new Xdr.ColumnOffset(targetOffsetX.ToString(CultureInfo.InvariantCulture));
            toMarker.RowId = new Xdr.RowId(targetRow.ToString(CultureInfo.InvariantCulture));
            toMarker.RowOffset = new Xdr.RowOffset(targetOffsetY.ToString(CultureInfo.InvariantCulture));
        }

        private Xdr.NonVisualDrawingProperties? DrawingProperties
            => _picture.NonVisualPictureProperties?.NonVisualDrawingProperties;

        private A.SourceRectangle? CropRectangle
            => _picture.BlipFill?.SourceRectangle;

        private A.Transform2D? Transform
            => _picture.ShapeProperties?.GetFirstChild<A.Transform2D>();

        private A.Transform2D? EnsureTransform() {
            var shapeProperties = _picture.ShapeProperties;
            if (shapeProperties == null) {
                return null;
            }

            var transform = shapeProperties.GetFirstChild<A.Transform2D>();
            if (transform == null) {
                transform = new A.Transform2D(
                    new A.Offset { X = 0, Y = 0 },
                    new A.Extents { Cx = GetExtentCx(), Cy = GetExtentCy() });
                shapeProperties.PrependChild(transform);
            }

            return transform;
        }

        private ImagePart? ImagePart {
            get {
                string? relationshipId = _picture.BlipFill?.Blip?.Embed?.Value;
                if (string.IsNullOrWhiteSpace(relationshipId)) {
                    return null;
                }

                try {
                    return _drawingsPart.GetPartById(relationshipId!) as ImagePart;
                } catch (ArgumentOutOfRangeException) {
                    return null;
                }
            }
        }

        private void Save() {
            _drawingsPart.WorksheetDrawing?.Save();
            _document.MarkPackageDirty();
        }

        private static int GetDrawingOrder(OpenXmlElement anchor, DrawingsPart drawingsPart) {
            Xdr.WorksheetDrawing? worksheetDrawing = drawingsPart.WorksheetDrawing;
            if (worksheetDrawing == null) {
                return 0;
            }

            OpenXmlElementList children = worksheetDrawing.ChildElements;
            for (int i = 0; i < children.Count; i++) {
                if (ReferenceEquals(children[i], anchor)) {
                    return i;
                }
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
            string? row = _anchor.GetFirstChild<Xdr.FromMarker>()?.RowId?.Text;
            return int.TryParse(row, out int value) && value >= 0 ? value : 0;
        }

        private int GetMarkerColumn() {
            string? column = _anchor.GetFirstChild<Xdr.FromMarker>()?.ColumnId?.Text;
            return int.TryParse(column, out int value) && value >= 0 ? value : 0;
        }

        private long GetMarkerColumnOffset() {
            string? offset = _anchor.GetFirstChild<Xdr.FromMarker>()?.ColumnOffset?.Text;
            return long.TryParse(offset, out long value) ? value : 0L;
        }

        private long GetMarkerRowOffset() {
            string? offset = _anchor.GetFirstChild<Xdr.FromMarker>()?.RowOffset?.Text;
            return long.TryParse(offset, out long value) ? value : 0L;
        }

        private bool TryGetToMarkerRow(out int row) {
            string? value = _anchor.GetFirstChild<Xdr.ToMarker>()?.RowId?.Text;
            return int.TryParse(value, out row) && row >= 0;
        }

        private bool TryGetToMarkerColumn(out int column) {
            string? value = _anchor.GetFirstChild<Xdr.ToMarker>()?.ColumnId?.Text;
            return int.TryParse(value, out column) && column >= 0;
        }

        private long GetToMarkerColumnOffset() {
            string? offset = _anchor.GetFirstChild<Xdr.ToMarker>()?.ColumnOffset?.Text;
            return long.TryParse(offset, out long value) ? value : 0L;
        }

        private long GetToMarkerRowOffset() {
            string? offset = _anchor.GetFirstChild<Xdr.ToMarker>()?.RowOffset?.Text;
            return long.TryParse(offset, out long value) ? value : 0L;
        }

        private long GetExtentCx() {
            if (TryGetTwoCellAnchorSizeEmu(horizontal: true, out long twoCellEmu)) {
                return twoCellEmu;
            }

            long? anchorExtent = _anchor.GetFirstChild<Xdr.Extent>()?.Cx?.Value;
            if (anchorExtent.HasValue && anchorExtent.Value > 0) {
                return anchorExtent.Value;
            }

            long? shapeExtent = _picture.ShapeProperties?.GetFirstChild<A.Transform2D>()?.GetFirstChild<A.Extents>()?.Cx?.Value;
            if (shapeExtent.HasValue && shapeExtent.Value > 0) {
                return shapeExtent.Value;
            }

            return 0L;
        }

        private long GetExtentCy() {
            if (TryGetTwoCellAnchorSizeEmu(horizontal: false, out long twoCellEmu)) {
                return twoCellEmu;
            }

            long? anchorExtent = _anchor.GetFirstChild<Xdr.Extent>()?.Cy?.Value;
            if (anchorExtent.HasValue && anchorExtent.Value > 0) {
                return anchorExtent.Value;
            }

            long? shapeExtent = _picture.ShapeProperties?.GetFirstChild<A.Transform2D>()?.GetFirstChild<A.Extents>()?.Cy?.Value;
            if (shapeExtent.HasValue && shapeExtent.Value > 0) {
                return shapeExtent.Value;
            }

            return 0L;
        }

        private bool TryGetTwoCellAnchorSizeEmu(bool horizontal, out long emu) {
            emu = 0;
            Xdr.TwoCellAnchor? twoCellAnchor = _anchor as Xdr.TwoCellAnchor;
            if (twoCellAnchor?.FromMarker == null || twoCellAnchor.ToMarker == null) {
                return false;
            }

            if (twoCellAnchor.EditAs != null
                && (twoCellAnchor.EditAs.Value == Xdr.EditAsValues.OneCell
                    || twoCellAnchor.EditAs.Value == Xdr.EditAsValues.Absolute)) {
                return false;
            }

            int from = horizontal
                ? ParseMarkerIndex(twoCellAnchor.FromMarker.ColumnId?.Text)
                : ParseMarkerIndex(twoCellAnchor.FromMarker.RowId?.Text);
            int to = horizontal
                ? ParseMarkerIndex(twoCellAnchor.ToMarker.ColumnId?.Text)
                : ParseMarkerIndex(twoCellAnchor.ToMarker.RowId?.Text);
            int maximumMarker = horizontal ? 16384 : 1048576;
            if (!ExcelImageExportLimits.IsValidMarkerSpan(from, to, maximumMarker)) {
                return false;
            }

            long fromOffset = ParseMarkerOffset(horizontal
                ? twoCellAnchor.FromMarker.ColumnOffset?.Text
                : twoCellAnchor.FromMarker.RowOffset?.Text);
            long toOffset = ParseMarkerOffset(horizontal
                ? twoCellAnchor.ToMarker.ColumnOffset?.Text
                : twoCellAnchor.ToMarker.RowOffset?.Text);

            if (to < from) {
                return false;
            }

            WorksheetPart? worksheetPart = _drawingsPart.GetParentParts().OfType<WorksheetPart>().FirstOrDefault();
            ExcelWorksheetGeometryIndex geometry = ExcelWorksheetGeometryIndex.Create(worksheetPart);
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
