using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using System.Globalization;
using A = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Represents a worksheet image anchored in the drawing layer.
    /// </summary>
    public sealed class ExcelImage {
        private readonly Xdr.Picture _picture;
        private readonly OpenXmlElement _anchor;
        private readonly DrawingsPart _drawingsPart;
        private readonly ExcelDocument _document;

        internal ExcelImage(Xdr.Picture picture, OpenXmlElement anchor, DrawingsPart drawingsPart, ExcelDocument document) {
            _picture = picture ?? throw new ArgumentNullException(nameof(picture));
            _anchor = anchor ?? throw new ArgumentNullException(nameof(anchor));
            _drawingsPart = drawingsPart ?? throw new ArgumentNullException(nameof(drawingsPart));
            _document = document ?? throw new ArgumentNullException(nameof(document));
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
        /// Gets the image width in pixels from the drawing extent.
        /// </summary>
        public int WidthPixels => EmuToPx(GetExtentCx());

        /// <summary>
        /// Gets the image height in pixels from the drawing extent.
        /// </summary>
        public int HeightPixels => EmuToPx(GetExtentCy());

        /// <summary>
        /// Gets the image horizontal offset from the anchor cell origin in pixels.
        /// </summary>
        public int OffsetXPixels => EmuOffsetToPx(GetMarkerColumnOffset());

        /// <summary>
        /// Gets the image vertical offset from the anchor cell origin in pixels.
        /// </summary>
        public int OffsetYPixels => EmuOffsetToPx(GetMarkerRowOffset());

        /// <summary>
        /// Gets whether the image is positioned with a two-cell drawing anchor.
        /// </summary>
        public bool HasTwoCellAnchor => _anchor is Xdr.TwoCellAnchor;

        /// <summary>
        /// Gets whether the image is positioned with an absolute drawing anchor.
        /// </summary>
        public bool HasAbsoluteAnchor => _anchor is Xdr.AbsoluteAnchor;

        /// <summary>
        /// Gets the 1-based ending row index for two-cell anchored images, when available.
        /// </summary>
        public int? ToRowIndex => TryGetToMarkerRow(out int row) ? row + 1 : null;

        /// <summary>
        /// Gets the 1-based ending column index for two-cell anchored images, when available.
        /// </summary>
        public int? ToColumnIndex => TryGetToMarkerColumn(out int column) ? column + 1 : null;

        /// <summary>
        /// Gets the horizontal offset from the two-cell ending marker column in pixels.
        /// </summary>
        public int ToOffsetXPixels => EmuOffsetToPx(GetToMarkerColumnOffset());

        /// <summary>
        /// Gets the vertical offset from the two-cell ending marker row in pixels.
        /// </summary>
        public int ToOffsetYPixels => EmuOffsetToPx(GetToMarkerRowOffset());

        /// <summary>
        /// Tries to read the image bounds when it is positioned by an absolute worksheet drawing anchor.
        /// </summary>
        public bool TryGetAbsoluteAnchorBounds(out int xPixels, out int yPixels, out int widthPixels, out int heightPixels) {
            xPixels = 0;
            yPixels = 0;
            widthPixels = 0;
            heightPixels = 0;
            if (_anchor is not Xdr.AbsoluteAnchor absoluteAnchor) {
                return false;
            }

            xPixels = ExcelImageExportLimits.ClampAbsoluteOffsetPixels(EmuOffsetToPx(absoluteAnchor.Position?.X?.Value ?? 0L));
            yPixels = ExcelImageExportLimits.ClampAbsoluteOffsetPixels(EmuOffsetToPx(absoluteAnchor.Position?.Y?.Value ?? 0L));
            widthPixels = ExcelImageExportLimits.ClampExtentPixels(EmuToPx(absoluteAnchor.Extent?.Cx?.Value ?? 0L));
            heightPixels = ExcelImageExportLimits.ClampExtentPixels(EmuToPx(absoluteAnchor.Extent?.Cy?.Value ?? 0L));
            if (widthPixels <= 0 || heightPixels <= 0) {
                widthPixels = WidthPixels;
                heightPixels = HeightPixels;
            }

            return widthPixels > 0 && heightPixels > 0;
        }

        /// <summary>
        /// Sets the ending marker for a two-cell anchored image using 1-based worksheet marker coordinates.
        /// </summary>
        /// <param name="toRow">1-based ending row marker.</param>
        /// <param name="toColumn">1-based ending column marker.</param>
        /// <param name="toOffsetXPixels">Horizontal offset from the ending marker column.</param>
        /// <param name="toOffsetYPixels">Vertical offset from the ending marker row.</param>
        public ExcelImage SetTwoCellEndingMarker(int toRow, int toColumn, int toOffsetXPixels = 0, int toOffsetYPixels = 0) {
            if (toRow < 1) throw new ArgumentOutOfRangeException(nameof(toRow));
            if (toColumn < 1) throw new ArgumentOutOfRangeException(nameof(toColumn));
            if (toOffsetXPixels < 0) throw new ArgumentOutOfRangeException(nameof(toOffsetXPixels));
            if (toOffsetYPixels < 0) throw new ArgumentOutOfRangeException(nameof(toOffsetYPixels));
            if (_anchor is not Xdr.TwoCellAnchor anchor || anchor.ToMarker == null) {
                throw new InvalidOperationException("Only two-cell anchored images expose an ending marker.");
            }

            anchor.ToMarker.RowId = new Xdr.RowId((toRow - 1).ToString(CultureInfo.InvariantCulture));
            anchor.ToMarker.ColumnId = new Xdr.ColumnId((toColumn - 1).ToString(CultureInfo.InvariantCulture));
            anchor.ToMarker.RowOffset = new Xdr.RowOffset(PxToEmu(toOffsetYPixels).ToString(CultureInfo.InvariantCulture));
            anchor.ToMarker.ColumnOffset = new Xdr.ColumnOffset(PxToEmu(toOffsetXPixels).ToString(CultureInfo.InvariantCulture));
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
        /// Sets the image size in pixels.
        /// </summary>
        public ExcelImage SetSize(int widthPixels, int heightPixels) {
            if (widthPixels <= 0) throw new ArgumentOutOfRangeException(nameof(widthPixels));
            if (heightPixels <= 0) throw new ArgumentOutOfRangeException(nameof(heightPixels));

            long cx = PxToEmu(widthPixels);
            long cy = PxToEmu(heightPixels);

            var extent = _anchor.GetFirstChild<Xdr.Extent>();
            if (extent != null) {
                extent.Cx = cx;
                extent.Cy = cy;
            } else if (_anchor is Xdr.TwoCellAnchor twoCellAnchor) {
                ResizeTwoCellAnchor(twoCellAnchor, widthPixels, heightPixels);
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

        private void ResizeTwoCellAnchor(Xdr.TwoCellAnchor anchor, int widthPixels, int heightPixels) {
            Xdr.FromMarker? fromMarker = anchor.FromMarker;
            Xdr.ToMarker? toMarker = anchor.ToMarker;
            if (fromMarker == null || toMarker == null) {
                return;
            }

            WorksheetPart? worksheetPart = _drawingsPart.GetParentParts().OfType<WorksheetPart>().FirstOrDefault();
            double maximumDigitWidth = GetDefaultMaximumDigitWidth(worksheetPart);
            int startColumn = ParseMarkerIndex(fromMarker.ColumnId?.Text);
            int startRow = ParseMarkerIndex(fromMarker.RowId?.Text);
            var columnMarker = ResolveEndMarker(
                startColumn,
                ParseMarkerOffset(fromMarker.ColumnOffset?.Text),
                widthPixels,
                A1.MaxColumns,
                index => GetColumnWidthPixels(worksheetPart, index + 1, maximumDigitWidth));
            var rowMarker = ResolveEndMarker(
                startRow,
                ParseMarkerOffset(fromMarker.RowOffset?.Text),
                heightPixels,
                A1.MaxRows,
                index => GetRowHeightPixels(worksheetPart, index + 1));

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

            int width = Math.Max(1, (int)Math.Round(WidthPixels * percent / 100.0));
            int height = Math.Max(1, (int)Math.Round(HeightPixels * percent / 100.0));
            return SetSize(width, height);
        }

        /// <summary>
        /// Moves the image to a one-cell anchor position and optionally applies pixel offsets.
        /// </summary>
        /// <param name="row">1-based target row.</param>
        /// <param name="column">1-based target column.</param>
        /// <param name="offsetXPixels">Horizontal offset from the cell edge, in pixels.</param>
        /// <param name="offsetYPixels">Vertical offset from the cell edge, in pixels.</param>
        public ExcelImage MoveTo(int row, int column, int offsetXPixels = 0, int offsetYPixels = 0) {
            if (row <= 0 || row > A1.MaxRows) throw new ArgumentOutOfRangeException(nameof(row), "Row must be between 1 and the Excel row limit.");
            if (column <= 0 || column > A1.MaxColumns) throw new ArgumentOutOfRangeException(nameof(column), "Column must be between 1 and the Excel column limit.");

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
            long targetColumnOffset = PxToEmu(offsetXPixels);
            long targetRowOffset = PxToEmu(offsetYPixels);

            MoveTwoCellEndMarkerIfNeeded(
                targetColumn - previousColumn,
                targetRow - previousRow,
                targetColumnOffset - previousColumnOffset,
                targetRowOffset - previousRowOffset);

            marker.ColumnId = new Xdr.ColumnId((column - 1).ToString(CultureInfo.InvariantCulture));
            marker.ColumnOffset = new Xdr.ColumnOffset(targetColumnOffset.ToString(CultureInfo.InvariantCulture));
            marker.RowId = new Xdr.RowId((row - 1).ToString(CultureInfo.InvariantCulture));
            marker.RowOffset = new Xdr.RowOffset(targetRowOffset.ToString(CultureInfo.InvariantCulture));
            Save();
            return this;
        }

        private void MoveTwoCellEndMarkerIfNeeded(int columnDelta, int rowDelta, long columnOffsetDelta, long rowOffsetDelta) {
            Xdr.ToMarker? toMarker = _anchor.GetFirstChild<Xdr.ToMarker>();
            if (toMarker == null) {
                return;
            }

            int targetColumn = ParseMarkerIndex(toMarker.ColumnId?.Text) + columnDelta;
            int targetRow = ParseMarkerIndex(toMarker.RowId?.Text) + rowDelta;
            if (targetColumn < 0 || targetColumn >= A1.MaxColumns) {
                throw new ArgumentOutOfRangeException("column", "Moving this two-cell image would place its end marker outside the Excel column limit.");
            }

            if (targetRow < 0 || targetRow >= A1.MaxRows) {
                throw new ArgumentOutOfRangeException("row", "Moving this two-cell image would place its end marker outside the Excel row limit.");
            }

            toMarker.ColumnId = new Xdr.ColumnId(targetColumn.ToString(CultureInfo.InvariantCulture));
            toMarker.ColumnOffset = new Xdr.ColumnOffset((ParseMarkerOffset(toMarker.ColumnOffset?.Text) + columnOffsetDelta).ToString(CultureInfo.InvariantCulture));
            toMarker.RowId = new Xdr.RowId(targetRow.ToString(CultureInfo.InvariantCulture));
            toMarker.RowOffset = new Xdr.RowOffset((ParseMarkerOffset(toMarker.RowOffset?.Text) + rowOffsetDelta).ToString(CultureInfo.InvariantCulture));
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

        private static long PxToEmu(int px) => (long)Math.Round(px * 9525.0);

        private static int ParseMarkerIndex(string? text) {
            return int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out int value) && value >= 0 ? value : 0;
        }

        private static long ParseMarkerOffset(string? text) {
            return long.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out long value) ? value : 0L;
        }

        private static int EmuToPx(long emu) {
            if (emu <= 0) {
                return 0;
            }

            double pixels = Math.Max(1D, Math.Round(emu / 9525.0));
            return ExcelImageExportLimits.ClampExtentPixels(pixels >= int.MaxValue ? int.MaxValue : (int)pixels);
        }

        private static int EmuOffsetToPx(long emu) {
            double pixels = Math.Round(emu / 9525.0);
            if (pixels <= -int.MaxValue) return -int.MaxValue;
            if (pixels >= int.MaxValue) return int.MaxValue;
            return ExcelImageExportLimits.ClampOffsetPixels((int)pixels);
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
            int startIndex,
            long startOffsetEmu,
            int sizePixels,
            int limit,
            Func<int, int> segmentSizePixels) {
            int index = Math.Max(0, startIndex);
            int remainingPixels = ExcelImageExportLimits.SaturatingAddExtent(sizePixels, EmuToPx(startOffsetEmu));
            int maximumIndex = Math.Min(limit - 1, index + ExcelImageExportLimits.MaximumAnchorSpanCells);
            while (index < maximumIndex) {
                int segmentPixels = Math.Max(1, segmentSizePixels(index));
                if (remainingPixels < segmentPixels) {
                    break;
                }

                remainingPixels -= segmentPixels;
                index++;
            }

            return (index, PxToEmu(Math.Max(0, remainingPixels)));
        }

        private static int GetColumnWidthPixels(WorksheetPart? worksheetPart, int columnIndex, double maximumDigitWidth) {
            DocumentFormat.OpenXml.Spreadsheet.Worksheet? worksheet = worksheetPart?.Worksheet;
            DocumentFormat.OpenXml.Spreadsheet.Column? column = worksheet?
                .GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Columns>()?
                .Elements<DocumentFormat.OpenXml.Spreadsheet.Column>()
                .FirstOrDefault(item => item.Min != null && item.Max != null && item.Min.Value <= (uint)columnIndex && item.Max.Value >= (uint)columnIndex);
            if (column?.Hidden?.Value == true) {
                return 0;
            }

            double width = column?.Width?.Value > 0 && column.CustomWidth?.Value == true
                ? column.Width.Value
                : worksheet?.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetFormatProperties>()?.DefaultColumnWidth?.Value > 0
                    ? worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetFormatProperties>()!.DefaultColumnWidth!.Value
                    : 8.43D;
            double pixels = Math.Truncate((256D * width + Math.Truncate(128D / maximumDigitWidth)) / 256D * maximumDigitWidth);
            return Math.Max(1, (int)Math.Round(pixels));
        }

        private static double GetDefaultMaximumDigitWidth(WorksheetPart? worksheetPart) {
            const double fallbackMaximumDigitWidth = 7D;
            try {
                WorkbookPart? workbookPart = worksheetPart?.GetParentParts().OfType<WorkbookPart>().FirstOrDefault();
                DocumentFormat.OpenXml.Spreadsheet.Font? font = workbookPart?.WorkbookStylesPart?.Stylesheet?.Fonts?
                    .Elements<DocumentFormat.OpenXml.Spreadsheet.Font>()
                    .FirstOrDefault();
                string? fontName = font?.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.FontName>()?.Val?.Value;
                if (string.IsNullOrWhiteSpace(fontName)) {
                    return fallbackMaximumDigitWidth;
                }

                double fontSize = font!.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.FontSize>()?.Val?.Value ?? 11D;
                OfficeFontStyle fontStyle = OfficeFontStyle.Regular;
                if (font.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Bold>() != null) {
                    fontStyle |= OfficeFontStyle.Bold;
                }

                if (font.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Italic>() != null) {
                    fontStyle |= OfficeFontStyle.Italic;
                }

                if (font.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Underline>() != null) {
                    fontStyle |= OfficeFontStyle.Underline;
                }

                float maximumDigitWidth = ExcelTextMeasurer.Create(new OfficeFontInfo(fontName, fontSize, fontStyle)).DefaultStyle.MaximumDigitWidth;
                return maximumDigitWidth > 0.0001f ? maximumDigitWidth : fallbackMaximumDigitWidth;
            } catch {
                return fallbackMaximumDigitWidth;
            }
        }

        private static int GetRowHeightPixels(WorksheetPart? worksheetPart, int rowIndex) {
            DocumentFormat.OpenXml.Spreadsheet.Worksheet? worksheet = worksheetPart?.Worksheet;
            DocumentFormat.OpenXml.Spreadsheet.Row? row = worksheet?
                .GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetData>()?
                .Elements<DocumentFormat.OpenXml.Spreadsheet.Row>()
                .FirstOrDefault(item => item.RowIndex != null && item.RowIndex.Value == (uint)rowIndex);
            if (row?.Hidden?.Value == true) {
                return 0;
            }

            double heightPoints = row?.Height?.Value > 0 && row.CustomHeight?.Value == true
                ? row.Height.Value
                : worksheet?.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetFormatProperties>()?.DefaultRowHeight?.Value > 0
                    ? worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetFormatProperties>()!.DefaultRowHeight!.Value
                    : 15D;
            return Math.Max(1, (int)Math.Round(heightPoints * 96D / 72D));
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
            double maximumDigitWidth = GetDefaultMaximumDigitWidth(worksheetPart);
            long basePixels = 0L;
            for (int index = from; index < to; index++) {
                basePixels += horizontal
                    ? GetColumnWidthPixels(worksheetPart, index + 1, maximumDigitWidth)
                    : GetRowHeightPixels(worksheetPart, index + 1);
            }

            int offsetPixels = EmuOffsetToPx(toOffset - fromOffset);
            int pixels = ExcelImageExportLimits.ClampExtentPixels((int)Math.Min(int.MaxValue, Math.Max(1L, basePixels + offsetPixels)));
            emu = PxToEmu(pixels);
            return true;
        }
    }
}
