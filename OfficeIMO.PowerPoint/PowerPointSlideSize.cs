using System;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Provides access to presentation slide size in various units.
    /// </summary>
    public sealed class PowerPointSlideSize {
        private readonly PresentationPart _presentationPart;

        internal PowerPointSlideSize(PresentationPart presentationPart) {
            _presentationPart = presentationPart ?? throw new ArgumentNullException(nameof(presentationPart));
        }

        private SlideSize EnsureSlideSize() {
            _presentationPart.Presentation ??= new Presentation();
            SlideSize? size = _presentationPart.Presentation.SlideSize;
            if (size == null) {
                size = new SlideSize { Cx = 12192000, Cy = 6858000, Type = SlideSizeValues.Screen16x9 };
                _presentationPart.Presentation.SlideSize = size;
            }
            return size;
        }

        /// <summary>
        ///     Slide width in EMUs.
        /// </summary>
        public long WidthEmus {
            get => EnsureSlideSize().Cx?.Value ?? 0;
            set => EnsureSlideSize().Cx = checked((int)value);
        }

        /// <summary>
        ///     Slide height in EMUs.
        /// </summary>
        public long HeightEmus {
            get => EnsureSlideSize().Cy?.Value ?? 0;
            set => EnsureSlideSize().Cy = checked((int)value);
        }

        /// <summary>
        ///     Slide width in centimeters.
        /// </summary>
        public double WidthCm {
            get => PowerPointUnits.ToCentimeters(WidthEmus);
            set => WidthEmus = PowerPointUnits.FromCentimeters(value);
        }

        /// <summary>
        ///     Slide height in centimeters.
        /// </summary>
        public double HeightCm {
            get => PowerPointUnits.ToCentimeters(HeightEmus);
            set => HeightEmus = PowerPointUnits.FromCentimeters(value);
        }

        /// <summary>
        ///     Slide width in inches.
        /// </summary>
        public double WidthInches {
            get => PowerPointUnits.ToInches(WidthEmus);
            set => WidthEmus = PowerPointUnits.FromInches(value);
        }

        /// <summary>
        ///     Slide height in inches.
        /// </summary>
        public double HeightInches {
            get => PowerPointUnits.ToInches(HeightEmus);
            set => HeightEmus = PowerPointUnits.FromInches(value);
        }

        /// <summary>
        ///     Slide width in points.
        /// </summary>
        public double WidthPoints {
            get => PowerPointUnits.ToPoints(WidthEmus);
            set => WidthEmus = PowerPointUnits.FromPoints(value);
        }

        /// <summary>
        ///     Slide height in points.
        /// </summary>
        public double HeightPoints {
            get => PowerPointUnits.ToPoints(HeightEmus);
            set => HeightEmus = PowerPointUnits.FromPoints(value);
        }

        /// <summary>
        ///     Slide size preset type.
        /// </summary>
        public SlideSizeValues? Type {
            get => EnsureSlideSize().Type?.Value;
            set => EnsureSlideSize().Type = value;
        }

        /// <summary>
        ///     Gets a content box that respects the specified margin (EMU units).
        /// </summary>
        /// <param name="marginEmus">Margin in EMUs.</param>
        public PowerPointLayoutBox GetContentBox(long marginEmus) {
            if (marginEmus < 0) {
                throw new ArgumentOutOfRangeException(nameof(marginEmus));
            }

            if (marginEmus * 2 > WidthEmus || marginEmus * 2 > HeightEmus) {
                throw new ArgumentOutOfRangeException(nameof(marginEmus), "Margin exceeds slide size.");
            }

            long width = WidthEmus - (2 * marginEmus);
            long height = HeightEmus - (2 * marginEmus);
            return new PowerPointLayoutBox(marginEmus, marginEmus, width, height);
        }

        /// <summary>
        ///     Gets a content box that respects the specified margin in centimeters.
        /// </summary>
        /// <param name="marginCm">Margin in centimeters.</param>
        public PowerPointLayoutBox GetContentBoxCm(double marginCm) {
            return GetContentBox(PowerPointUnits.FromCentimeters(marginCm));
        }

        /// <summary>
        ///     Gets a content box that respects the specified margin in inches.
        /// </summary>
        /// <param name="marginInches">Margin in inches.</param>
        public PowerPointLayoutBox GetContentBoxInches(double marginInches) {
            return GetContentBox(PowerPointUnits.FromInches(marginInches));
        }

        /// <summary>
        ///     Gets a content box that respects the specified margin in points.
        /// </summary>
        /// <param name="marginPoints">Margin in points.</param>
        public PowerPointLayoutBox GetContentBoxPoints(double marginPoints) {
            return GetContentBox(PowerPointUnits.FromPoints(marginPoints));
        }

        /// <summary>
        ///     Creates column layout boxes inside a content area (EMU units).
        /// </summary>
        /// <param name="columnCount">Number of columns to create.</param>
        /// <param name="marginEmus">Margin in EMUs.</param>
        /// <param name="gutterEmus">Gutter between columns in EMUs.</param>
        public PowerPointLayoutBox[] GetColumns(int columnCount, long marginEmus, long gutterEmus) {
            if (columnCount <= 0) {
                throw new ArgumentOutOfRangeException(nameof(columnCount));
            }
            if (gutterEmus < 0) {
                throw new ArgumentOutOfRangeException(nameof(gutterEmus));
            }

            PowerPointLayoutBox content = GetContentBox(marginEmus);
            long totalGutter = gutterEmus * (columnCount - 1);
            if (totalGutter > content.Width) {
                throw new ArgumentOutOfRangeException(nameof(gutterEmus), "Gutter exceeds available width.");
            }

            long columnWidth = (content.Width - totalGutter) / columnCount;
            if (columnWidth <= 0) {
                throw new InvalidOperationException("Column width is not positive.");
            }

            var columns = new PowerPointLayoutBox[columnCount];
            long left = content.Left;
            for (int i = 0; i < columnCount; i++) {
                columns[i] = new PowerPointLayoutBox(left, content.Top, columnWidth, content.Height);
                left += columnWidth + gutterEmus;
            }

            return columns;
        }

        /// <summary>
        ///     Creates column layout boxes inside a content area in centimeters.
        /// </summary>
        /// <param name="columnCount">Number of columns to create.</param>
        /// <param name="marginCm">Margin in centimeters.</param>
        /// <param name="gutterCm">Gutter between columns in centimeters.</param>
        public PowerPointLayoutBox[] GetColumnsCm(int columnCount, double marginCm, double gutterCm) {
            return GetColumns(columnCount,
                PowerPointUnits.FromCentimeters(marginCm),
                PowerPointUnits.FromCentimeters(gutterCm));
        }

        /// <summary>
        ///     Creates column layout boxes inside a content area in inches.
        /// </summary>
        /// <param name="columnCount">Number of columns to create.</param>
        /// <param name="marginInches">Margin in inches.</param>
        /// <param name="gutterInches">Gutter between columns in inches.</param>
        public PowerPointLayoutBox[] GetColumnsInches(int columnCount, double marginInches, double gutterInches) {
            return GetColumns(columnCount,
                PowerPointUnits.FromInches(marginInches),
                PowerPointUnits.FromInches(gutterInches));
        }

        /// <summary>
        ///     Creates column layout boxes inside a content area in points.
        /// </summary>
        /// <param name="columnCount">Number of columns to create.</param>
        /// <param name="marginPoints">Margin in points.</param>
        /// <param name="gutterPoints">Gutter between columns in points.</param>
        public PowerPointLayoutBox[] GetColumnsPoints(int columnCount, double marginPoints, double gutterPoints) {
            return GetColumns(columnCount,
                PowerPointUnits.FromPoints(marginPoints),
                PowerPointUnits.FromPoints(gutterPoints));
        }

        /// <summary>
        ///     Creates row layout boxes inside a content area (EMU units).
        /// </summary>
        /// <param name="rowCount">Number of rows to create.</param>
        /// <param name="marginEmus">Margin in EMUs.</param>
        /// <param name="gutterEmus">Gutter between rows in EMUs.</param>
        public PowerPointLayoutBox[] GetRows(int rowCount, long marginEmus, long gutterEmus) {
            if (rowCount <= 0) {
                throw new ArgumentOutOfRangeException(nameof(rowCount));
            }
            if (gutterEmus < 0) {
                throw new ArgumentOutOfRangeException(nameof(gutterEmus));
            }

            PowerPointLayoutBox content = GetContentBox(marginEmus);
            long totalGutter = gutterEmus * (rowCount - 1);
            if (totalGutter > content.Height) {
                throw new ArgumentOutOfRangeException(nameof(gutterEmus), "Gutter exceeds available height.");
            }

            long rowHeight = (content.Height - totalGutter) / rowCount;
            if (rowHeight <= 0) {
                throw new InvalidOperationException("Row height is not positive.");
            }

            var rows = new PowerPointLayoutBox[rowCount];
            long top = content.Top;
            for (int i = 0; i < rowCount; i++) {
                rows[i] = new PowerPointLayoutBox(content.Left, top, content.Width, rowHeight);
                top += rowHeight + gutterEmus;
            }

            return rows;
        }

        /// <summary>
        ///     Creates row layout boxes inside a content area in centimeters.
        /// </summary>
        /// <param name="rowCount">Number of rows to create.</param>
        /// <param name="marginCm">Margin in centimeters.</param>
        /// <param name="gutterCm">Gutter between rows in centimeters.</param>
        public PowerPointLayoutBox[] GetRowsCm(int rowCount, double marginCm, double gutterCm) {
            return GetRows(rowCount,
                PowerPointUnits.FromCentimeters(marginCm),
                PowerPointUnits.FromCentimeters(gutterCm));
        }

        /// <summary>
        ///     Creates row layout boxes inside a content area in inches.
        /// </summary>
        /// <param name="rowCount">Number of rows to create.</param>
        /// <param name="marginInches">Margin in inches.</param>
        /// <param name="gutterInches">Gutter between rows in inches.</param>
        public PowerPointLayoutBox[] GetRowsInches(int rowCount, double marginInches, double gutterInches) {
            return GetRows(rowCount,
                PowerPointUnits.FromInches(marginInches),
                PowerPointUnits.FromInches(gutterInches));
        }

        /// <summary>
        ///     Creates row layout boxes inside a content area in points.
        /// </summary>
        /// <param name="rowCount">Number of rows to create.</param>
        /// <param name="marginPoints">Margin in points.</param>
        /// <param name="gutterPoints">Gutter between rows in points.</param>
        public PowerPointLayoutBox[] GetRowsPoints(int rowCount, double marginPoints, double gutterPoints) {
            return GetRows(rowCount,
                PowerPointUnits.FromPoints(marginPoints),
                PowerPointUnits.FromPoints(gutterPoints));
        }
    }
}
