using System;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Represents a rectangular layout region on a slide.
    /// </summary>
    public readonly struct PowerPointLayoutBox {
        /// <summary>
        ///     Creates a new layout box in EMUs.
        /// </summary>
        /// <param name="left">Left position in EMUs.</param>
        /// <param name="top">Top position in EMUs.</param>
        /// <param name="width">Width in EMUs.</param>
        /// <param name="height">Height in EMUs.</param>
        public PowerPointLayoutBox(long left, long top, long width, long height) {
            Left = left;
            Top = top;
            Width = width;
            Height = height;
        }

        /// <summary>
        ///     Left position in EMUs.
        /// </summary>
        public long Left { get; }

        /// <summary>
        ///     Top position in EMUs.
        /// </summary>
        public long Top { get; }

        /// <summary>
        ///     Width in EMUs.
        /// </summary>
        public long Width { get; }

        /// <summary>
        ///     Height in EMUs.
        /// </summary>
        public long Height { get; }

        /// <summary>
        ///     Right position in EMUs.
        /// </summary>
        public long Right => Left + Width;

        /// <summary>
        ///     Bottom position in EMUs.
        /// </summary>
        public long Bottom => Top + Height;

        /// <summary>
        ///     Left position in centimeters.
        /// </summary>
        public double LeftCm => PowerPointUnits.ToCentimeters(Left);

        /// <summary>
        ///     Top position in centimeters.
        /// </summary>
        public double TopCm => PowerPointUnits.ToCentimeters(Top);

        /// <summary>
        ///     Width in centimeters.
        /// </summary>
        public double WidthCm => PowerPointUnits.ToCentimeters(Width);

        /// <summary>
        ///     Height in centimeters.
        /// </summary>
        public double HeightCm => PowerPointUnits.ToCentimeters(Height);

        /// <summary>
        ///     Left position in inches.
        /// </summary>
        public double LeftInches => PowerPointUnits.ToInches(Left);

        /// <summary>
        ///     Top position in inches.
        /// </summary>
        public double TopInches => PowerPointUnits.ToInches(Top);

        /// <summary>
        ///     Width in inches.
        /// </summary>
        public double WidthInches => PowerPointUnits.ToInches(Width);

        /// <summary>
        ///     Height in inches.
        /// </summary>
        public double HeightInches => PowerPointUnits.ToInches(Height);

        /// <summary>
        ///     Left position in points.
        /// </summary>
        public double LeftPoints => PowerPointUnits.ToPoints(Left);

        /// <summary>
        ///     Top position in points.
        /// </summary>
        public double TopPoints => PowerPointUnits.ToPoints(Top);

        /// <summary>
        ///     Width in points.
        /// </summary>
        public double WidthPoints => PowerPointUnits.ToPoints(Width);

        /// <summary>
        ///     Height in points.
        /// </summary>
        public double HeightPoints => PowerPointUnits.ToPoints(Height);

        /// <summary>
        ///     Applies the layout box to the provided shape.
        /// </summary>
        /// <param name="shape">Shape to update.</param>
        public void ApplyTo(PowerPointShape shape) {
            if (shape == null) {
                throw new ArgumentNullException(nameof(shape));
            }

            shape.Left = Left;
            shape.Top = Top;
            shape.Width = Width;
            shape.Height = Height;
        }
    }
}
