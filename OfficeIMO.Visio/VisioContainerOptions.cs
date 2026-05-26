using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Options used when creating an OfficeIMO-authored Visio container shape.
    /// </summary>
    public sealed class VisioContainerOptions {
        /// <summary>
        /// Outer margin around member shapes, in page units.
        /// </summary>
        public double Margin { get; set; } = 0.25D;

        /// <summary>
        /// Additional space reserved for the container heading, in page units.
        /// </summary>
        public double HeadingHeight { get; set; } = 0.35D;

        /// <summary>
        /// Whether Visio may resize the container around members.
        /// </summary>
        public bool AutoResize { get; set; } = true;

        /// <summary>
        /// Whether the container is locked.
        /// </summary>
        public bool Locked { get; set; }

        /// <summary>
        /// Fill color used for the generated container background.
        /// </summary>
        public Color FillColor { get; set; } = Color.FromRgb(218, 242, 252);

        /// <summary>
        /// Border color used for the generated container background.
        /// </summary>
        public Color LineColor { get; set; } = Color.FromRgb(91, 155, 213);

        /// <summary>
        /// Border weight in inches.
        /// </summary>
        public double LineWeight { get; set; } = 0.014D;
    }
}
