using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Shape types supported by <see cref="WordDocument.AddShape(ShapeType,double,double,string,string,double,double)"/>.
    /// </summary>
    public enum ShapeType {
        /// <summary>
        /// A rectangular shape.
        /// </summary>
        Rectangle,

        /// <summary>
        /// An elliptical shape.
        /// </summary>
        Ellipse,

        /// <summary>
        /// A straight line shape.
        /// </summary>
        Line,

        /// <summary>
        /// A rectangle with rounded corners.
        /// </summary>
        RoundedRectangle,

        /// <summary>
        /// A triangle.
        /// </summary>
        Triangle,

        /// <summary>
        /// A diamond.
        /// </summary>
        Diamond,

        /// <summary>
        /// A pentagon.
        /// </summary>
        Pentagon,

        /// <summary>
        /// A hexagon.
        /// </summary>
        Hexagon,

        /// <summary>
        /// Arrow pointing right.
        /// </summary>
        RightArrow,

        /// <summary>
        /// Arrow pointing left.
        /// </summary>
        LeftArrow,

        /// <summary>
        /// Arrow pointing up.
        /// </summary>
        UpArrow,

        /// <summary>
        /// Arrow pointing down.
        /// </summary>
        DownArrow,

        /// <summary>
        /// A 5-point star.
        /// </summary>
        Star5,

        /// <summary>
        /// A heart shape.
        /// </summary>
        Heart,

        /// <summary>
        /// A cloud shape.
        /// </summary>
        Cloud,

        /// <summary>
        /// A donut shape.
        /// </summary>
        Donut,

        /// <summary>
        /// A cylindrical can.
        /// </summary>
        Can,

        /// <summary>
        /// A cube shape.
        /// </summary>
        Cube
    }
}
