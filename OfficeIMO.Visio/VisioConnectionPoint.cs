namespace OfficeIMO.Visio {
    /// <summary>
    /// Represents a connection point on a Visio shape.
    /// </summary>
    public class VisioConnectionPoint {
        /// <summary>
        /// Initializes a new instance of the <see cref="VisioConnectionPoint"/> class.
        /// </summary>
        /// <param name="x">X coordinate relative to the shape.</param>
        /// <param name="y">Y coordinate relative to the shape.</param>
        /// <param name="dirX">Directional X component.</param>
        /// <param name="dirY">Directional Y component.</param>
        public VisioConnectionPoint(double x, double y, double dirX, double dirY) {
            X = x;
            Y = y;
            DirX = dirX;
            DirY = dirY;
        }

        /// <summary>
        /// Gets or sets the X coordinate of the connection point relative to the shape.
        /// </summary>
        public double X { get; set; }

        /// <summary>
        /// Gets or sets the Y coordinate of the connection point relative to the shape.
        /// </summary>
        public double Y { get; set; }

        /// <summary>
        /// Gets or sets the directional X component of the connection point.
        /// </summary>
        public double DirX { get; set; }

        /// <summary>
        /// Gets or sets the directional Y component of the connection point.
        /// </summary>
        public double DirY { get; set; }
    }
}
