namespace OfficeIMO.Visio {
    /// <summary>
    /// Represents a connection point on a Visio shape.
    /// </summary>
    public class VisioConnectionPoint {
        public VisioConnectionPoint(double x, double y, double dirX, double dirY) {
            X = x;
            Y = y;
            DirX = dirX;
            DirY = dirY;
        }

        /// <summary>
        /// X coordinate of the connection point relative to the shape.
        /// </summary>
        public double X { get; set; }

        /// <summary>
        /// Y coordinate of the connection point relative to the shape.
        /// </summary>
        public double Y { get; set; }

        /// <summary>
        /// Directional X component of the connection point.
        /// </summary>
        public double DirX { get; set; }

        /// <summary>
        /// Directional Y component of the connection point.
        /// </summary>
        public double DirY { get; set; }
    }
}
