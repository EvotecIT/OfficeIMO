namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// Represents a shape on a PowerPoint slide.
    /// </summary>
    public class PowerPointShape {
        public PowerPointShape(string id) {
            Id = id;
        }

        /// <summary>
        /// Identifier of the shape.
        /// </summary>
        public string Id { get; }
    }
}
