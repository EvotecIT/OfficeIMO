namespace OfficeIMO.Visio {
    /// <summary>
    /// Represents a shape on a Visio page.
    /// </summary>
    public class VisioShape {
        public VisioShape(string id) {
            Id = id;
        }

        /// <summary>
        /// Identifier of the shape.
        /// </summary>
        public string Id { get; }
    }
}

