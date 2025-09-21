namespace OfficeIMO.Visio {
    /// <summary>
    /// Represents a Visio master.
    /// </summary>
    public class VisioMaster {
        /// <summary>
        /// Initializes a new instance of the <see cref="VisioMaster"/> class.
        /// </summary>
        /// <param name="id">Identifier of the master.</param>
        /// <param name="nameU">Universal name of the master.</param>
        /// <param name="shape">Associated master shape.</param>
        public VisioMaster(string id, string nameU, VisioShape shape) {
            Id = id;
            NameU = nameU;
            Shape = shape;
        }

        /// <summary>
        /// Gets the master identifier.
        /// </summary>
        public string Id { get; }

        /// <summary>
        /// Gets the universal name of the master.
        /// </summary>
        public string NameU { get; }

        /// <summary>
        /// Gets the shape that defines the master.
        /// </summary>
        public VisioShape Shape { get; }
    }
}
