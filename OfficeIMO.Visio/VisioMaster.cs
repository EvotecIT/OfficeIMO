namespace OfficeIMO.Visio {
    /// <summary>
    /// Represents a Visio master.
    /// </summary>
    public class VisioMaster {
        /// <summary>
        /// Initializes a new instance of the <see cref="VisioMaster"/> class.
        /// </summary>
        /// <param name="id">Master identifier.</param>
        /// <param name="nameU">Universal master name.</param>
        /// <param name="shape">Shape defining the master.</param>
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
        /// Gets the universal master name.
        /// </summary>
        public string NameU { get; }

        /// <summary>
        /// Gets the shape associated with the master.
        /// </summary>
        public VisioShape Shape { get; }
    }
}