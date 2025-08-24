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
        /// <param name="shape">The shape associated with the master.</param>
        public VisioMaster(string id, string nameU, VisioShape shape) {
            Id = id;
            NameU = nameU;
            Shape = shape;
        }

        /// <summary>
        /// Identifier of the master.
        /// </summary>
        public string Id { get; }

        /// <summary>
        /// Universal name of the master.
        /// </summary>
        public string NameU { get; }

        /// <summary>
        /// Shape that defines the master.
        /// </summary>
        public VisioShape Shape { get; }
    }
}