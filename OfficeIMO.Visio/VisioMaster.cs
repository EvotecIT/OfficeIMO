namespace OfficeIMO.Visio {
    /// <summary>
    /// Represents a Visio master.
    /// </summary>
    public class VisioMaster {
        public VisioMaster(string id, string nameU, VisioShape shape) {
            Id = id;
            NameU = nameU;
            Shape = shape;
        }

        public string Id { get; }

        public string NameU { get; }

        public VisioShape Shape { get; }
    }
}