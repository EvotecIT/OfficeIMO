namespace OfficeIMO.Visio {
    /// <summary>
    /// Connects two shapes together.
    /// </summary>
    public class VisioConnector {
        private static int _idCounter = 1;

        public VisioConnector(VisioShape from, VisioShape to) : this($"C{_idCounter++}", from, to) {
        }

        public VisioConnector(string id, VisioShape from, VisioShape to) {
            Id = id;
            From = from;
            To = to;
        }

        /// <summary>
        /// Identifier of the connector.
        /// </summary>
        public string Id { get; }

        /// <summary>
        /// Shape from which the connector starts.
        /// </summary>
        public VisioShape From { get; }

        /// <summary>
        /// Shape at which the connector ends.
        /// </summary>
        public VisioShape To { get; }
    }
}

