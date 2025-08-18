namespace OfficeIMO.Visio {
    /// <summary>
    /// Connects two shapes together.
    /// </summary>
    public class VisioConnector {
        public VisioConnector(VisioShape from, VisioShape to) {
            From = from;
            To = to;
        }

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

