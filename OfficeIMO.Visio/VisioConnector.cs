using System;
using System.Globalization;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Connects two shapes together.
    /// </summary>
    public class VisioConnector {
        private static int _idCounter;

        public VisioConnector(VisioShape from, VisioShape to) : this(GetNextId(from, to), from, to) {
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

        public VisioConnectionPoint? FromConnectionPoint { get; set; }

        public VisioConnectionPoint? ToConnectionPoint { get; set; }

        private static string GetNextId(VisioShape from, VisioShape to) {
            int fromId = int.TryParse(from.Id, out int fi) ? fi : 0;
            int toId = int.TryParse(to.Id, out int ti) ? ti : 0;
            int newId = Math.Max(Math.Max(fromId, toId) + 1, _idCounter + 1);
            _idCounter = newId;
            return newId.ToString(CultureInfo.InvariantCulture);
        }
    }
}

