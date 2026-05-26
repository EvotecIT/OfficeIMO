namespace OfficeIMO.Visio {
    /// <summary>
    /// Describes a visual quality issue found in a Visio page.
    /// </summary>
    public sealed class VisioDiagramQualityIssue {
        /// <summary>
        /// Initializes a new visual quality issue.
        /// </summary>
        public VisioDiagramQualityIssue(
            VisioDiagramQualityIssueSeverity severity,
            string kind,
            string message,
            string? pageName = null,
            string? shapeId = null,
            string? otherShapeId = null,
            string? connectorId = null,
            string? otherConnectorId = null) {
            Severity = severity;
            Kind = kind;
            Message = message;
            PageName = pageName;
            ShapeId = shapeId;
            OtherShapeId = otherShapeId;
            ConnectorId = connectorId;
            OtherConnectorId = otherConnectorId;
        }

        /// <summary>Issue severity.</summary>
        public VisioDiagramQualityIssueSeverity Severity { get; }

        /// <summary>Stable issue kind.</summary>
        public string Kind { get; }

        /// <summary>Human-readable message.</summary>
        public string Message { get; }

        /// <summary>Page name, if available.</summary>
        public string? PageName { get; }

        /// <summary>Primary shape identifier, if relevant.</summary>
        public string? ShapeId { get; }

        /// <summary>Secondary shape identifier, if relevant.</summary>
        public string? OtherShapeId { get; }

        /// <summary>Connector identifier, if relevant.</summary>
        public string? ConnectorId { get; }

        /// <summary>Secondary connector identifier, if relevant.</summary>
        public string? OtherConnectorId { get; }

        /// <inheritdoc />
        public override string ToString() {
            string location = string.IsNullOrWhiteSpace(PageName) ? string.Empty : $" on '{PageName}'";
            return $"{Severity}: {Kind}{location}: {Message}";
        }
    }
}
