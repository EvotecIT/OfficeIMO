namespace OfficeIMO.Visio {
    /// <summary>
    /// Specifies when Visio may reroute a dynamic connector.
    /// </summary>
    public enum VisioConnectorRerouteBehavior {
        /// <summary>Reroute freely.</summary>
        Freely = 0,

        /// <summary>Reroute as needed when manually requested.</summary>
        AsNeeded = 1,

        /// <summary>Never reroute.</summary>
        Never = 2,

        /// <summary>Reroute when the connector crosses another connector.</summary>
        OnCrossover = 3
    }
}
