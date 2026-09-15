namespace OfficeIMO.Visio {
    /// <summary>
    /// Controls which endpoints of matching connectors should be retargeted.
    /// </summary>
    public enum VisioConnectorEndpointScope {
        /// <summary>Retarget only connector start points.</summary>
        Start = 0,
        /// <summary>Retarget only connector end points.</summary>
        End,
        /// <summary>Retarget both start and end points.</summary>
        Both
    }
}
