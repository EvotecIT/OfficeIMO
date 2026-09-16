namespace OfficeIMO.Visio {
    /// <summary>
    /// Specifies the type of connector used to link shapes.
    /// </summary>
    public enum ConnectorKind {
        /// <summary>Straight line connector.</summary>
        Straight,
        /// <summary>Right-angle connector.</summary>
        RightAngle,
        /// <summary>Curved connector.</summary>
        Curved,
        /// <summary>Connector whose type is determined dynamically.</summary>
        Dynamic
    }
}
