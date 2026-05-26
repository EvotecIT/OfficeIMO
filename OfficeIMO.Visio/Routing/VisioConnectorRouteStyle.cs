namespace OfficeIMO.Visio {
    /// <summary>
    /// Determines how an orthogonal connector route is generated.
    /// </summary>
    public enum VisioConnectorRouteStyle {
        /// <summary>Choose the cleaner route based on endpoint distance.</summary>
        Auto,

        /// <summary>Route horizontally first, then vertically, then horizontally to the target.</summary>
        HorizontalThenVertical,

        /// <summary>Route vertically first, then horizontally, then vertically to the target.</summary>
        VerticalThenHorizontal
    }
}
