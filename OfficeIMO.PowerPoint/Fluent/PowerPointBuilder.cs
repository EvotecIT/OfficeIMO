namespace OfficeIMO.PowerPoint.Fluent {
    /// <summary>
    /// Provides fluent helpers for creating PowerPoint documents.
    /// </summary>
    public static class PowerPointBuilder {
        /// <summary>
        /// Creates a new PowerPoint document.
        /// </summary>
        public static PowerPointDocument Create() {
            return new PowerPointDocument();
        }
    }
}
