namespace OfficeIMO.PowerPoint.Fluent {
    /// <summary>
    /// Provides fluent helpers for creating PowerPoint presentations.
    /// </summary>
    public static class PowerPointBuilder {
        /// <summary>
        /// Creates a new PowerPoint presentation at the given path.
        /// </summary>
        public static PowerPointPresentation Create(string filePath) {
            return PowerPointPresentation.Create(filePath);
        }
    }
}
