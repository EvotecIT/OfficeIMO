namespace OfficeIMO.PowerPoint.Fluent {
    /// <summary>
    ///     Provides fluent helpers for creating PowerPoint presentations.
    /// </summary>
    public static class PowerPointBuilder {
        /// <summary>
        ///     Creates a new PowerPoint presentation at the given path.
        /// </summary>
        public static PowerPointPresentation Create(string filePath) {
            return PowerPointPresentation.Create(filePath);
        }

        /// <summary>
        ///     Creates a new PowerPoint presentation backed by a stream.
        /// </summary>
        public static PowerPointPresentation Create(System.IO.Stream stream, bool autoSave = true) {
            return PowerPointPresentation.Create(stream, autoSave);
        }
    }
}
