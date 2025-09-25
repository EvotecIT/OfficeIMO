namespace OfficeIMO.Word {
    /// <summary>
    /// Options controlling how an embedded object is displayed.
    /// </summary>
    public class WordEmbeddedObjectOptions {
        /// <summary>
        /// Gets or sets a value indicating whether the object should be displayed as an icon.
        /// </summary>
        public bool DisplayAsIcon { get; set; } = true;

        /// <summary>
        /// Gets or sets the path to the icon used when <see cref="DisplayAsIcon"/> is <c>true</c>.
        /// </summary>
        public string? IconPath { get; set; }

        /// <summary>
        /// Gets or sets the width of the embedded object when displayed as an icon.
        /// </summary>
        public double Width { get; set; } = 64.8;

        /// <summary>
        /// Gets or sets the height of the embedded object when displayed as an icon.
        /// </summary>
        public double Height { get; set; } = 64.8;

        /// <summary>
        /// Creates a new instance configured to display the embedded object as an icon.
        /// </summary>
        /// <param name="iconPath">The icon image path.</param>
        /// <param name="width">Optional icon width in points.</param>
        /// <param name="height">Optional icon height in points.</param>
        /// <returns>A configured <see cref="WordEmbeddedObjectOptions"/> instance.</returns>
        public static WordEmbeddedObjectOptions Icon(string? iconPath = null, double? width = null, double? height = null) {
            return new WordEmbeddedObjectOptions {
                DisplayAsIcon = true,
                IconPath = iconPath,
                Width = width ?? 64.8,
                Height = height ?? 64.8
            };
        }
    }
}
