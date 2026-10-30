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
        public global::OfficeIMO.OfficeLength Width { get; set; } = global::OfficeIMO.OfficeLength.Points(64.8);

        /// <summary>
        /// Gets or sets the height of the embedded object when displayed as an icon.
        /// </summary>
        public global::OfficeIMO.OfficeLength Height { get; set; } = global::OfficeIMO.OfficeLength.Points(64.8);

        /// <summary>
        /// Creates a new instance configured to display the embedded object as an icon.
        /// </summary>
        /// <param name="iconPath">The icon image path.</param>
        /// <param name="width">Optional physical icon width; defaults to 64.8 points.</param>
        /// <param name="height">Optional physical icon height; defaults to 64.8 points.</param>
        /// <returns>A configured <see cref="WordEmbeddedObjectOptions"/> instance.</returns>
        public static WordEmbeddedObjectOptions Icon(string? iconPath = null, global::OfficeIMO.OfficeLength? width = null, global::OfficeIMO.OfficeLength? height = null) {
            return new WordEmbeddedObjectOptions {
                DisplayAsIcon = true,
                IconPath = iconPath,
                Width = width ?? global::OfficeIMO.OfficeLength.Points(64.8),
                Height = height ?? global::OfficeIMO.OfficeLength.Points(64.8)
            };
        }
    }
}
