namespace OfficeIMO.Word {
    /// <summary>
    /// Determines how an image fills its bounding box.
    /// </summary>
    public enum ImageFillMode {
        /// <summary>Stretch the image to fill the area.</summary>
        Stretch,
        /// <summary>Tile the image to fill the area.</summary>
        Tile,
        /// <summary>Scale the image uniformly to fit within the area.</summary>
        Fit,
        /// <summary>Place the image at the center without scaling.</summary>
        Center
    }
}
