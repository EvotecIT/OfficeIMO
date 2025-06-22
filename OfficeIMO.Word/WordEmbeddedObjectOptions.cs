using System;

namespace OfficeIMO.Word {
    /// <summary>
    /// Represents the WordEmbeddedObjectOptions.
    /// </summary>
    public class WordEmbeddedObjectOptions {
        public bool DisplayAsIcon { get; set; } = true;
        public string IconPath { get; set; }
        public double Width { get; set; } = 64.8;
        public double Height { get; set; } = 64.8;

        /// <summary>
        /// Executes the Icon method.
        /// </summary>
        /// <param name="iconPath">iconPath.</param>
        /// <param name="width">width.</param>
        /// <param name="height">height.</param>
        /// <returns>The result.</returns>
        public static WordEmbeddedObjectOptions Icon(string iconPath = null, double? width = null, double? height = null) {
            return new WordEmbeddedObjectOptions {
                DisplayAsIcon = true,
                IconPath = iconPath,
                Width = width ?? 64.8,
                Height = height ?? 64.8
            };
        }
    }
}
