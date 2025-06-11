using System;

namespace OfficeIMO.Word {
    public class WordEmbeddedObjectOptions {
        public bool DisplayAsIcon { get; set; } = true;
        public string IconPath { get; set; }
        public double Width { get; set; } = 64.8;
        public double Height { get; set; } = 40.8;

        public static WordEmbeddedObjectOptions Icon(string iconPath = null, double? width = null, double? height = null) {
            return new WordEmbeddedObjectOptions {
                DisplayAsIcon = true,
                IconPath = iconPath,
                Width = width ?? 64.8,
                Height = height ?? 40.8
            };
        }
    }
}
