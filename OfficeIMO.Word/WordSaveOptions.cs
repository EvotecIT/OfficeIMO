namespace OfficeIMO.Word {
    /// <summary>
    /// Selects the physical document format for stream saves, where no file extension is available.
    /// </summary>
    public enum WordStreamSaveFormat {
        /// <summary>
        /// Save streams as the standard Office Open XML package format.
        /// </summary>
        OpenXml = 0,

        /// <summary>
        /// Save streams as a native Word 97-2003 legacy .doc compound file.
        /// </summary>
        LegacyDoc = 1
    }

    /// <summary>
    /// Optional behaviors applied during Word document saves.
    /// </summary>
    public sealed class WordSaveOptions {
        /// <summary>
        /// Selects the physical document format for <see cref="WordDocument.Save(System.IO.Stream, WordSaveOptions?)"/>.
        /// File-path saves continue to use the destination extension.
        /// </summary>
        public WordStreamSaveFormat StreamFormat { get; set; }

        /// <summary>
        /// Returns an options instance with all features disabled.
        /// </summary>
        public static WordSaveOptions None => new WordSaveOptions();
    }
}
