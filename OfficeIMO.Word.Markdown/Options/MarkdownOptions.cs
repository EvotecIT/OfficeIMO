namespace OfficeIMO.Word.Markdown {
    public class MarkdownOptions {
        /// <summary>
        /// Font family to use for the document. Default is "Calibri".
        /// </summary>
        public string FontFamily { get; set; } = "Calibri";
        
        /// <summary>
        /// Font family to use for code blocks. Default is "Consolas".
        /// </summary>
        public string CodeFontFamily { get; set; } = "Consolas";
        
        /// <summary>
        /// Whether to use GitHub Flavored Markdown extensions. Default is true.
        /// </summary>
        public bool UseGitHubFlavored { get; set; } = true;
        
        /// <summary>
        /// Whether to preserve empty lines between paragraphs. Default is true.
        /// </summary>
        public bool PreserveEmptyLines { get; set; } = true;
        
        /// <summary>
        /// Whether to download and embed images from URLs. Default is false.
        /// </summary>
        public bool DownloadImages { get; set; } = false;
        
        /// <summary>
        /// Base path for resolving relative image URLs.
        /// </summary>
        public string ImageBasePath { get; set; }
        
        /// <summary>
        /// Whether to include table of contents if headings are present. Default is false.
        /// </summary>
        public bool GenerateTableOfContents { get; set; } = false;
        
        /// <summary>
        /// Default font size in points. Default is 11.
        /// </summary>
        public int DefaultFontSize { get; set; } = 11;
    }
}