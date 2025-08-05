namespace OfficeIMO.Word.Html {
    public class HtmlOptions {
        /// <summary>
        /// Font family to use for the document. Default is "Calibri".
        /// </summary>
        public string FontFamily { get; set; } = "Calibri";
        
        /// <summary>
        /// Font family to use for code blocks. Default is "Consolas".
        /// </summary>
        public string CodeFontFamily { get; set; } = "Consolas";
        
        /// <summary>
        /// Whether to include default CSS styles in the HTML output. Default is true.
        /// </summary>
        public bool IncludeCss { get; set; } = true;
        
        /// <summary>
        /// Whether to preserve Word styles as inline CSS. Default is true.
        /// </summary>
        public bool PreserveStyles { get; set; } = true;
        
        /// <summary>
        /// Whether to download and embed images from URLs. Default is false.
        /// </summary>
        public bool DownloadImages { get; set; } = false;
        
        /// <summary>
        /// Whether to embed images as base64 data URIs. Default is false.
        /// </summary>
        public bool EmbedImages { get; set; } = false;
        
        /// <summary>
        /// Path to save extracted images when not embedding. 
        /// </summary>
        public string ImageOutputPath { get; set; }
        
        /// <summary>
        /// URL prefix for image src attributes when not embedding.
        /// </summary>
        public string ImageUrlPrefix { get; set; } = "images/";
        
        /// <summary>
        /// Whether to generate a complete HTML document or just the body content. Default is true.
        /// </summary>
        public bool GenerateCompleteDocument { get; set; } = true;
        
        /// <summary>
        /// Document title for the HTML head. 
        /// </summary>
        public string DocumentTitle { get; set; }
        
        /// <summary>
        /// Default font size in points. Default is 11.
        /// </summary>
        public int DefaultFontSize { get; set; } = 11;
    }
}