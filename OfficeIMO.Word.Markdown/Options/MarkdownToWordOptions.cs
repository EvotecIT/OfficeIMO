using System;
using OfficeIMO.Word;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Markdown {
    /// <summary>
    /// Options controlling Markdown to Word conversion.
    /// </summary>
    public class MarkdownToWordOptions {
        /// <summary>
        /// Optional font family applied to created runs during conversion.
        /// </summary>
        public string FontFamily { get; set; }
        
        /// <summary>
        /// Optional default page size applied when creating new documents.
        /// </summary>
        public WordPageSize? DefaultPageSize { get; set; }
        
        /// <summary>
        /// Optional default page orientation applied when creating new documents.
        /// </summary>
        public PageOrientationValues? DefaultOrientation { get; set; }
        
        /// <summary>
        /// Applies default page settings to the provided document instance.
        /// </summary>
        public void ApplyDefaults(WordDocument document) {
            if (document == null) {
                throw new ArgumentNullException(nameof(document));
            }
            
            if (DefaultPageSize.HasValue) {
                document.PageSettings.PageSize = DefaultPageSize.Value;
            }
            
            if (DefaultOrientation.HasValue) {
                document.PageOrientation = DefaultOrientation.Value;
            }
        }
    }
}
