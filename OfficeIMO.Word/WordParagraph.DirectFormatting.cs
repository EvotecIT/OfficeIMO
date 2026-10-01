using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

public partial class WordParagraph {
    /// <summary>Maps authored emphasis overrides from document-format adapters, including explicit off values beneath styles.</summary>
    internal void ApplyDirectEmphasis(bool? bold, bool? italic, UnderlineValues? underline) {
        RunProperties properties = IsHyperLink && _stdRun == null
            ? VerifyRunProperties(Hyperlink!._hyperlink!, Hyperlink._run!, Hyperlink._runProperties)
            : VerifyRunProperties();
        properties.Bold = bold.HasValue ? new Bold { Val = bold.Value } : null;
        properties.BoldComplexScript = bold.HasValue ? new BoldComplexScript { Val = bold.Value } : null;
        properties.Italic = italic.HasValue ? new Italic { Val = italic.Value } : null;
        properties.ItalicComplexScript = italic.HasValue ? new ItalicComplexScript { Val = italic.Value } : null;
        properties.Underline = underline.HasValue ? new Underline { Val = underline.Value } : null;
    }
}
