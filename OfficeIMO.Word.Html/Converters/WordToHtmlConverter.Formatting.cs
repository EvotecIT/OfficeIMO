using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Globalization;

namespace OfficeIMO.Word.Html {
    internal partial class WordToHtmlConverter {

            string MimeFromFileName(string fileName) {
                var ext = Path.GetExtension(fileName)?.ToLowerInvariant();
                return ext switch {
                    ".jpg" => "image/jpeg",
                    ".jpeg" => "image/jpeg",
                    ".png" => "image/png",
                    ".gif" => "image/gif",
                    ".bmp" => "image/bmp",
                    ".tif" => "image/tiff",
                    ".tiff" => "image/tiff",
                    _ => "image/png"
                };
            }

            string FormatNumber(double value) {
                return value.ToString("0.##", CultureInfo.InvariantCulture);
            }

            string FormatTwips(int twips) {
                return FormatNumber(twips / 20.0) + "pt";
            }

            string? GetHighlightKey(HighlightColorValues value) {
                if (value is IEnumValue enumValue && !string.IsNullOrWhiteSpace(enumValue.Value)) {
                    return enumValue.Value;
                }
                return value.ToString();
            }

            string? GetHighlightCss(HighlightColorValues? highlight) {
                if (highlight == null) {
                    return null;
                }
                var key = GetHighlightKey(highlight.Value);
                if (key == null) {
                    return null;
                }
                key = key.Trim();
                if (key.Length == 0) {
                    return null;
                }
                key = key.ToLowerInvariant();
                return key switch {
                    "none" => null,
                    "yellow" => "#ffff00",
                    "green" => "#00ff00",
                    "cyan" => "#00ffff",
                    "magenta" => "#ff00ff",
                    "blue" => "#0000ff",
                    "red" => "#ff0000",
                    "darkblue" => "#00008b",
                    "darkcyan" => "#008b8b",
                    "darkgreen" => "#006400",
                    "darkmagenta" => "#8b008b",
                    "darkred" => "#8b0000",
                    "darkyellow" => "#808000",
                    "darkgray" => "#a9a9a9",
                    "lightgray" => "#d3d3d3",
                    "black" => "#000000",
                    "white" => "#ffffff",
                    _ => null
                };
            }

            bool IsStructuralTag(string tag) {
                switch (tag) {
                    case "section":
                    case "article":
                    case "aside":
                    case "nav":
                    case "header":
                    case "footer":
                    case "main":
                        return true;
                    default:
                        return false;
                }
            }
    }
}
