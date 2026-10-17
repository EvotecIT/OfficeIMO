using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Html;
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

            private static string? GetHighlightCss(HighlightColorValues? highlight) {
                string? color = WordRunFormatting.GetHighlightColorHex(highlight);
                return color == null ? null : "#" + color;
            }

            private INode ApplyWordTextDecorations(
                IHtmlDocument htmlDocument,
                WordParagraph run,
                INode node,
                WordToHtmlOptions options,
                bool suppressUnderline,
                bool suppressStrikethrough,
                string source) {
                if (!suppressStrikethrough && (run.Strike || run.DoubleStrike)) {
                    if (run.DoubleStrike) {
                        var span = CreateOutputElement(htmlDocument, "span");
                        SetOutputAttribute(htmlDocument, span, "style", "text-decoration-line:line-through;text-decoration-style:double", source + ":double-strike");
                        SetOutputAttribute(htmlDocument, span, "data-officeimo-word-double-strike", "true", source + ":double-strike-metadata");
                        span.AppendChild(node);
                        node = span;
                    } else {
                        var strike = CreateOutputElement(htmlDocument, "s");
                        strike.AppendChild(node);
                        node = strike;
                    }
                }

                WordUnderlineStyle? underline = run.Underline;
                if (!suppressUnderline && underline.HasValue && underline.Value != WordUnderlineStyle.None) {
                    if (underline.Value == WordUnderlineStyle.Single) {
                        var element = CreateOutputElement(htmlDocument, "u");
                        element.AppendChild(node);
                        node = element;
                    } else {
                        string cssStyle = MapWordUnderlineToCssStyle(underline.Value);
                        var span = CreateOutputElement(htmlDocument, "span");
                        SetOutputAttribute(htmlDocument, span, "style", "text-decoration-line:underline;text-decoration-style:" + cssStyle, source + ":underline");
                        SetOutputAttribute(htmlDocument, span, "data-officeimo-word-underline", underline.Value.ToString(), source + ":underline-metadata");
                        span.AppendChild(node);
                        node = span;
                        if (!IsExactCssUnderline(underline.Value)) {
                            AddWordTextStyleApproximation(
                                options,
                                "WordUnderlineStyleApproximated",
                                "Word underline style '" + underline.Value + "' uses the closest CSS " + cssStyle + " pattern; private round-trip metadata retains the exact Word value.",
                                source);
                        }
                    }
                }

                return node;
            }

            private static string MapWordUnderlineToCssStyle(WordUnderlineStyle underline) => underline switch {
                WordUnderlineStyle.Double => "double",
                WordUnderlineStyle.Dotted or WordUnderlineStyle.DottedHeavy => "dotted",
                WordUnderlineStyle.Wave or WordUnderlineStyle.WavyHeavy or WordUnderlineStyle.WavyDouble => "wavy",
                WordUnderlineStyle.Dash or WordUnderlineStyle.DashedHeavy or WordUnderlineStyle.DashLong or WordUnderlineStyle.DashLongHeavy or
                    WordUnderlineStyle.DotDash or WordUnderlineStyle.DashDotHeavy or WordUnderlineStyle.DotDotDash or WordUnderlineStyle.DashDotDotHeavy => "dashed",
                _ => "solid"
            };

            private static bool IsExactCssUnderline(WordUnderlineStyle underline) => underline is
                WordUnderlineStyle.Single or WordUnderlineStyle.Double or WordUnderlineStyle.Dotted or WordUnderlineStyle.Dash or WordUnderlineStyle.Wave;

            private static void AddWordTextStyleApproximation(WordToHtmlOptions options, string code, string message, string source) {
                if (options.ConversionReport.Diagnostics.Any(diagnostic => diagnostic.Code == code && diagnostic.Source == source)) return;
                options.ConversionReport.Add(
                    "OfficeIMO.Word.Html",
                    code,
                    message,
                    HtmlDiagnosticSeverity.Warning,
                    source,
                    null,
                    OfficeConversionLossKind.Approximation);
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
