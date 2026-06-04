using AngleSharp.Dom;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text;
using System.Threading;

namespace OfficeIMO.Word.Html {
    internal partial class WordToHtmlConverter {
        private static void AppendStyleDefinitions(
            WordDocument document,
            IDocument htmlDoc,
            IElement head,
            HashSet<string> paragraphStyles,
            HashSet<string> runStyles,
            CancellationToken cancellationToken) {
            if (paragraphStyles.Count == 0 && runStyles.Count == 0) {
                return;
            }

            var stylePart = document._wordprocessingDocument?.MainDocumentPart?.StyleDefinitionsPart;
            var styleMap = (stylePart?.Styles?.OfType<Style>() ?? Enumerable.Empty<Style>())
                .ToDictionary<Style, string, Style>(s => s.StyleId!, s => s, StringComparer.OrdinalIgnoreCase);

            string BuildCss(string styleId) {
                var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                void Merge(string id) {
                    var key = id;
                    if (string.IsNullOrEmpty(key)) {
                        return;
                    }
                    if (!visited.Add(key)) {
                        return;
                    }
                    if (!styleMap.TryGetValue(key, out var def)) {
                        return;
                    }
                    var baseId = def.BasedOn?.Val;
                    if (!string.IsNullOrEmpty(baseId)) {
                        Merge(baseId!);
                    }
                    var pPr = def.StyleParagraphProperties;
                    if (pPr?.Justification?.Val != null) {
                        var justifyVal = pPr.Justification.Val.Value;
                        var alignment = "left";
                        if (justifyVal == JustificationValues.Center) {
                            alignment = "center";
                        } else if (justifyVal == JustificationValues.Right) {
                            alignment = "right";
                        } else if (justifyVal == JustificationValues.Both) {
                            alignment = "justify";
                        }
                        props["text-align"] = alignment;
                    }
                    var rPr = def.StyleRunProperties;
                    if (rPr != null) {
                        if (rPr.Bold != null) {
                            props["font-weight"] = "bold";
                        }
                        if (rPr.Italic != null) {
                            props["font-style"] = "italic";
                        }
                        var underline = rPr.Underline?.Val?.Value;
                        if (underline != null && underline != UnderlineValues.None) {
                            props["text-decoration"] = "underline";
                        }
                        var colorVal = rPr.Color?.Val?.Value;
                        if (!string.IsNullOrEmpty(colorVal)) {
                            var cv = colorVal!;
                            props["color"] = "#" + cv.ToLowerInvariant();
                        }
                        var sizeVal = rPr.FontSize?.Val;
                        if (!string.IsNullOrEmpty(sizeVal) && int.TryParse(sizeVal, out int sz)) {
                            props["font-size"] = (sz / 2.0).ToString("0.##") + "pt";
                        }
                        var font = rPr.RunFonts?.Ascii?.Value;
                        if (!string.IsNullOrEmpty(font)) {
                            var value = font!;
                            props["font-family"] = value.Contains(' ') ? $"\"{value}\"" : value;
                        }
                    }
                }

                Merge(styleId);

                return string.Join(" ", props.Select(kv => kv.Key + ':' + kv.Value + ';'));
            }

            var styleElement = htmlDoc.CreateElement("style");
            var sb = new StringBuilder();

            foreach (var s in paragraphStyles) {
                cancellationToken.ThrowIfCancellationRequested();
                var css = BuildCss(s);
                sb.Append('.').Append(s).Append(" { ").Append(css).Append(" }\n");
            }
            foreach (var s in runStyles) {
                cancellationToken.ThrowIfCancellationRequested();
                var css = BuildCss(s);
                sb.Append('.').Append(s).Append(" { ").Append(css).Append(" }\n");
            }
            styleElement.TextContent = sb.ToString();
            head.AppendChild(styleElement);
        }
    }
}
