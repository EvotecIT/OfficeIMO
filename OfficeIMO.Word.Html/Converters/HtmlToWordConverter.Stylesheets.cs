using AngleSharp;
using AngleSharp.Css;
using AngleSharp.Css.Dom;
using AngleSharp.Css.Parser;
using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using AngleSharp.Io;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Concurrent;
using System.Net;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Word.Html {
    internal partial class HtmlToWordConverter {
        private async Task LoadAndParseCssAsync(IBrowsingContext context, Url url, CancellationToken cancellationToken) {
            var loader = context.GetService<IResourceLoader>();
            if (loader == null) {
                return;
            }
            var request = new ResourceRequest(null!, url);
            var download = loader.FetchAsync(request);
            var response = await download.Task.ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
            if (response.StatusCode == HttpStatusCode.OK) {
                using var reader = new StreamReader(response.Content);
#if NET8_0_OR_GREATER
                var css = await reader.ReadToEndAsync(cancellationToken).ConfigureAwait(false);
#else
                var css = await reader.ReadToEndAsync().ConfigureAwait(false);
                cancellationToken.ThrowIfCancellationRequested();
#endif
                ParseCss(css, url.Href);
            }
        }

        private void TryLoadCssFromFileUrl(Url url) {
            if (!Uri.TryCreate(url.Href, UriKind.Absolute, out var fileUri) || !fileUri.IsFile) {
                return;
            }

            var localPath = fileUri.LocalPath;
            if (string.IsNullOrEmpty(localPath) || !File.Exists(localPath)) {
                return;
            }

            ParseCss(File.ReadAllText(localPath), localPath);
        }
        private void ApplyClassStyle(IElement element, WordParagraph paragraph, HtmlToWordOptions options) {
            string? classAttr = element.GetAttribute("class");
            if (string.IsNullOrWhiteSpace(classAttr)) {
                return;
            }

            foreach (var cls in (classAttr ?? string.Empty).Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)) {
                if (options.ClassStyles.TryGetValue(cls, out var style) || _cssClassStyles.TryGetValue(cls, out style)) {
                    paragraph.Style = style;
                    break;
                }

                var args = WordHtmlConverterExtensions.OnStyleMissing(paragraph, cls);
                if (args.Style.HasValue) {
                    paragraph.Style = args.Style.Value;
                    break;
                }
                if (!string.IsNullOrEmpty(args.StyleId)) {
                    paragraph.SetStyleId(args.StyleId!);
                    break;
                }
            }
        }

        private void ParseCss(string css, string? key = null) {
            if (string.IsNullOrWhiteSpace(css)) {
                return;
            }

            key ??= ComputeHash(css);
            if (!_stylesheetCache.TryGetValue(key, out var rules)) {
                try {
                    var sheet = _cssParser.ParseStyleSheet(css);
                    rules = sheet.Rules.OfType<ICssStyleRule>().ToArray();
                    _stylesheetCache[key] = rules;
                } catch (Exception) {
                    _stylesheetCache[key] = Array.Empty<ICssStyleRule>();
                    return;
                }
            }

            foreach (var rule in rules) {
                _cssRules.Add(rule);

                var mapped = CssStyleMapper.MapParagraphStyle(rule.Style?.CssText);
                if (mapped.HasValue) {
                    var selectorText = rule.SelectorText;
                    if (!string.IsNullOrWhiteSpace(selectorText)) {
                        foreach (var part in selectorText.Split(',')) {
                            foreach (Match match in _classRegex.Matches(part)) {
                                _cssClassStyles[match.Groups[1].Value] = mapped.Value;
                            }
                        }
                    }
                }
            }
        }

        private static string ComputeHash(string content) {
            using var sha = SHA256.Create();
            var bytes = sha.ComputeHash(Encoding.UTF8.GetBytes(content));
            return BitConverter.ToString(bytes).Replace("-", "");
        }

        private void ApplyCssToElement(IElement element) {
            if (_cssRules.Count == 0) {
                return;
            }

            var accumulated = new Dictionary<string, (string Value, Priority Specificity, bool Important, int Order)>(
                StringComparer.OrdinalIgnoreCase);
            for (int ruleIndex = 0; ruleIndex < _cssRules.Count; ruleIndex++) {
                var rule = _cssRules[ruleIndex];
                var selector = rule.Selector;
                if (selector != null && selector.Match(element, null)) {
                    var specificity = selector.Specificity;
                    foreach (var property in rule.Style) {
                        var name = property.Name;
                        var important = property.IsImportant;
                        var candidate = (Value: property.Value, Specificity: specificity, Important: important, Order: ruleIndex);
                        if (!accumulated.TryGetValue(name, out var existing)) {
                            accumulated[name] = candidate;
                        } else if (candidate.Important != existing.Important) {
                            if (candidate.Important) {
                                accumulated[name] = candidate;
                            }
                        } else if (specificity > existing.Specificity || (specificity == existing.Specificity && ruleIndex >= existing.Order)) {
                            accumulated[name] = candidate;
                        }
                    }
                }
            }

            var inline = element.GetAttribute("style");
            if (!string.IsNullOrEmpty(inline)) {
                try {
                    var declaration = _cssParser.ParseDeclaration(inline);
                    foreach (var property in declaration) {
                        var name = property.Name;
                        var important = property.IsImportant;
                        var candidate = (Value: property.Value, Specificity: Priority.Inline, Important: important, Order: int.MaxValue);
                        if (!accumulated.TryGetValue(name, out var existing)) {
                            accumulated[name] = candidate;
                        } else if (candidate.Important != existing.Important) {
                            if (candidate.Important) {
                                accumulated[name] = candidate;
                            }
                        } else if (Priority.Inline > existing.Specificity || (Priority.Inline == existing.Specificity && candidate.Order >= existing.Order)) {
                            accumulated[name] = candidate;
                        }
                    }
                } catch (Exception) {
                    // ignore invalid inline style
                }
            }

            if (accumulated.Count > 0) {
                var sb = new StringBuilder();
                foreach (var kvp in accumulated) {
                    sb.Append(kvp.Key).Append(':').Append(kvp.Value.Value).Append(';');
                }
                element.SetAttribute("style", sb.ToString());
            }
        }
    }
}
