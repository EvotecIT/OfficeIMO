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

        private void ParseLeadingStylesheetChildren(IElement element) {
            foreach (var child in element.ChildNodes) {
                if (child is IText text && string.IsNullOrWhiteSpace(text.Text)) {
                    continue;
                }

                if (child is IHtmlStyleElement styleElement) {
                    ParseCss(styleElement.TextContent);
                    continue;
                }

                if (child is IHtmlLinkElement linkElement) {
                    var rel = linkElement.GetAttribute("rel");
                    if (!string.Equals(rel, "stylesheet", StringComparison.OrdinalIgnoreCase)) {
                        break;
                    }

                    var hrefAttr = linkElement.GetAttribute("href");
                    var href = linkElement.Href ?? hrefAttr;
                    if (string.IsNullOrEmpty(href)) {
                        continue;
                    }

                    if (!string.IsNullOrEmpty(hrefAttr) && File.Exists(hrefAttr)) {
                        ParseCss(File.ReadAllText(hrefAttr), hrefAttr);
                        continue;
                    }

                    var url = new Url(href);
                    if (!url.IsAbsolute && linkElement.BaseUrl != null) {
                        url = new Url(new Url(linkElement.BaseUrl), href);
                    }

                    if (url.Scheme == "file") {
                        TryLoadCssFromFileUrl(url);
                    }
                    continue;
                }

                break;
            }
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

            ValidateCssLimit(css, key);

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

        private static readonly HashSet<string> _inheritedCssProperties = new(StringComparer.OrdinalIgnoreCase) {
            "color",
            "direction",
            "font",
            "font-family",
            "font-size",
            "font-style",
            "font-variant",
            "font-weight",
            "letter-spacing",
            "line-height",
            "text-align",
            "text-decoration",
            "text-transform",
            "white-space",
        };

        private void ApplyCssToElement(IElement element) {
            if (_cssRules.Count == 0) {
                ReportUnsupportedInlineCssDiagnostics(element);
                return;
            }

            var accumulated = new Dictionary<string, (string Value, Priority Specificity, bool Important, int Order)>(
                StringComparer.OrdinalIgnoreCase);

            var ancestors = new List<IElement>();
            var parent = element.ParentElement;
            while (parent != null) {
                ancestors.Add(parent);
                parent = parent.ParentElement;
            }
            ancestors.Reverse();

            foreach (var ancestor in ancestors) {
                var inherited = CollectCssDeclarations(ancestor, inheritedOnly: true);
                foreach (var kvp in inherited) {
                    accumulated[kvp.Key] = kvp.Value;
                }
            }

            var own = CollectCssDeclarations(element, inheritedOnly: false);
            foreach (var kvp in own) {
                accumulated[kvp.Key] = kvp.Value;
            }

            if (accumulated.Count > 0) {
                ReportUnsupportedCssDiagnostics(
                    element,
                    accumulated.ToDictionary(pair => pair.Key, pair => pair.Value.Value, StringComparer.OrdinalIgnoreCase));
                var sb = new StringBuilder();
                foreach (var kvp in accumulated) {
                    sb.Append(kvp.Key).Append(':').Append(kvp.Value.Value).Append(';');
                }
                element.SetAttribute("style", sb.ToString());
            }
        }

        private Dictionary<string, (string Value, Priority Specificity, bool Important, int Order)> CollectCssDeclarations(IElement element, bool inheritedOnly) {
            var accumulated = new Dictionary<string, (string Value, Priority Specificity, bool Important, int Order)>(
                StringComparer.OrdinalIgnoreCase);

            for (int ruleIndex = 0; ruleIndex < _cssRules.Count; ruleIndex++) {
                var rule = _cssRules[ruleIndex];
                var selector = rule.Selector;
                if (selector != null && SelectorMatches(rule, element)) {
                    var specificity = selector.Specificity;
                    foreach (var property in rule.Style) {
                        if (inheritedOnly && !_inheritedCssProperties.Contains(property.Name)) {
                            continue;
                        }
                        ApplyCssCandidate(accumulated, property.Name, property.Value, specificity, property.IsImportant, ruleIndex);
                    }
                }
            }

            var inline = element.GetAttribute("style");
            if (!string.IsNullOrEmpty(inline)) {
                try {
                    var declaration = _cssParser.ParseDeclaration(inline);
                    foreach (var property in declaration) {
                        if (inheritedOnly && !_inheritedCssProperties.Contains(property.Name)) {
                            continue;
                        }
                        ApplyCssCandidate(accumulated, property.Name, property.Value, Priority.Inline, property.IsImportant, int.MaxValue);
                    }
                } catch (Exception) {
                    // ignore invalid inline style
                }
            }

            return accumulated;
        }

        private static bool SelectorMatches(ICssStyleRule rule, IElement element) {
            if (rule.Selector?.Match(element, null) == true) {
                return true;
            }

            var selectorText = rule.SelectorText;
            if (string.IsNullOrWhiteSpace(selectorText)) {
                return false;
            }

            foreach (var selector in selectorText.Split(',')) {
                if (SimpleSelectorMatches(selector.Trim(), element)) {
                    return true;
                }
            }

            return false;
        }

        private static bool SimpleSelectorMatches(string selector, IElement element) {
            if (string.IsNullOrWhiteSpace(selector) ||
                selector.IndexOfAny(new[] { ' ', '>', '+', '~', '[', ']' }) >= 0) {
                return false;
            }

            var pseudoIndex = selector.IndexOf(':');
            if (pseudoIndex >= 0) {
                selector = selector.Substring(0, pseudoIndex);
            }

            if (string.IsNullOrWhiteSpace(selector)) {
                return false;
            }

            string? expectedId = null;
            var hashIndex = selector.IndexOf('#');
            if (hashIndex >= 0) {
                var idEnd = selector.IndexOf('.', hashIndex + 1);
                expectedId = idEnd >= 0 ? selector.Substring(hashIndex + 1, idEnd - hashIndex - 1) : selector.Substring(hashIndex + 1);
                selector = selector.Remove(hashIndex, expectedId.Length + 1);
            }

            if (!string.IsNullOrEmpty(expectedId) && !string.Equals(element.Id, expectedId, StringComparison.Ordinal)) {
                return false;
            }

            var classMatches = _classRegex.Matches(selector);
            var classAttribute = element.GetAttribute("class") ?? string.Empty;
            var classes = new HashSet<string>(
                classAttribute.Split(new[] { ' ', '\t', '\r', '\n', '\f' }, StringSplitOptions.RemoveEmptyEntries),
                StringComparer.Ordinal);
            foreach (Match match in classMatches) {
                if (!classes.Contains(match.Groups[1].Value)) {
                    return false;
                }
            }

            var tagEnd = selector.IndexOf('.');
            var tag = tagEnd >= 0 ? selector.Substring(0, tagEnd) : selector;
            if (!string.IsNullOrWhiteSpace(tag) && !string.Equals(element.TagName, tag, StringComparison.OrdinalIgnoreCase)) {
                return false;
            }

            return !string.IsNullOrWhiteSpace(tag) || classMatches.Count > 0 || !string.IsNullOrEmpty(expectedId);
        }

        private static void ApplyCssCandidate(
            Dictionary<string, (string Value, Priority Specificity, bool Important, int Order)> accumulated,
            string name,
            string value,
            Priority specificity,
            bool important,
            int order) {
            var candidate = (Value: value, Specificity: specificity, Important: important, Order: order);
            if (!accumulated.TryGetValue(name, out var existing)) {
                accumulated[name] = candidate;
            } else if (candidate.Important != existing.Important) {
                if (candidate.Important) {
                    accumulated[name] = candidate;
                }
            } else if (specificity > existing.Specificity || (specificity == existing.Specificity && order >= existing.Order)) {
                accumulated[name] = candidate;
            }
        }
    }
}
