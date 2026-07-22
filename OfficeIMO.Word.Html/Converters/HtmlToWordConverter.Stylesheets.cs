using AngleSharp;
using AngleSharp.Css;
using AngleSharp.Css.Dom;
using AngleSharp.Css.Parser;
using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using AngleSharp.Io;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Html;
using System.Collections.Concurrent;
using System.Net;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Word.Html {
    internal partial class HtmlToWordConverter {
        private async Task LoadAndParseCssAsync(Url url, CancellationToken cancellationToken) {
            if (!Uri.TryCreate(url.Href, UriKind.Absolute, out var uri)) {
                return;
            }

            if (!TryApplyStylesheetUriPolicy(uri, url.Href)) {
                return;
            }

            try {
                var css = await FetchCssStringAsync(uri, url.Href, cancellationToken).ConfigureAwait(false);
                if (css != null) {
                    ParseCss(css, url.Href);
                }
            } catch (OperationCanceledException ex) when (!cancellationToken.IsCancellationRequested) {
                AddDiagnostic(_options, "StylesheetLoadTimedOut", "Stylesheet resource timed out or was canceled by the resource pipeline and was skipped.", url.Href, ex);
            } catch (OperationCanceledException) {
                throw;
            } catch (HttpRequestException ex) {
                AddDiagnostic(_options, "StylesheetTransportFailed", "Stylesheet resource could not be fetched and was skipped.", url.Href, ex);
            } catch (HtmlConversionLimitException) {
                throw;
            } catch (Exception ex) {
                AddDiagnostic(_options, "StylesheetLoadFailed", "Stylesheet resource could not be loaded and was skipped.", url.Href, ex);
            }
        }

        private void TryLoadCssFromFileUrl(Url url) {
            if (!Uri.TryCreate(url.Href, UriKind.Absolute, out var fileUri) || !fileUri.IsFile) {
                return;
            }

            if (!TryApplyStylesheetUriPolicy(fileUri, url.Href)) {
                return;
            }

            var localPath = fileUri.LocalPath;
            if (string.IsNullOrEmpty(localPath) || !File.Exists(localPath)) {
                return;
            }

            ParseCss(ReadCssFileWithLimit(localPath), localPath);
        }

        private async Task<string?> FetchCssStringAsync(Uri uri, string source, CancellationToken cancellationToken) {
            using var cts = _resourceTimeout.HasValue
                ? CancellationTokenSource.CreateLinkedTokenSource(cancellationToken)
                : null;
            var token = cts?.Token ?? cancellationToken;
            if (cts != null && _resourceTimeout.HasValue) {
                cts.CancelAfter(_resourceTimeout.Value);
            }

            using var request = new HttpRequestMessage(System.Net.Http.HttpMethod.Get, uri);
            using var response = await _httpClient.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, token).ConfigureAwait(false);
            if (!response.IsSuccessStatusCode) {
                AddDiagnostic(_options, "StylesheetHttpStatusRejected", "Stylesheet resource returned a non-success status and was skipped.", source, new HttpRequestException($"{(int)response.StatusCode} {response.StatusCode}"));
                return null;
            }

            var contentType = response.Content.Headers.ContentType?.MediaType;
            if (!IsStylesheetContentTypeAllowed(contentType)) {
                AddDiagnostic(_options, "StylesheetContentTypeRejected", "Stylesheet resource content type is not allowed and was skipped.", source, new HtmlResourceContentTypeException($"Stylesheet content type '{contentType}' is not allowed."));
                return null;
            }

            return await ReadCssContentWithLimitAsync(response.Content, source, token).ConfigureAwait(false);
        }

        private async Task<string> ReadCssContentWithLimitAsync(HttpContent content, string source, CancellationToken cancellationToken) {
            var readLimit = GetCssReadLimit();
            if (readLimit.Limit.HasValue && content.Headers.ContentLength.HasValue && content.Headers.ContentLength.Value > readLimit.Limit.Value) {
                if (readLimit.LimitedByTotalBudget) {
                    ThrowLimitExceeded(_options, "CssTotalSizeLimitExceeded", "Total CSS size exceeded the configured conversion limit.", source, _cssBytesUsed + content.Headers.ContentLength.Value, _options.MaxTotalCssBytes!.Value);
                }

                ThrowLimitExceeded(_options, "CssSizeLimitExceeded", "CSS size exceeded the configured conversion limit.", source, content.Headers.ContentLength.Value, readLimit.Limit.Value);
            }

            using var stream = await content.ReadAsStreamAsync().ConfigureAwait(false);
            using var memory = new MemoryStream();
            var buffer = new byte[81920];
            long total = 0;
            while (true) {
                cancellationToken.ThrowIfCancellationRequested();
#if NET8_0_OR_GREATER
                var read = await stream.ReadAsync(buffer.AsMemory(0, buffer.Length), cancellationToken).ConfigureAwait(false);
#else
                var read = await stream.ReadAsync(buffer, 0, buffer.Length).ConfigureAwait(false);
                cancellationToken.ThrowIfCancellationRequested();
#endif
                if (read == 0) {
                    break;
                }

                total += read;
                if (readLimit.Limit.HasValue && total > readLimit.Limit.Value) {
                    if (readLimit.LimitedByTotalBudget) {
                        ThrowLimitExceeded(_options, "CssTotalSizeLimitExceeded", "Total CSS size exceeded the configured conversion limit.", source, _cssBytesUsed + total, _options.MaxTotalCssBytes!.Value);
                    }

                    ThrowLimitExceeded(_options, "CssSizeLimitExceeded", "CSS size exceeded the configured conversion limit.", source, total, readLimit.Limit.Value);
                }

                memory.Write(buffer, 0, read);
            }

            return Encoding.UTF8.GetString(memory.ToArray());
        }

        private string ReadCssFileWithLimit(string path) {
            var readLimit = GetCssReadLimit();
            if (readLimit.Limit.HasValue) {
                var length = new FileInfo(path).Length;
                if (length > readLimit.Limit.Value) {
                    if (readLimit.LimitedByTotalBudget) {
                        ThrowLimitExceeded(_options, "CssTotalSizeLimitExceeded", "Total CSS size exceeded the configured conversion limit.", path, _cssBytesUsed + length, _options.MaxTotalCssBytes!.Value);
                    }

                    ThrowLimitExceeded(_options, "CssSizeLimitExceeded", "CSS size exceeded the configured conversion limit.", path, length, readLimit.Limit.Value);
                }
            }

            return File.ReadAllText(path);
        }

        private bool TryApplyStylesheetUriPolicy(Uri uri, string source) {
            if (!IsStylesheetSchemeAllowed(uri.Scheme, out var detail)) {
                AddDiagnostic(_options, "StylesheetResourceRejectedByPolicy", "Stylesheet resource was skipped because its URI is not allowed by the current stylesheet policy.", source, new HtmlResourcePolicyException(detail));
                return false;
            }

            if (!uri.IsFile && _options.AllowedStylesheetHosts.Count > 0 && !_options.AllowedStylesheetHosts.Contains(uri.Host)) {
                detail = $"Stylesheet host '{uri.Host}' is not allowed.";
                AddDiagnostic(_options, "StylesheetResourceRejectedByPolicy", "Stylesheet resource was skipped because its URI is not allowed by the current stylesheet policy.", source, new HtmlResourcePolicyException(detail));
                return false;
            }

            return true;
        }

        private bool TryApplyLocalStylesheetPolicy(string source) {
            if (IsStylesheetSchemeAllowed(Uri.UriSchemeFile, out var detail)) {
                return true;
            }

            AddDiagnostic(_options, "StylesheetResourceRejectedByPolicy", "Stylesheet resource was skipped because its URI is not allowed by the current stylesheet policy.", source, new HtmlResourcePolicyException(detail));
            return false;
        }

        private bool IsStylesheetSchemeAllowed(string scheme, out string detail) {
            if (_options.AllowedStylesheetUriSchemes.Contains(scheme)) {
                detail = string.Empty;
                return true;
            }

            detail = $"Stylesheet URI scheme '{scheme}' is not allowed.";
            return false;
        }

        private bool IsStylesheetContentTypeAllowed(string? contentType) {
            if (!_options.ValidateStylesheetContentTypes || string.IsNullOrWhiteSpace(contentType)) {
                return true;
            }

            return _options.AllowedStylesheetContentTypes.Contains(contentType!.Trim());
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

                var args = WordHtmlConverterExtensions.OnStyleMissing(options, paragraph, cls);
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

            var cacheKey = ComputeHash(css);
            if (!_stylesheetCache.TryGetValue(cacheKey, out var rules)) {
                try {
                    var sheet = _cssParser.ParseStyleSheet(css);
                    rules = sheet.Rules.OfType<ICssStyleRule>().ToArray();
                    _stylesheetCache[cacheKey] = rules;
                } catch (Exception) {
                    _stylesheetCache[cacheKey] = Array.Empty<ICssStyleRule>();
                    return;
                }
            }

            foreach (var rule in rules) {
                RecordCssRule(rule);
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
                if (selector != null) {
                    RecordSelectorEvaluation();
                }
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

        private void RecordCssRule(ICssStyleRule rule) {
            int declarationCount = rule.Style?.Length ?? 0;
            if (declarationCount == 0) return;
            foreach (string selector in HtmlComputedStyleEngine.SplitSelectorList(rule.SelectorText)) {
                if (string.IsNullOrWhiteSpace(selector)) continue;
                try {
                    _cssProcessingBudget.RecordRule(declarationCount);
                } catch (HtmlDomLimitException exception) {
                    ThrowLimitExceeded(_options, exception.Code, exception.Message,
                        exception.LimitSource, exception.Actual, exception.Limit);
                }
            }
        }

        private void RecordSelectorEvaluation() {
            try {
                _cssProcessingBudget.RecordSelectorEvaluation();
            } catch (HtmlDomLimitException exception) {
                ThrowLimitExceeded(_options, exception.Code, exception.Message,
                    exception.LimitSource, exception.Actual, exception.Limit);
            }
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
