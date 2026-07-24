using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Drawing.Internal;

namespace OfficeIMO.Word.Markdown {
    internal partial class WordToMarkdownConverter {
        internal string ConvertParagraph(WordParagraph paragraph, WordToMarkdownOptions options, bool? hasCheckboxOverride = null, bool checkboxCheckedOverride = false) {
            const string codeLangPrefix = "CodeLang_";
            string? styleId = paragraph.StyleId;
            if (styleId is { Length: > 0 } sid && sid.StartsWith(codeLangPrefix, StringComparison.Ordinal)) {
                var runs = paragraph.GetRuns()
                    .Where(r => !string.IsNullOrEmpty(r.Text))
                    .ToList();
                if (runs.Count > 0) {
                    string language = sid.Substring(codeLangPrefix.Length);
                    string code = string.Concat(runs.Select(r => r.Text));
                    return $"```{language}\n{code}\n```";
                }
            }

            var sb = new StringBuilder();

            if (paragraph.IndentationBefore.HasValue && paragraph.IndentationBefore.Value > 0) {
                int depth = GetBoundedBlockquoteDepth(paragraph.IndentationBefore.Value);
                if (depth > 0) {
                    sb.Append(string.Join(" ", Enumerable.Repeat(">", depth))).Append(' ');
                }
            }

            int? headingLevel = paragraph.Style.HasValue
                ? HeadingStyleMapper.GetLevelForHeadingStyle(paragraph.Style.Value)
                : (int?)null;
            if (headingLevel.HasValue && headingLevel.Value > 0) {
                sb.Append(new string('#', headingLevel.Value)).Append(' ');
            }

            var listInfo = DocumentTraversal.GetListInfo(paragraph);
            if (listInfo != null) {
                int level = ValidateListLevel(listInfo.Value.Level, options.MaxListNestingDepth);
                sb.Append(new string(' ', checked(level * 2)));
                sb.Append(listInfo.Value.Ordered ? "1. " : "- ");
                // Task list (checkbox) mapping — look across all runs in the underlying paragraph
                bool hasCheckbox = hasCheckboxOverride ?? paragraph.IsCheckBox;
                bool done = hasCheckboxOverride.HasValue ? checkboxCheckedOverride : (paragraph.CheckBox?.IsChecked == true);
                if (!hasCheckbox && !hasCheckboxOverride.HasValue) {
                    try {
                        foreach (var r in paragraph.GetRuns()) { if (r.IsCheckBox) { hasCheckbox = true; done = r.CheckBox?.IsChecked == true; break; } }
                    } catch { /* best-effort */ }
                }
                if (hasCheckbox) sb.Append(done ? "[x] " : "[ ] ");
            }

            sb.Append(RenderRuns(paragraph, options));

            return sb.ToString();
        }

        private static readonly System.Collections.Generic.HashSet<string> KnownMonospaceFonts = new System.Collections.Generic.HashSet<string>(System.StringComparer.OrdinalIgnoreCase) {
            "Consolas", "Courier", "Courier New", "Lucida Console", "DejaVu Sans Mono",
            "Menlo", "Monaco", "Inconsolata", "Source Code Pro", "Fira Code",
            "Cascadia Mono", "Cascadia Code", "JetBrains Mono"
        };

        private static string? ResolveConfiguredCodeFont(string? configuredFontFamily) {
            if (string.IsNullOrWhiteSpace(configuredFontFamily)) {
                return null;
            }

            return FontResolver.Resolve(configuredFontFamily) ?? configuredFontFamily;
        }

        private static string? ResolveImplicitCodeFont() {
            var font = FontResolver.Resolve("monospace");
            if (string.IsNullOrWhiteSpace(font)) {
                return null;
            }

            var fontValue = font!;

            if (KnownMonospaceFonts.Contains(fontValue) || fontValue.IndexOf("Mono", StringComparison.OrdinalIgnoreCase) >= 0) {
                return fontValue;
            }

            return null;
        }

        internal string RenderRuns(WordParagraph paragraph, WordToMarkdownOptions options) {
            var sb = new StringBuilder();
            // Inline code detection:
            // 1) If caller specifies options.FontFamily, treat runs with that font as code
            // 2) Else, treat runs with the platform monospace (FontResolver.Resolve("monospace")) as code
            // 3) Else, fallback to a conservative known-monospace allowlist or names containing "Mono"
            string? preferredCodeFont = ResolveConfiguredCodeFont(options.FontFamily);
            string? implicitCodeFont = ResolveImplicitCodeFont();
            foreach (var run in paragraph.GetRuns()) {
                // Respect explicit line breaks embedded in runs (non-page breaks)
                if (run.Break != null && run.PageBreak == null) {
                    // Emit as <br/> marker to stay safe inside tables; the Markdown reader will
                    // translate this back into a hard break when converting to Word/HTML.
                    if (sb.Length > 0) sb.Append("<br/>");
                }
                if (run.IsFootNote && run.FootNote != null && run.FootNote.ReferenceId.HasValue) {
                    long id = run.FootNote.ReferenceId.Value;
                    sb.Append($"[^{id}]");
                    continue;
                }

                if (run.IsImage && run.Image != null) {
                    sb.Append(RenderImage(run.Image, options));
                    continue;
                }

                if (run.IsCheckBox) {
                    continue;
                }

                string? text = run.Text;
                if (string.IsNullOrEmpty(text)) {
                    continue;
                }

                if (run.Bold && run.Italic) {
                    text = $"***{text}***";
                } else if (run.Bold) {
                    text = $"**{text}**";
                } else if (run.Italic) {
                    text = $"*{text}*";
                }

                if (options.EnableUnderline && run.Underline.HasValue && run.Underline.Value != UnderlineValues.None) {
                    text = $"<u>{text}</u>";
                }

                if (run.Strike) {
                    text = $"~~{text}~~";
                }

                if (options.EnableHighlight && run.Highlight.HasValue && run.Highlight.Value != HighlightColorValues.None) {
                    text = $"=={text}==";
                }

                bool code = false;
                var runFont = run.FontFamily;
                if (!string.IsNullOrEmpty(runFont)) {
                    if (!string.IsNullOrEmpty(preferredCodeFont)) {
                        code = string.Equals(runFont, preferredCodeFont, StringComparison.OrdinalIgnoreCase);
                    }
                    if (!code && !string.IsNullOrEmpty(implicitCodeFont)) {
                        code = string.Equals(runFont, implicitCodeFont, StringComparison.OrdinalIgnoreCase);
                    }
                    if (!code) {
                        code = KnownMonospaceFonts.Contains(runFont!) || runFont!.IndexOf("Mono", StringComparison.OrdinalIgnoreCase) >= 0;
                    }
                }
                if (code) {
                    // Choose a fence that is one longer than the longest run of backticks in the text
                    int longest = 0; int current = 0;
                    foreach (char ch in text) { if (ch == '`') { current++; longest = current > longest ? current : longest; } else { current = 0; } }
                    int fenceLen = longest + 1; if (fenceLen < 1) fenceLen = 1;
                    string fence = new string('`', fenceLen);
                    text = fence + text + fence;
                }

                if (run.IsHyperLink && run.Hyperlink != null && run.Hyperlink.Uri != null) {
                    var uri = run.Hyperlink.Uri;
                    string url;
                    if (uri.IsAbsoluteUri) {
                        url = uri.GetComponents(UriComponents.AbsoluteUri, UriFormat.UriEscaped);
                        var original = uri.OriginalString;
                        if (!string.IsNullOrEmpty(original) &&
                            !original.EndsWith("/", StringComparison.Ordinal) &&
                            uri.AbsolutePath == "/" &&
                            url.EndsWith("/", StringComparison.Ordinal)) {
                            url = url.TrimEnd('/');
                        }
                    } else {
                        url = uri.ToString();
                    }
                    text = $"[{text}]({url})";
                }

                sb.Append(text);
            }

            return sb.ToString();
        }

        internal string RenderFootnote(WordFootNote footNote, WordToMarkdownOptions options) {
            var paragraphs = footNote.Paragraphs;
            if (paragraphs == null || paragraphs.Count == 0) return string.Empty;
            var sb = new StringBuilder();
            for (int i = 0; i < paragraphs.Count; i++) {
                if (i > 0) sb.Append(' ');
                sb.Append(RenderRuns(paragraphs[i], options));
            }
            return sb.ToString();
        }

        internal string RenderImage(WordImage image, WordToMarkdownOptions options) {
            if (image == null) {
                return string.Empty;
            }

            string alt = image.Description ?? string.Empty;
            if (image.IsExternal && options.FallbackExternalImagesToLinks) {
                string source = image.ExternalUri?.ToString() ?? image.FilePath;
                if (string.IsNullOrWhiteSpace(source)) {
                    source = image.ExternalRelationshipId ?? string.Empty;
                }

                options.OnWarning?.Invoke($"Externally linked image '{source}' was emitted as a Markdown image reference because the binary payload is not stored in the Word package.");
                return $"![{alt}]({source})";
            }

            if (options.ImageExportMode == ImageExportMode.File) {
                string directory = options.ImageDirectory ?? Directory.GetCurrentDirectory();
                Directory.CreateDirectory(directory);
                string extension = Path.GetExtension(image.FilePath);
                if (string.IsNullOrEmpty(extension)) {
                    extension = ".png";
                }
                string fileName = BuildSafeImageFileName(image.FileName, extension);

                string targetPath = Path.Combine(directory, fileName);
                WriteImageFile(image, targetPath);

                return $"![{alt}]({fileName})";
            } else {
                byte[] bytes = ReadEmbeddedImageBytes(image, options);
                string extension = Path.GetExtension(image.FilePath);
                string mime = extension switch {
                    ".jpg" => "image/jpeg",
                    ".jpeg" => "image/jpeg",
                    ".gif" => "image/gif",
                    ".bmp" => "image/bmp",
                    _ => "image/png"
                };
                string base64 = System.Convert.ToBase64String(bytes);
                return $"![{alt}](data:{mime};base64,{base64})";
            }
        }

        private const int MaximumBlockquoteDepth = 64;

        private static int GetBoundedBlockquoteDepth(double indentationBefore) {
            double computed = Math.Round(indentationBefore / 720d);
            if (double.IsNaN(computed) || computed <= 0) {
                return 0;
            }

            return computed >= MaximumBlockquoteDepth
                ? MaximumBlockquoteDepth
                : (int)computed;
        }

        private string BuildSafeImageFileName(string? suppliedName, string extension) {
            string normalized = (suppliedName ?? string.Empty).Replace('\\', '/');
            int separator = normalized.LastIndexOf('/');
            if (separator >= 0) {
                normalized = normalized.Substring(separator + 1);
            }

            normalized = Path.GetFileName(normalized);
            var invalid = Path.GetInvalidFileNameChars();
            var safe = new StringBuilder(normalized.Length);
            foreach (char character in normalized) {
                safe.Append(character < ' ' || invalid.Contains(character) ? '_' : character);
            }

            string fileName = safe.ToString().Trim();
            if (string.IsNullOrEmpty(fileName) || fileName == "." || fileName == "..") {
                fileName = Guid.NewGuid().ToString("N");
            }

            if (string.IsNullOrEmpty(Path.GetExtension(fileName))) {
                fileName += extension;
            }

            if (_exportedImageFileNames.Add(fileName)) {
                return fileName;
            }

            string baseName = Path.GetFileNameWithoutExtension(fileName);
            string fileExtension = Path.GetExtension(fileName);
            for (int suffix = 2; ; suffix++) {
                string candidate = baseName + "-" + suffix.ToString(System.Globalization.CultureInfo.InvariantCulture) + fileExtension;
                if (_exportedImageFileNames.Add(candidate)) {
                    return candidate;
                }
            }
        }

        private static byte[] ReadEmbeddedImageBytes(WordImage image, WordToMarkdownOptions options) {
            long maximumBytes = options.MaxEmbeddedImageBytes;
            if (maximumBytes <= 0) {
                throw new ArgumentOutOfRangeException(nameof(options.MaxEmbeddedImageBytes), "MaxEmbeddedImageBytes must be greater than zero.");
            }

            using Stream input = image.OpenRead();
            if (input.CanSeek && input.Length > maximumBytes) {
                throw new InvalidDataException($"The embedded image exceeds the {maximumBytes}-byte Markdown limit.");
            }

            using var output = new MemoryStream();
            var buffer = new byte[81920];
            long total = 0;
            while (true) {
                int read = input.Read(buffer, 0, buffer.Length);
                if (read == 0) break;
                total += read;
                if (total > maximumBytes) {
                    throw new InvalidDataException($"The embedded image exceeds the {maximumBytes}-byte Markdown limit.");
                }
                output.Write(buffer, 0, read);
            }

            return output.ToArray();
        }

        private static void WriteImageFile(WordImage image, string targetPath) {
            OfficeFileCommit.Write(targetPath, output => {
                using Stream input = !string.IsNullOrEmpty(image.FilePath) && File.Exists(image.FilePath)
                    ? File.OpenRead(image.FilePath)
                    : image.OpenRead();
                input.CopyTo(output);
            });
        }
    }
}
