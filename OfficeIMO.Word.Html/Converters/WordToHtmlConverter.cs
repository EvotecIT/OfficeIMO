using AngleSharp;
using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using AngleSharp.Html.Parser;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Word.Html {
    /// <summary>
    /// Converts <see cref="WordDocument"/> instances into HTML markup.
    /// </summary>
    internal class WordToHtmlConverter {
        public string Convert(WordDocument document, WordToHtmlOptions options) {
            return ConvertAsync(document, options, CancellationToken.None).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Converts the specified document to HTML asynchronously using provided options.
        /// </summary>
        /// <param name="document">Document to convert.</param>
        /// <param name="options">Conversion options controlling HTML output.</param>
        /// <param name="cancellationToken">Token to monitor for cancellation requests.</param>
        /// <returns>HTML representation of the document.</returns>
        public async Task<string> ConvertAsync(WordDocument document, WordToHtmlOptions options, CancellationToken cancellationToken = default) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            options ??= new WordToHtmlOptions();
            cancellationToken.ThrowIfCancellationRequested();

            var context = BrowsingContext.New(Configuration.Default);
            var htmlDoc = await context.OpenNewAsync().ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();

            var head = htmlDoc.Head ?? throw new InvalidOperationException("HTML document missing head element.");
            var body = htmlDoc.Body ?? throw new InvalidOperationException("HTML document missing body element.");

            var charset = htmlDoc.CreateElement("meta");
            charset.SetAttribute("charset", "UTF-8");
            head.AppendChild(charset);

            var props = document.BuiltinDocumentProperties;
            var title = htmlDoc.CreateElement("title");
            var titleText = string.IsNullOrEmpty(props?.Title) ? "Document" : props!.Title!;
            title.TextContent = titleText;
            head.AppendChild(title);

            void AddMeta(string name, string? value) {
                if (!string.IsNullOrEmpty(value)) {
                    var meta = htmlDoc.CreateElement("meta");
                    meta.SetAttribute("name", name);
                    meta.SetAttribute("content", value);
                    head.AppendChild(meta);
                }
            }

            if (props != null) {
                AddMeta("author", props.Creator);
                AddMeta("description", props.Description);
                AddMeta("keywords", props.Keywords);
                AddMeta("subject", props.Subject);
            }

            foreach (var (name, content) in options.AdditionalMetaTags) {
                cancellationToken.ThrowIfCancellationRequested();
                if (!string.IsNullOrEmpty(name)) {
                    var meta = htmlDoc.CreateElement("meta");
                    meta.SetAttribute("name", name);
                    if (!string.IsNullOrEmpty(content)) {
                        meta.SetAttribute("content", content);
                    }
                    head.AppendChild(meta);
                }
            }

            foreach (var (rel, href) in options.AdditionalLinkTags) {
                cancellationToken.ThrowIfCancellationRequested();
                if (!string.IsNullOrEmpty(rel) && !string.IsNullOrEmpty(href)) {
                    var link = htmlDoc.CreateElement("link");
                    link.SetAttribute("rel", rel);
                    link.SetAttribute("href", href);
                    head.AppendChild(link);
                }
            }

            if (options.IncludeDefaultCss) {
                var style = htmlDoc.CreateElement("style");
                style.TextContent = WordHtmlResources.DefaultCss;
                head.AppendChild(style);
            }

            if (!string.IsNullOrEmpty(options.FontFamily)) {
                body.SetAttribute("style", $"font-family:{options.FontFamily}");
            }

            Stack<IElement> listStack = new Stack<IElement>();
            Stack<IElement> itemStack = new Stack<IElement>();

            List<(int Number, WordFootNote Note)> footnotes = new();
            Dictionary<long, int> footnoteMap = new();

            HashSet<string> paragraphStyles = new();
            HashSet<string> runStyles = new();

            void CloseLists() {
                while (listStack.Count > 0) {
                    listStack.Pop();
                }
                while (itemStack.Count > 0) {
                    itemStack.Pop();
                }
            }

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

            void AppendRuns(IElement parent, WordParagraph para, bool processFootnotes = true) {
                var runs = para.GetRuns().ToList();
                List<INode> nodes = new();
                bool inQuote = false;
                IElement? quote = null;
                for (int i = 0; i < runs.Count; i++) {
                    var run = runs[i];
                    if (processFootnotes && options.ExportFootnotes && run.FootNote != null) {
                        var note = run.FootNote;
                        if (string.Equals(run.CharacterStyleId, "HtmlAbbr", StringComparison.OrdinalIgnoreCase) && nodes.Count > 0) {
                            string text = string.Join(string.Empty, note.Paragraphs?.Skip(1).Select(r => r.Text) ?? Enumerable.Empty<string>());
                            var abbr = htmlDoc.CreateElement("abbr");
                            abbr.SetAttribute("title", text);
                            var lastNode = nodes[nodes.Count - 1];
                            abbr.AppendChild(lastNode);
                            nodes[nodes.Count - 1] = abbr;
                        } else {
                            long id = note.ReferenceId ?? 0;
                            if (!footnoteMap.TryGetValue(id, out int number)) {
                                number = footnotes.Count + 1;
                                footnoteMap[id] = number;
                                footnotes.Add((number, note));
                            }
                            var sup = htmlDoc.CreateElement("sup");
                            var a = htmlDoc.CreateElement("a");
                            a.SetAttribute("href", $"#fn{number}");
                            a.SetAttribute("id", $"fnref{number}");
                            a.TextContent = number.ToString();
                            sup.AppendChild(a);
                            nodes.Add(sup);
                        }
                        continue;
                    }

                    if (run.IsImage && run.Image != null) {
                        var imgObj = run.Image;
                        var ext = Path.GetExtension(imgObj.FileName)?.ToLowerInvariant();
                        if (ext == ".svg") {
                            if (options.EmbedImagesAsBase64) {
                                var svgXml = Encoding.UTF8.GetString(imgObj.GetBytes());
                                var parser = new HtmlParser();
                                var fragment = parser.ParseFragment(svgXml, body);
                                var svgElement = fragment.OfType<IElement>().FirstOrDefault();
                                if (svgElement != null) {
                                    nodes.Add(svgElement);
                                }
                            } else {
                                var imgSvg = htmlDoc.CreateElement("img") as IHtmlImageElement;
                                string srcSvg;
                                if (imgObj.IsExternal && imgObj.ExternalUri != null) {
                                    srcSvg = imgObj.ExternalUri.ToString();
                                } else {
                                    srcSvg = string.IsNullOrEmpty(imgObj.FilePath) ? (imgObj.FileName ?? string.Empty) : imgObj.FilePath!;
                                }
                                imgSvg!.Source = srcSvg;
                                if (imgObj.Width.HasValue) imgSvg.DisplayWidth = (int)Math.Round(imgObj.Width.Value);
                                if (imgObj.Height.HasValue) imgSvg.DisplayHeight = (int)Math.Round(imgObj.Height.Value);
                                if (!string.IsNullOrEmpty(imgObj.Description)) {
                                    imgSvg.AlternativeText = imgObj.Description;
                                }
                                nodes.Add(imgSvg);
                            }
                            continue;
                        }

                        var img = htmlDoc.CreateElement("img") as IHtmlImageElement;
                        string src;
                        if (imgObj.IsExternal && imgObj.ExternalUri != null) {
                            src = imgObj.ExternalUri.ToString();
                        } else if (!options.EmbedImagesAsBase64) {
                            src = string.IsNullOrEmpty(imgObj.FilePath) ? (imgObj.FileName ?? string.Empty) : imgObj.FilePath!;
                        } else {
                            var bytes = imgObj.GetBytes();
                            var mime = MimeFromFileName(imgObj.FileName ?? string.Empty);
                            src = $"data:{mime};base64,{System.Convert.ToBase64String(bytes)}";
                        }
                        img!.Source = src;
                        if (imgObj.Width.HasValue) img.DisplayWidth = (int)Math.Round(imgObj.Width.Value);
                        if (imgObj.Height.HasValue) img.DisplayHeight = (int)Math.Round(imgObj.Height.Value);
                        if (!string.IsNullOrEmpty(imgObj.Description)) {
                            img.AlternativeText = imgObj.Description;
                        }
                        nodes.Add(img);
                        continue;
                    }

                    if (string.IsNullOrEmpty(run.Text)) {
                        // Still honor explicit line breaks even when the run carries no text
                        if (run.Break != null && run.PageBreak == null) {
                            nodes.Add(htmlDoc.CreateElement("br"));
                        }
                        continue;
                    }

                    if (string.Equals(run.CharacterStyleId, "HtmlQuote", StringComparison.OrdinalIgnoreCase)) {
                        if (!inQuote) {
                            quote = htmlDoc.CreateElement("q");
                            nodes.Add(quote);
                        } else {
                            quote = null;
                        }
                        inQuote = !inQuote;
                        continue;
                    }

                    // Ensure null-safe text handling
                    INode node = htmlDoc.CreateTextNode(run.Text ?? string.Empty);

                    if (run.Bold) {
                        var strong = htmlDoc.CreateElement("strong");
                        strong.AppendChild(node);
                        node = strong;
                    }

                    if (run.Italic) {
                        var em = htmlDoc.CreateElement("em");
                        em.AppendChild(node);
                        node = em;
                    }

                    if (run.Strike || run.DoubleStrike) {
                        var s = htmlDoc.CreateElement("s");
                        s.AppendChild(node);
                        node = s;
                    }

                    if (run.Underline != null) {
                        var u = htmlDoc.CreateElement("u");
                        u.AppendChild(node);
                        node = u;
                    }

                    if (run.VerticalTextAlignment == VerticalPositionValues.Superscript) {
                        var sup = htmlDoc.CreateElement("sup");
                        sup.AppendChild(node);
                        node = sup;
                    }

                    if (run.VerticalTextAlignment == VerticalPositionValues.Subscript) {
                        var sub = htmlDoc.CreateElement("sub");
                        sub.AppendChild(node);
                        node = sub;
                    }

                    if (run.IsHyperLink && run.Hyperlink != null) {
                        var href = run.Hyperlink.Uri?.ToString();
                        if (string.IsNullOrEmpty(href)) {
                            var anchor = run.Hyperlink.Anchor;
                            if (!string.IsNullOrEmpty(anchor)) {
                                href = "#" + anchor;
                            }
                        }
                        if (!string.IsNullOrEmpty(href)) {
                            var a = htmlDoc.CreateElement("a");
                            a.SetAttribute("href", href);
                            a.AppendChild(node);
                            node = a;
                        }
                        // if href is null/empty, fall back to plain text       
                    }

                    bool handledHtmlStyle = false;
                    if (string.Equals(run.CharacterStyleId, "HtmlCite", StringComparison.OrdinalIgnoreCase)) {
                        var cite = htmlDoc.CreateElement("cite");
                        cite.AppendChild(node);
                        node = cite;
                        handledHtmlStyle = true;
                    } else if (string.Equals(run.CharacterStyleId, "HtmlDfn", StringComparison.OrdinalIgnoreCase)) {
                        var dfn = htmlDoc.CreateElement("dfn");
                        dfn.AppendChild(node);
                        node = dfn;
                        handledHtmlStyle = true;
                    } else if (string.Equals(run.CharacterStyleId, "HtmlTime", StringComparison.OrdinalIgnoreCase)) {
                        var time = htmlDoc.CreateElement("time");
                        string dt = run.Text ?? string.Empty;
                        if (DateTime.TryParse(run.Text, out var parsed)) {
                            dt = parsed.ToString("o");
                        }
                        time.SetAttribute("datetime", dt);
                        time.AppendChild(node);
                        node = time;
                        handledHtmlStyle = true;
                    }

                    if (options.IncludeFontStyles) {
                        var font = run.FontFamily ?? options.FontFamily;
                        if (!string.IsNullOrEmpty(font)) {
                            var span = htmlDoc.CreateElement("span");
                            var value = font.Contains(' ') ? $"\"{font}\"" : font;
                            span.SetAttribute("style", $"font-family:{value}");
                            span.AppendChild(node);
                            node = span;
                        }
                    }

                    // Caps / SmallCaps
                    if (run.CapsStyle == CapsStyle.SmallCaps) {
                        var span = htmlDoc.CreateElement("span");
                        span.SetAttribute("style", "font-variant:small-caps");
                        span.AppendChild(node);
                        node = span;
                    } else if (run.CapsStyle == CapsStyle.Caps) {
                        var span = htmlDoc.CreateElement("span");
                        span.SetAttribute("style", "text-transform:uppercase");
                        span.AppendChild(node);
                        node = span;
                    }

                    if (options.IncludeRunColorStyles || options.IncludeRunHighlightStyles) {
                        var inlineStyles = new List<string>();
                        if (options.IncludeRunColorStyles) {
                            var colorHex = run.ColorHex;
                            if (!string.IsNullOrEmpty(colorHex) &&
                                !string.Equals(colorHex, "auto", StringComparison.OrdinalIgnoreCase)) {
                                var normalized = colorHex.Trim().TrimStart('#').ToLowerInvariant();
                                inlineStyles.Add($"color:#{normalized}");
                            }
                        }
                        if (options.IncludeRunHighlightStyles) {
                            var highlightCss = GetHighlightCss(run.Highlight);
                            if (!string.IsNullOrEmpty(highlightCss)) {
                                inlineStyles.Add($"background-color:{highlightCss}");
                            }
                        }
                        if (inlineStyles.Count > 0) {
                            var span = htmlDoc.CreateElement("span");
                            span.SetAttribute("style", string.Join(";", inlineStyles));
                            span.AppendChild(node);
                            node = span;
                        }
                    }

                    if (options.IncludeRunClasses && !string.IsNullOrEmpty(run.CharacterStyleId) && !handledHtmlStyle) {
                        var spanClass = htmlDoc.CreateElement("span");
                        spanClass.SetAttribute("class", run.CharacterStyleId);
                        spanClass.AppendChild(node);
                        node = spanClass;
                        runStyles.Add(run.CharacterStyleId!);
                    }

                    if (inQuote && quote != null) {
                        quote.AppendChild(node);
                    } else {
                        nodes.Add(node);
                    }

                    // Preserve hard line breaks
                    if (run.Break != null && run.PageBreak == null) {
                        nodes.Add(htmlDoc.CreateElement("br"));
                    }
                }
                foreach (var node in nodes) {
                    cancellationToken.ThrowIfCancellationRequested();
                    parent.AppendChild(node);
                }
            }

            bool IsCodeParagraph(WordParagraph para) {
                if (string.Equals(para.StyleId, "Code", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(para.StyleId, "HTMLPreformatted", StringComparison.OrdinalIgnoreCase)) {
                    return true;
                }
                var runs = FormattingHelper.GetFormattedRuns(para).ToList();
                return runs.Count > 0 && runs.All(r => r.Code);
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

            void AppendParagraph(IElement parent, WordParagraph para) {
                if (para.IsBookmark && para.Bookmark != null) {
                    var name = para.Bookmark.Name ?? string.Empty;
                    var parts = name.Split(new[] { ':' }, 2);
                    if (parts.Length == 2 && IsStructuralTag(parts[0])) {
                        var structEl = htmlDoc.CreateElement(parts[0]);
                        structEl.SetAttribute("id", parts[1]);
                        AppendRuns(structEl, para);
                        parent.AppendChild(structEl);
                        return;
                    }
                }

                if (para.Borders.BottomStyle != null && string.IsNullOrWhiteSpace(para.Text)) {
                    var hr = htmlDoc.CreateElement("hr");
                    parent.AppendChild(hr);
                    return;
                }

                if (IsCodeParagraph(para)) {
                    var pre = htmlDoc.CreateElement("pre");
                    var code = htmlDoc.CreateElement("code");
                    code.TextContent = para.Text ?? string.Empty;
                    pre.AppendChild(code);
                    parent.AppendChild(pre);
                    return;
                }

                int level = para.Style.HasValue ? HeadingStyleMapper.GetLevelForHeadingStyle(para.Style.Value) : 0;
                bool isBlockQuote = (!string.IsNullOrEmpty(para.StyleId) && (string.Equals(para.StyleId, "Quote", StringComparison.OrdinalIgnoreCase) || string.Equals(para.StyleId, "IntenseQuote", StringComparison.OrdinalIgnoreCase)))
                    || (para.IndentationBefore.HasValue && para.IndentationBefore.Value > 0);
                var element = htmlDoc.CreateElement(isBlockQuote ? "blockquote" : (level > 0 ? $"h{level}" : "p"));
                if (options.IncludeParagraphClasses && !string.IsNullOrEmpty(para.StyleId)) {
                    element.SetAttribute("class", para.StyleId);
                    paragraphStyles.Add(para.StyleId!);
                }
                if (para.BiDi) {
                    element.SetAttribute("dir", "rtl");
                }
                // Inline paragraph styles: alignment, shading background, and paragraph borders
                List<string> pStyles = new();
                var alignCss = GetTextAlignCss(para.ParagraphAlignment);
                if (!string.IsNullOrEmpty(alignCss)) {
                    pStyles.Add($"text-align:{alignCss}");
                }
                var pBg = para.ShadingFillColorHex;
                if (!string.IsNullOrEmpty(pBg)) {
                    pStyles.Add($"background-color:#{pBg}");
                }
                var pBorderCss = GetParagraphBorderCss(para);
                if (pBorderCss.Count > 0) {
                    pStyles.AddRange(pBorderCss);
                }
                if (options.IncludeParagraphIndentationStyles) {
                    if (para.IndentationBefore.HasValue && para.IndentationBefore.Value != 0) {
                        pStyles.Add($"margin-left:{FormatTwips(para.IndentationBefore.Value)}");
                    }
                    if (para.IndentationAfter.HasValue && para.IndentationAfter.Value != 0) {
                        pStyles.Add($"margin-right:{FormatTwips(para.IndentationAfter.Value)}");
                    }
                    if (para.IndentationFirstLine.HasValue && para.IndentationFirstLine.Value != 0) {
                        pStyles.Add($"text-indent:{FormatTwips(para.IndentationFirstLine.Value)}");
                    } else if (para.IndentationHanging.HasValue && para.IndentationHanging.Value != 0) {
                        pStyles.Add($"text-indent:{FormatTwips(-para.IndentationHanging.Value)}");
                    }
                }
                if (options.IncludeParagraphSpacingStyles) {
                    if (para.LineSpacingBefore.HasValue && para.LineSpacingBefore.Value != 0) {
                        pStyles.Add($"margin-top:{FormatTwips(para.LineSpacingBefore.Value)}");
                    }
                    if (para.LineSpacingAfter.HasValue && para.LineSpacingAfter.Value != 0) {
                        pStyles.Add($"margin-bottom:{FormatTwips(para.LineSpacingAfter.Value)}");
                    }
                    if (para.LineSpacing.HasValue && para.LineSpacing.Value != 0) {
                        var rule = para.LineSpacingRule;
                        if (rule == null || rule == LineSpacingRuleValues.Auto) {
                            var multiple = para.LineSpacing.Value / 240.0;
                            if (multiple > 0) {
                                pStyles.Add($"line-height:{FormatNumber(multiple)}");
                            }
                        } else {
                            pStyles.Add($"line-height:{FormatTwips(para.LineSpacing.Value)}");
                        }
                    }
                }
                if (pStyles.Count > 0) {
                    element.SetAttribute("style", string.Join(";", pStyles));
                }
                AppendRuns(element, para);
                parent.AppendChild(element);
            }

            string? GetWidthCss(TableWidthUnitValues? type, int? width) {
                if (type == null || width == null) {
                    return null;
                }

                if (type == TableWidthUnitValues.Pct) {
                    return $"{width.Value / 50}%";
                }

                if (type == TableWidthUnitValues.Dxa) {
                    double points = width.Value / 20.0;
                    double pixels = points * 96 / 72;
                    return $"{Math.Round(pixels)}px";
                }

                return null;
            }

            string? GetTextAlignCss(JustificationValues? justification) {
                if (justification == null) {
                    return null;
                }

                if (justification == JustificationValues.Center) {
                    return "center";
                }

                if (justification == JustificationValues.Right) {
                    return "right";
                }

                if (justification == JustificationValues.Left) {
                    return "left";
                }

                if (justification == JustificationValues.Both) {
                    return "justify";
                }

                return null;
            }

            JustificationValues? GetCellAlignment(WordTableCell cell) {
                JustificationValues? align = null;
                foreach (var p in cell.Paragraphs) {
                    if (p.ParagraphAlignment == null) {
                        continue;
                    }
                    if (align == null) {
                        align = p.ParagraphAlignment;
                    } else if (align != p.ParagraphAlignment) {
                        return null;
                    }
                }
                return align;
            }

            string? BuildBorderCss(BorderValues? style, string? colorHex, UInt32Value? size) {
                if (style == null) {
                    return null;
                }

                string cssStyle = "solid";
                if (style == BorderValues.Dashed) {
                    cssStyle = "dashed";
                } else if (style == BorderValues.Dotted) {
                    cssStyle = "dotted";
                } else if (style == BorderValues.Double) {
                    cssStyle = "double";
                }

                string color = !string.IsNullOrEmpty(colorHex) ? $"#{colorHex}" : "black";
                double widthPt = size != null ? size.Value / 8.0 : 1.0;
                double widthPx = widthPt * 96 / 72;
                string width = $"{Math.Round(widthPx)}px";
                return $"{width} {cssStyle} {color}";
            }

            List<string> GetBorderCss(WordTableCell cell) {
                List<string> styles = new();
                var b = cell.Borders;
                if (b == null) {
                    return styles;
                }

                var left = BuildBorderCss(b.LeftStyle, b.LeftColorHex, b.LeftSize);
                var right = BuildBorderCss(b.RightStyle, b.RightColorHex, b.RightSize);
                var top = BuildBorderCss(b.TopStyle, b.TopColorHex, b.TopSize);
                var bottom = BuildBorderCss(b.BottomStyle, b.BottomColorHex, b.BottomSize);

                if (left == null && right == null && top == null && bottom == null) {
                    return styles;
                }

                if (left == top && top == right && right == bottom && left != null) {
                    styles.Add($"border:{left}");
                } else {
                    if (left != null) {
                        styles.Add($"border-left:{left}");
                    }
                    if (right != null) {
                        styles.Add($"border-right:{right}");
                    }
                    if (top != null) {
                        styles.Add($"border-top:{top}");
                    }
                    if (bottom != null) {
                        styles.Add($"border-bottom:{bottom}");
                    }
                }

                return styles;
            }

            List<string> GetParagraphBorderCss(WordParagraph p) {
                List<string> styles = new();
                var b = p.Borders;
                if (b == null) return styles;

                var left = BuildBorderCss(b.LeftStyle, b.LeftColorHex, b.LeftSize);
                var right = BuildBorderCss(b.RightStyle, b.RightColorHex, b.RightSize);
                var top = BuildBorderCss(b.TopStyle, b.TopColorHex, b.TopSize);
                var bottom = BuildBorderCss(b.BottomStyle, b.BottomColorHex, b.BottomSize);

                if (left == null && right == null && top == null && bottom == null) {
                    return styles;
                }
                if (left == top && top == right && right == bottom && left != null) {
                    styles.Add($"border:{left}");
                } else {
                    if (left != null) styles.Add($"border-left:{left}");
                    if (right != null) styles.Add($"border-right:{right}");
                    if (top != null) styles.Add($"border-top:{top}");
                    if (bottom != null) styles.Add($"border-bottom:{bottom}");
                }
                return styles;
            }

            bool CellHasBorder(WordTableCell cell) {
                var b = cell.Borders;
                return b != null && (b.LeftStyle != null || b.RightStyle != null || b.TopStyle != null || b.BottomStyle != null);
            }

            bool TableHasBorder(WordTable table) {
                return table.Rows.Any(r => r.Cells.Any(CellHasBorder));
            }

            void AppendTable(IElement parent, WordTable table) {
                var tableEl = htmlDoc.CreateElement("table");
                var tableStyles = new List<string>();
                var tableWidth = GetWidthCss(table.WidthType, table.Width);
                if (!string.IsNullOrEmpty(tableWidth)) {
                    tableStyles.Add($"width:{tableWidth}");
                }
                if (TableHasBorder(table)) {
                    tableStyles.Add("border:1px solid black");
                    tableStyles.Add("border-collapse:collapse");
                }
                if (tableStyles.Count > 0) {
                    tableEl.SetAttribute("style", string.Join(";", tableStyles));
                }

                for (int r = 0; r < table.Rows.Count; r++) {
                    var row = table.Rows[r];
                    var tr = htmlDoc.CreateElement("tr");
                    for (int c = 0; c < row.Cells.Count; c++) {
                        var cell = row.Cells[c];
                        if (cell.HorizontalMerge == MergedCellValues.Continue || cell.VerticalMerge == MergedCellValues.Continue) {
                            continue;
                        }
                        var td = htmlDoc.CreateElement("td");
                        int colSpan = 1;
                        int rowSpan = 1;
                        if (cell.HorizontalMerge == MergedCellValues.Restart) {
                            int cc = c + 1;
                            while (cc < row.Cells.Count && row.Cells[cc].HorizontalMerge == MergedCellValues.Continue) {
                                colSpan++;
                                cc++;
                            }
                            if (colSpan > 1) {
                                td.SetAttribute("colspan", colSpan.ToString());
                            }
                        }
                        if (cell.VerticalMerge == MergedCellValues.Restart) {
                            int rr = r + 1;
                            while (rr < table.Rows.Count && table.Rows[rr].Cells[c].VerticalMerge == MergedCellValues.Continue) {
                                rowSpan++;
                                rr++;
                            }
                            if (rowSpan > 1) {
                                td.SetAttribute("rowspan", rowSpan.ToString());
                            }
                        }

                        var cellStyles = new List<string>();
                        var width = GetWidthCss(cell.WidthType, cell.Width);
                        if (!string.IsNullOrEmpty(width)) {
                            cellStyles.Add($"width:{width}");
                        }
                        var cellAlignment = GetCellAlignment(cell);
                        var align = GetTextAlignCss(cellAlignment);
                        if (!string.IsNullOrEmpty(align)) {
                            cellStyles.Add($"text-align:{align}");
                        }
                        // Vertical alignment within table cells
                        if (cell.VerticalAlignment != null) {
                            // Avoid enum switch expressions for broad TFM support
                            var s = cell.VerticalAlignment.Value.ToString();
                            string vAlign = "top";
                            if (string.Equals(s, nameof(TableVerticalAlignmentValues.Center), StringComparison.Ordinal)) {
                                vAlign = "middle";
                            } else if (string.Equals(s, nameof(TableVerticalAlignmentValues.Bottom), StringComparison.Ordinal)) {
                                vAlign = "bottom";
                            }
                            cellStyles.Add($"vertical-align:{vAlign}");
                        }
                        var bg = cell.ShadingFillColorHex;
                        if (!string.IsNullOrEmpty(bg)) {
                            cellStyles.Add($"background-color:#{bg}");
                        }
                        var borderCss = GetBorderCss(cell);
                        if (borderCss.Count > 0) {
                            cellStyles.AddRange(borderCss);
                        }
                        if (cellStyles.Count > 0) {
                            td.SetAttribute("style", string.Join(";", cellStyles));
                        }

                        for (int pIdx = 0; pIdx < cell.Paragraphs.Count; pIdx++) {
                            var p = cell.Paragraphs[pIdx];
                            if (IsCodeParagraph(p)) {
                                List<string> lines = new();
                                lines.Add(p.Text);
                                while (pIdx + 1 < cell.Paragraphs.Count && IsCodeParagraph(cell.Paragraphs[pIdx + 1])) {
                                    lines.Add(cell.Paragraphs[pIdx + 1].Text);
                                    pIdx++;
                                }
                                var pre = htmlDoc.CreateElement("pre");
                                var code = htmlDoc.CreateElement("code");
                                code.TextContent = string.Join("\n", lines);
                                pre.AppendChild(code);
                                td.AppendChild(pre);
                            } else {
                                AppendParagraph(td, p);
                            }
                        }

                        if (cell.HasNestedTables) {
                            foreach (var nested in cell.NestedTables) {
                                cancellationToken.ThrowIfCancellationRequested();
                                AppendTable(td, nested);
                            }
                        }

                        tr.AppendChild(td);
                    }
                    tableEl.AppendChild(tr);
                }
                parent.AppendChild(tableEl);
            }

            var formatMap = new Dictionary<NumberFormatValues, (string? Type, string Css)>{
                { NumberFormatValues.Decimal, ("1", "decimal") },
                { NumberFormatValues.DecimalZero, (null, "decimal-leading-zero") },
                { NumberFormatValues.LowerLetter, ("a", "lower-alpha") },
                { NumberFormatValues.UpperLetter, ("A", "upper-alpha") },
                { NumberFormatValues.LowerRoman, ("i", "lower-roman") },
                { NumberFormatValues.UpperRoman, ("I", "upper-roman") },
            };

            string? GetListStyle(DocumentTraversal.ListInfo info) {
                var format = info.NumberFormat;
                if (format == NumberFormatValues.Bullet) {
                    return info.LevelText switch {
                        "o" or "◦" => "circle",
                        "■" or "§" => "square",
                        _ => "disc",
                    };
                }
                if (format != null && formatMap.TryGetValue(format.Value, out var map)) {
                    return map.Css;
                }
                return null;
            }

            string? GetListType(DocumentTraversal.ListInfo info) {
                var format = info.NumberFormat;
                if (format == NumberFormatValues.Bullet) {
                    return info.LevelText switch {
                        "o" or "◦" => "circle",
                        "■" or "§" => "square",
                        _ => "disc",
                    };
                }
                if (format != null && formatMap.TryGetValue(format.Value, out var map)) {
                    return map.Type;
                }
                return null;
            }

            var listIndices = DocumentTraversal.BuildListIndices(document);

            var processedParagraphs = new HashSet<WordParagraph>();
            foreach (var section in DocumentTraversal.EnumerateSections(document)) {
                cancellationToken.ThrowIfCancellationRequested();
                var elements = section.Elements;
                if (elements == null || elements.Count == 0) {
                    // Fallback: compose elements from paragraphs and tables when section enumeration yields none
                    var composed = new List<WordElement>(section.Paragraphs.Count + section.Tables.Count);
                    composed.AddRange(section.Paragraphs);
                    composed.AddRange(section.Tables);
                    elements = composed;
                }
                if (elements == null) {
                    continue;
                }
                for (int idx = 0; idx < elements.Count; idx++) {
                    var element = elements[idx];
                    if (element is WordParagraph paragraph) {
                        // Render each underlying OpenXml paragraph exactly once.
                        // Prefer the bookmark-bearing wrapper when multiple wrappers exist for the same paragraph.
                        if (processedParagraphs.Contains(paragraph)) {
                            continue;
                        }
                        if (!paragraph.IsBookmark) {
                            // Look ahead for a sibling wrapper (same underlying paragraph) that carries a bookmark
                            for (int j = idx + 1; j < elements.Count; j++) {
                                if (elements[j] is WordParagraph sibling && sibling.Equals(paragraph)) {
                                    if (sibling.IsBookmark) { paragraph = sibling; }
                                    continue;
                                }
                                break;
                            }
                        }
                        processedParagraphs.Add(paragraph);
                        var listInfo = DocumentTraversal.GetListInfo(paragraph);
                        if (listInfo != null) {
                            int level = listInfo.Value.Level;
                            while (listStack.Count > level) {
                                listStack.Pop();
                                itemStack.Pop();
                            }
                            while (listStack.Count <= level) {
                                bool ordered = listInfo.Value.Ordered;
                                var listTag = ordered ? "ol" : "ul";
                                var listEl = htmlDoc.CreateElement(listTag);
                                if (ordered) {
                                    // Continue numbering across gaps by using the numeric index of the current item when available
                                    if (listIndices.TryGetValue(paragraph, out var indexInfo)) {
                                        listEl.SetAttribute("start", indexInfo.Index.ToString());
                                    } else {
                                        listEl.SetAttribute("start", listInfo.Value.Start.ToString());
                                    }
                                }
                                var typeAttr = GetListType(listInfo.Value);
                                if (!string.IsNullOrEmpty(typeAttr)) {
                                    listEl.SetAttribute("type", typeAttr);
                                }
                                if (options.IncludeListStyles) {
                                    var css = GetListStyle(listInfo.Value);
                                    if (!string.IsNullOrEmpty(css)) {
                                        listEl.SetAttribute("style", $"list-style-type:{css}");
                                    }
                                }
                                if (itemStack.Count > 0) {
                                    itemStack.Peek().AppendChild(listEl);
                                } else {
                                    body.AppendChild(listEl);
                                }
                                listStack.Push(listEl);
                            }
                            while (itemStack.Count > level) {
                                itemStack.Pop();
                            }
                            var li = htmlDoc.CreateElement("li");
                            listStack.Peek().AppendChild(li);
                            itemStack.Push(li);
                            AppendRuns(li, paragraph);
                        } else {
                            CloseLists();
                            if (paragraph.IsImage && idx + 1 < elements.Count && elements[idx + 1] is WordParagraph captionPara && string.Equals(captionPara.StyleId, "Caption", StringComparison.OrdinalIgnoreCase)) {
                                var figure = htmlDoc.CreateElement("figure");
                                AppendRuns(figure, paragraph);
                                var figCap = htmlDoc.CreateElement("figcaption");
                                if (options.IncludeParagraphClasses && !string.IsNullOrEmpty(captionPara.StyleId)) {
                                    figCap.SetAttribute("class", captionPara.StyleId);
                                    paragraphStyles.Add(captionPara.StyleId!);
                                }
                                AppendRuns(figCap, captionPara);
                                figure.AppendChild(figCap);
                                body.AppendChild(figure);
                                idx++;
                            } else if (IsCodeParagraph(paragraph)) {
                                List<string> lines = new();
                                lines.Add(paragraph.Text);
                                while (idx + 1 < elements.Count && elements[idx + 1] is WordParagraph nextPara && DocumentTraversal.GetListInfo(nextPara) == null && IsCodeParagraph(nextPara)) {
                                    lines.Add(nextPara.Text);
                                    idx++;
                                }
                                var pre = htmlDoc.CreateElement("pre");
                                var code = htmlDoc.CreateElement("code");
                                code.TextContent = string.Join("\n", lines);
                                pre.AppendChild(code);
                                body.AppendChild(pre);
                            } else {
                                AppendParagraph(body, paragraph);
                            }
                        }
                    } else if (element is WordTable table) {
                        CloseLists();
                        AppendTable(body, table);
                    }
                }
            }

            CloseLists();

            if (options.ExportFootnotes && footnotes.Count > 0) {
                var footSection = htmlDoc.CreateElement("section");
                footSection.SetAttribute("class", "footnotes");
                var hr = htmlDoc.CreateElement("hr");
                footSection.AppendChild(hr);
                var ol = htmlDoc.CreateElement("ol");
                foreach (var (number, note) in footnotes) {
                    cancellationToken.ThrowIfCancellationRequested();
                    var li = htmlDoc.CreateElement("li");
                    li.SetAttribute("id", $"fn{number}");
                    var p = htmlDoc.CreateElement("p");
                    string text = string.Join(string.Empty, note.Paragraphs?.Skip(1).Select(r => r.Text) ?? Enumerable.Empty<string>());
                    p.TextContent = text;
                    li.AppendChild(p);
                    ol.AppendChild(li);
                }
                footSection.AppendChild(ol);
                body.AppendChild(footSection);
            }

            if (paragraphStyles.Count > 0 || runStyles.Count > 0) {
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

            return htmlDoc.DocumentElement.OuterHtml;
        }
    }
}
