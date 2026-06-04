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
    internal partial class WordToHtmlConverter {
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

            AppendHeadMetadata(document, htmlDoc, head, options, cancellationToken);

            if (!string.IsNullOrEmpty(options.FontFamily)) {
                body.SetAttribute("style", $"font-family:{options.FontFamily}");
            }

            Stack<IElement> listStack = new Stack<IElement>();
            Stack<IElement> itemStack = new Stack<IElement>();

            List<(int Number, WordFootNote Note)> footnotes = new();
            List<(int Number, WordEndNote Note)> endnotes = new();
            List<(int Number, WordComment Comment)> comments = new();
            Dictionary<long, int> footnoteMap = new();
            Dictionary<long, int> endnoteMap = new();
            Dictionary<string, WordComment> commentsById = options.ExportComments
                ? document.Comments
                    .Where(comment => !string.IsNullOrEmpty(comment.Id))
                    .ToDictionary(comment => comment.Id!, comment => comment, StringComparer.OrdinalIgnoreCase)
                : new Dictionary<string, WordComment>(StringComparer.OrdinalIgnoreCase);
            Dictionary<string, int> commentMap = new(StringComparer.OrdinalIgnoreCase);

            HashSet<string> paragraphStyles = new();
            HashSet<string> runStyles = new();
            HashSet<HtmlListDefinition> listDefinitions = new();
            int formListIndex = 0;

            void CloseLists() {
                while (listStack.Count > 0) {
                    listStack.Pop();
                }
                while (itemStack.Count > 0) {
                    itemStack.Pop();
                }
            }


            void AppendRuns(IElement parent, WordParagraph para, bool processNotes = true) {
                var runs = para.GetRuns().ToList();
                List<INode> nodes = new();
                bool inQuote = false;
                IElement? quote = null;
                for (int i = 0; i < runs.Count; i++) {
                    var run = runs[i];
                    if (TryAppendNoteReference(htmlDoc, run, options, processNotes, nodes, footnotes, footnoteMap, endnotes, endnoteMap)) {
                        continue;
                    }

                    if (TryAppendCommentReference(htmlDoc, run, options, commentsById, comments, commentMap, nodes)) {
                        continue;
                    }

                    if (run.IsCheckBox && run.CheckBox != null) {
                        nodes.Add(CreateCheckBoxInput(htmlDoc, run.CheckBox));
                        continue;
                    }

                    if (run.IsDropDownList && run.DropDownList != null) {
                        nodes.Add(CreateDropDownListSelect(htmlDoc, run.DropDownList));
                        continue;
                    }

                    if (run.IsComboBox && run.ComboBox != null) {
                        formListIndex++;
                        nodes.AddRange(CreateComboBoxNodes(htmlDoc, run.ComboBox, formListIndex));
                        continue;
                    }

                    if (run.IsDatePicker && run.DatePicker != null) {
                        nodes.Add(CreateDatePickerInput(htmlDoc, run.DatePicker));
                        continue;
                    }

                    if (run.IsStructuredDocumentTag && run.StructuredDocumentTag != null && !run.IsPictureControl && !run.IsRepeatingSection) {
                        nodes.Add(CreateStructuredDocumentTagInput(htmlDoc, run.StructuredDocumentTag));
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

                    // Still honor explicit line breaks even when the run carries no text
                    if (run.Break != null && run.PageBreak == null) {
                        nodes.Add(htmlDoc.CreateElement("br"));
                    }
                    if (string.IsNullOrEmpty(run.Text)) {
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
                    } else if (string.Equals(run.CharacterStyleId, "HtmlCode", StringComparison.OrdinalIgnoreCase)) {
                        var code = htmlDoc.CreateElement("code");
                        code.AppendChild(node);
                        node = code;
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

                    if (run.FontSize != null) {
                        var span = htmlDoc.CreateElement("span");
                        span.SetAttribute("style", $"font-size:{run.FontSize.Value}pt");
                        span.AppendChild(node);
                        node = span;
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
                    ApplyBookmarkId(hr, para);
                    parent.AppendChild(hr);
                    return;
                }

                if (IsCodeParagraph(para)) {
                    var pre = htmlDoc.CreateElement("pre");
                    ApplyBookmarkId(pre, para);
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
                ApplyBookmarkId(element, para);
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
                if (options.IncludeTableColumnGroups) {
                    AppendColumnGroup(htmlDoc, tableEl, table);
                }

                int headerRowCount = 0;
                while (headerRowCount < table.Rows.Count && table.Rows[headerRowCount].RepeatHeaderRowAtTheTopOfEachPage) {
                    headerRowCount++;
                }
                IElement? thead = null;
                IElement? tbody = null;

                for (int r = 0; r < table.Rows.Count; r++) {
                    var row = table.Rows[r];
                    var tr = htmlDoc.CreateElement("tr");
                    bool isHeaderRow = headerRowCount > 0 && r < headerRowCount;
                    for (int c = 0; c < row.Cells.Count; c++) {
                        var cell = row.Cells[c];
                        if (cell.HorizontalMerge == MergedCellValues.Continue || cell.VerticalMerge == MergedCellValues.Continue) {
                            continue;
                        }
                        var cellElement = htmlDoc.CreateElement(isHeaderRow ? "th" : "td");
                        int colSpan = 1;
                        int rowSpan = 1;
                        if (cell.HorizontalMerge == MergedCellValues.Restart) {
                            int cc = c + 1;
                            while (cc < row.Cells.Count && row.Cells[cc].HorizontalMerge == MergedCellValues.Continue) {
                                colSpan++;
                                cc++;
                            }
                            if (colSpan > 1) {
                                cellElement.SetAttribute("colspan", colSpan.ToString());
                            }
                        }
                        if (cell.VerticalMerge == MergedCellValues.Restart) {
                            int rr = r + 1;
                            while (rr < table.Rows.Count && table.Rows[rr].Cells[c].VerticalMerge == MergedCellValues.Continue) {
                                rowSpan++;
                                rr++;
                            }
                            if (rowSpan > 1) {
                                cellElement.SetAttribute("rowspan", rowSpan.ToString());
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
                            cellElement.SetAttribute("style", string.Join(";", cellStyles));
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
                                cellElement.AppendChild(pre);
                            } else {
                                AppendParagraph(cellElement, p);
                            }
                        }

                        if (cell.HasNestedTables) {
                            foreach (var nested in cell.NestedTables) {
                                cancellationToken.ThrowIfCancellationRequested();
                                AppendTable(cellElement, nested);
                            }
                        }

                        tr.AppendChild(cellElement);
                    }
                    if (headerRowCount == 0) {
                        tableEl.AppendChild(tr);
                    } else if (isHeaderRow) {
                        if (thead == null) {
                            thead = htmlDoc.CreateElement("thead");
                            tableEl.AppendChild(thead);
                        }
                        thead.AppendChild(tr);
                    } else {
                        if (tbody == null) {
                            tbody = htmlDoc.CreateElement("tbody");
                            tableEl.AppendChild(tbody);
                        }
                        tbody.AppendChild(tr);
                    }
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
                { NumberFormatValues.RussianLower, (null, "lower-russian") },
                { NumberFormatValues.RussianUpper, (null, "upper-russian") },
                { NumberFormatValues.Hebrew1, (null, "hebrew") },
                { NumberFormatValues.Hebrew2, (null, "hebrew-2") },
                { NumberFormatValues.ArabicAlpha, (null, "arabic-alpha") },
                { NumberFormatValues.ArabicAbjad, (null, "arabic-abjad") },
                { NumberFormatValues.Aiueo, (null, "hiragana") },
                { NumberFormatValues.Iroha, (null, "hiragana-iroha") },
                { NumberFormatValues.AiueoFullWidth, (null, "katakana") },
                { NumberFormatValues.IrohaFullWidth, (null, "katakana-iroha") },
            };

            string? GetListStyle(DocumentTraversal.ListInfo info) {
                var format = info.NumberFormat;
                if (format == NumberFormatValues.Bullet) {
                    return info.LevelText switch {
                        "o" or "◦" => "circle",
                        "■" or "§" => "square",
                        "-" => "'-'",
                        "\u2013" => "'\\2013'",
                        "\u2014" => "'\\2014'",
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
                        "-" or "\u2013" or "\u2014" => null,
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
            int sectionIndex = 0;
            foreach (var section in DocumentTraversal.EnumerateSections(document)) {
                cancellationToken.ThrowIfCancellationRequested();
                IElement sectionParent = body;
                if (options.IncludeSectionMetadata) {
                    sectionParent = CreateSectionElement(htmlDoc, section, sectionIndex, sectionIndex == 0);
                    body.AppendChild(sectionParent);
                }
                AppendHeaderFooterRegions(htmlDoc, sectionParent, section, sectionIndex, true, AppendParagraph, AppendTable, options, cancellationToken);

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
                                var listStyle = GetListStyle(listInfo.Value);
                                if (options.IncludeListStyles && !string.IsNullOrEmpty(listStyle)) {
                                    listEl.SetAttribute("style", $"list-style-type:{listStyle}");
                                }
                                if (options.IncludeListDefinitions) {
                                    ApplyListDefinition(listEl, listInfo.Value, listStyle, listDefinitions);
                                }
                                if (itemStack.Count > 0) {
                                    itemStack.Peek().AppendChild(listEl);
                                } else {
                                    sectionParent.AppendChild(listEl);
                                }
                                listStack.Push(listEl);
                            }
                            while (itemStack.Count > level) {
                                itemStack.Pop();
                            }
                            var li = htmlDoc.CreateElement("li");
                            ApplyBookmarkId(li, paragraph);
                            listStack.Peek().AppendChild(li);
                            itemStack.Push(li);
                            AppendRuns(li, paragraph);
                        } else {
                            CloseLists();
                            if (paragraph.IsImage && idx + 1 < elements.Count && elements[idx + 1] is WordParagraph captionPara && string.Equals(captionPara.StyleId, "Caption", StringComparison.OrdinalIgnoreCase)) {
                                var figure = htmlDoc.CreateElement("figure");
                                ApplyBookmarkId(figure, paragraph);
                                AppendRuns(figure, paragraph);
                                var figCap = htmlDoc.CreateElement("figcaption");
                                if (options.IncludeParagraphClasses && !string.IsNullOrEmpty(captionPara.StyleId)) {
                                    figCap.SetAttribute("class", captionPara.StyleId);
                                    paragraphStyles.Add(captionPara.StyleId!);
                                }
                                AppendRuns(figCap, captionPara);
                                figure.AppendChild(figCap);
                                sectionParent.AppendChild(figure);
                                idx++;
                            } else if (IsCodeParagraph(paragraph)) {
                                List<string> lines = new();
                                lines.Add(paragraph.Text);
                                while (idx + 1 < elements.Count && elements[idx + 1] is WordParagraph nextPara && DocumentTraversal.GetListInfo(nextPara) == null && IsCodeParagraph(nextPara)) {
                                    lines.Add(nextPara.Text);
                                    idx++;
                                }
                                var pre = htmlDoc.CreateElement("pre");
                                ApplyBookmarkId(pre, paragraph);
                                var code = htmlDoc.CreateElement("code");
                                code.TextContent = string.Join("\n", lines);
                                pre.AppendChild(code);
                                sectionParent.AppendChild(pre);
                            } else {
                                AppendParagraph(sectionParent, paragraph);
                            }
                        }
                    } else if (element is WordTable table) {
                        CloseLists();
                        AppendTable(sectionParent, table);
                    }
                }
                if (options.ExportHeadersAndFooters) {
                    CloseLists();
                    AppendHeaderFooterRegions(htmlDoc, sectionParent, section, sectionIndex, false, AppendParagraph, AppendTable, options, cancellationToken);
                }
                if (options.IncludeSectionMetadata) {
                    CloseLists();
                }
                sectionIndex++;
            }

            CloseLists();

            AppendFootnotes(htmlDoc, body, footnotes, options, cancellationToken);
            AppendEndnotes(htmlDoc, body, endnotes, options, cancellationToken);
            AppendComments(htmlDoc, body, comments, options, cancellationToken);
            AppendListDefinitions(htmlDoc, head, listDefinitions, cancellationToken);
            AppendStyleDefinitions(document, htmlDoc, head, paragraphStyles, runStyles, cancellationToken);

            return htmlDoc.DocumentElement.OuterHtml;
        }
    }
}
