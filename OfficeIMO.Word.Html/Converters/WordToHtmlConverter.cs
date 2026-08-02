using AngleSharp.Dom;
using AngleSharp.Html;
using AngleSharp.Html.Dom;
using AngleSharp.Html.Parser;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Globalization;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeIMO.Word.Html {
    /// <summary>
    /// Converts <see cref="WordDocument"/> instances into HTML markup.
    /// </summary>
    internal partial class WordToHtmlConverter {
        public string Convert(WordDocument document, WordToHtmlOptions options) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            options ??= new WordToHtmlOptions();
            ExportInspection exportInspection = InspectExport(document, options);
            ReportKnownExportLimitations(document, options, exportInspection);
            var htmlDoc = new HtmlParser().ParseDocument("<!DOCTYPE html><html><head></head><body></body></html>");
            CancellationToken cancellationToken = CancellationToken.None;
            long embeddedImageBytes = 0;

            long MeasureCurrentHtmlCharacters() {
                using var countingWriter = new CountingHtmlWriter();
                htmlDoc.DocumentElement.ToHtml(countingWriter, HtmlMarkupFormatter.Instance);
                return countingWriter.CharacterCount;
            }

            long GetBase64ImageOutputCharacters(long imageBytes, string mime) {
                long base64Characters = imageBytes > (long.MaxValue / 4L) * 3L
                    ? long.MaxValue
                    : ((imageBytes + 2L) / 3L) * 4L;
                long prefixCharacters = "data:;base64,".Length + mime.Length;
                return SaturatingAdd(prefixCharacters, base64Characters);
            }

            void ReserveImageOutputCharacters(
                long imageOutputCharacters,
                ref long reservedImageOutputCharacters,
                string message,
                string source) {
                if (imageOutputCharacters <= reservedImageOutputCharacters) return;
                ReserveOutputCharacters(
                    htmlDoc,
                    imageOutputCharacters - reservedImageOutputCharacters,
                    message,
                    source);
                reservedImageOutputCharacters = imageOutputCharacters;
            }

            byte[] ReadEmbeddedImageBytes(WordImage image, string source, string? base64Mime = null, bool inlineMarkup = false) {
                using Stream input = image.OpenRead();
                long reservedImageOutputCharacters = 0;
                if (input.CanSeek && input.Length > options.MaxEmbeddedImageBytes) {
                    ThrowExportLimitExceeded(options, "WordImageSizeLimitExceeded", "A Word image exceeds the configured per-image HTML export limit.", source, input.Length, options.MaxEmbeddedImageBytes);
                }
                if (input.CanSeek && input.Length > options.MaxTotalEmbeddedImageBytes - embeddedImageBytes) {
                    ThrowExportLimitExceeded(options, "WordImageTotalSizeLimitExceeded", "Embedded Word images exceed the configured aggregate HTML export limit.", source, SaturatingAdd(embeddedImageBytes, input.Length), options.MaxTotalEmbeddedImageBytes);
                }
                if (base64Mime != null && input.CanSeek) {
                    ReserveImageOutputCharacters(
                        GetBase64ImageOutputCharacters(input.Length, base64Mime),
                        ref reservedImageOutputCharacters,
                        "An embedded Word image cannot fit within the configured HTML output-character limit.",
                        source);
                }
                if (inlineMarkup && input.CanSeek) {
                    ReserveImageOutputCharacters(
                        input.Length,
                        ref reservedImageOutputCharacters,
                        "An inline SVG image cannot fit within the configured HTML output-character limit.",
                        source);
                }

                int capacity = input.CanSeek && input.Length <= int.MaxValue ? (int)input.Length : 0;
                using var output = new MemoryStream(capacity);
                byte[] buffer = new byte[81920];
                long imageBytes = 0;
                int read;
                while ((read = input.Read(buffer, 0, buffer.Length)) > 0) {
                    long nextImageBytes = imageBytes + read;
                    if (nextImageBytes > options.MaxEmbeddedImageBytes) {
                        ThrowExportLimitExceeded(options, "WordImageSizeLimitExceeded", "A Word image exceeds the configured per-image HTML export limit.", source, nextImageBytes, options.MaxEmbeddedImageBytes);
                    }
                    if (nextImageBytes > options.MaxTotalEmbeddedImageBytes - embeddedImageBytes) {
                        ThrowExportLimitExceeded(options, "WordImageTotalSizeLimitExceeded", "Embedded Word images exceed the configured aggregate HTML export limit.", source, SaturatingAdd(embeddedImageBytes, nextImageBytes), options.MaxTotalEmbeddedImageBytes);
                    }
                    if (base64Mime != null) {
                        ReserveImageOutputCharacters(
                            GetBase64ImageOutputCharacters(nextImageBytes, base64Mime),
                            ref reservedImageOutputCharacters,
                            "An embedded Word image cannot fit within the configured HTML output-character limit.",
                            source);
                    }
                    if (inlineMarkup) {
                        ReserveImageOutputCharacters(
                            nextImageBytes,
                            ref reservedImageOutputCharacters,
                            "An inline SVG image cannot fit within the configured HTML output-character limit.",
                            source);
                    }
                    output.Write(buffer, 0, read);
                    imageBytes = nextImageBytes;
                }

                embeddedImageBytes += imageBytes;
                if (base64Mime != null) {
                    ReserveImageOutputCharacters(
                        GetBase64ImageOutputCharacters(imageBytes, base64Mime),
                        ref reservedImageOutputCharacters,
                        "An embedded Word image cannot fit within the configured HTML output-character limit.",
                        source);
                } else if (inlineMarkup) {
                    ReserveImageOutputCharacters(
                        imageBytes,
                        ref reservedImageOutputCharacters,
                        "An inline SVG image cannot fit within the configured HTML output-character limit.",
                        source);
                }
                byte[] bytes = output.ToArray();
                return bytes;
            }

            long SaturatingAdd(long left, long right) => left > long.MaxValue - right ? long.MaxValue : left + right;

            var head = htmlDoc.Head ?? throw new InvalidOperationException("HTML document missing head element.");
            var body = htmlDoc.Body ?? throw new InvalidOperationException("HTML document missing body element.");

            RegisterOutputConstructionBudget(
                htmlDoc,
                options,
                SaturatingAdd(MeasureCurrentHtmlCharacters(), exportInspection.OutputConstructionCharacters));

            AppendHeadMetadata(document, htmlDoc, head, options, CancellationToken.None);

            if (!string.IsNullOrEmpty(options.FontFamily)) {
                SetOutputAttribute(
                    htmlDoc,
                    body,
                    "style",
                    $"font-family:{options.FontFamily}",
                    "DocumentFontStyle");
            }

            Stack<IElement> listStack = new Stack<IElement>();
            Stack<IElement> itemStack = new Stack<IElement>();
            Stack<int> listNumberStack = new Stack<int>();

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
                while (listNumberStack.Count > 0) {
                    listNumberStack.Pop();
                }
            }


            void AppendRuns(IElement parent, WordParagraph para, bool processNotes = true) {
                var runs = para.GetRuns().ToList();
                var paragraphChildren = para._paragraph.ChildElements.ToList();
                IReadOnlyList<WordEquationOccurrence> equations = WordEquation.GetOccurrences(para._document, para._paragraph);
                List<INode> nodes = new();
                int nextEquation = 0;
                var expandedEquationContainers = new HashSet<DocumentFormat.OpenXml.OpenXmlElement>();
                bool inQuote = false;
                IElement? quote = null;

                void AppendNode(INode node) {
                    if (inQuote && quote != null) {
                        quote.AppendChild(node);
                    } else {
                        nodes.Add(node);
                    }
                }

                bool AppendRunArtifacts(WordParagraph run, List<INode> target, DocumentFormat.OpenXml.OpenXmlElement? artifactElement = null) {
                    bool includeAll = artifactElement == null;
                    if ((includeAll || artifactElement is FootnoteReference || artifactElement is EndnoteReference) &&
                        TryAppendNoteReference(htmlDoc, run, options, processNotes, target, footnotes, footnoteMap, endnotes, endnoteMap)) {
                        return true;
                    }

                    if ((includeAll || artifactElement is CommentReference) &&
                        TryAppendCommentReference(htmlDoc, run, options, commentsById, comments, commentMap, target)) {
                        return true;
                    }

                    if ((includeAll || artifactElement is SdtRun) && run.IsCheckBox && run.CheckBox != null) {
                        target.Add(CreateCheckBoxInput(htmlDoc, run.CheckBox));
                        return true;
                    }

                    if ((includeAll || artifactElement is SdtRun) && run.IsDropDownList && run.DropDownList != null) {
                        target.Add(CreateDropDownListSelect(htmlDoc, run.DropDownList));
                        return true;
                    }

                    if ((includeAll || artifactElement is SdtRun) && run.IsComboBox && run.ComboBox != null) {
                        formListIndex++;
                        target.AddRange(CreateComboBoxNodes(htmlDoc, run.ComboBox, formListIndex));
                        return true;
                    }

                    if ((includeAll || artifactElement is SdtRun) && run.IsDatePicker && run.DatePicker != null) {
                        target.Add(CreateDatePickerInput(htmlDoc, run.DatePicker));
                        return true;
                    }

                    if ((includeAll || artifactElement is SdtRun) && run.IsStructuredDocumentTag && run.StructuredDocumentTag != null && !run.IsPictureControl && !run.IsRepeatingSection) {
                        target.Add(CreateStructuredDocumentTagInput(htmlDoc, run.StructuredDocumentTag));
                        return true;
                    }

                    if ((includeAll || artifactElement is DocumentFormat.OpenXml.Wordprocessing.Drawing || artifactElement is DocumentFormat.OpenXml.Vml.ImageData) &&
                        run.IsImage && run.Image != null) {
                        var imgObj = run.Image;
                        var ext = Path.GetExtension(imgObj.FileName)?.ToLowerInvariant();
                        if (ext == ".svg") {
                            if (options.EmbedImagesAsBase64) {
                                var svgXml = Encoding.UTF8.GetString(ReadEmbeddedImageBytes(imgObj, imgObj.FileName ?? "image.svg", inlineMarkup: true));
                                var parser = new HtmlParser();
                                var fragment = parser.ParseFragment(svgXml, body);
                                var svgElement = fragment.OfType<IElement>().FirstOrDefault();
                                if (svgElement != null) {
                                    target.Add(svgElement);
                                }
                            } else {
                                var imgSvg = (IHtmlImageElement)CreateOutputElement(htmlDoc, "img");
                                string srcSvg;
                                if (imgObj.IsExternal && imgObj.ExternalUri != null) {
                                    srcSvg = imgObj.ExternalUri.ToString();
                                } else {
                                    srcSvg = string.IsNullOrEmpty(imgObj.FilePath) ? (imgObj.FileName ?? string.Empty) : imgObj.FilePath!;
                                }
                                SetOutputAttribute(htmlDoc, imgSvg, "src", srcSvg, "Image:src");
                                if (imgObj.Width.HasValue) imgSvg.DisplayWidth = (int)Math.Round(imgObj.Width.Value);
                                if (imgObj.Height.HasValue) imgSvg.DisplayHeight = (int)Math.Round(imgObj.Height.Value);
                                if (!string.IsNullOrEmpty(imgObj.Description)) {
                                    imgSvg.AlternativeText = imgObj.Description;
                                }
                                if (!string.IsNullOrEmpty(imgObj.Title)) {
                                    imgSvg.SetAttribute("title", imgObj.Title!);
                                }
                                target.Add(imgSvg);
                            }
                        } else {
                            var img = (IHtmlImageElement)CreateOutputElement(htmlDoc, "img");
                            string src;
                            if (imgObj.IsExternal && imgObj.ExternalUri != null) {
                                src = imgObj.ExternalUri.ToString();
                            } else if (!options.EmbedImagesAsBase64) {
                                src = string.IsNullOrEmpty(imgObj.FilePath) ? (imgObj.FileName ?? string.Empty) : imgObj.FilePath!;
                            } else {
                                var mime = MimeFromFileName(imgObj.FileName ?? string.Empty);
                                var bytes = ReadEmbeddedImageBytes(imgObj, imgObj.FileName ?? "image", mime);
                                src = $"data:{mime};base64,{System.Convert.ToBase64String(bytes)}";
                            }
                            if (imgObj.IsExternal || !options.EmbedImagesAsBase64) {
                                SetOutputAttribute(htmlDoc, img, "src", src, "Image:src");
                            } else {
                                img.Source = src;
                            }
                            if (imgObj.Width.HasValue) img.DisplayWidth = (int)Math.Round(imgObj.Width.Value);
                            if (imgObj.Height.HasValue) img.DisplayHeight = (int)Math.Round(imgObj.Height.Value);
                            if (!string.IsNullOrEmpty(imgObj.Description)) {
                                img.AlternativeText = imgObj.Description;
                            }
                            if (!string.IsNullOrEmpty(imgObj.Title)) {
                                img.SetAttribute("title", imgObj.Title!);
                            }
                            target.Add(img);
                        }
                        return true;
                    }

                    bool appendedBreak = false;
                    if ((includeAll || artifactElement is Break || artifactElement is CarriageReturn) && run.Break != null && run.PageBreak == null) {
                        target.Add(CreateOutputElement(htmlDoc, "br"));
                        appendedBreak = true;
                    }
                    if (TryCreateRubyNode(htmlDoc, run, out var rubyNode)) {
                        target.Add(rubyNode);
                        return true;
                    }

                    return appendedBreak && string.IsNullOrEmpty(run.Text);
                }

                List<INode> CreateExpandedEquationContainerNodes(
                    DocumentFormat.OpenXml.OpenXmlElement container,
                    IReadOnlyList<WordEquationOccurrence> coveringEquations,
                    WordParagraph fallbackRun) {
                    var expandedNodes = new List<INode>();
                    IElement? hyperlinkNode = container is Hyperlink hyperlink
                        ? CreateEquationHyperlinkNode(
                            htmlDoc,
                            new WordHyperLink(para._document, para._paragraph, hyperlink))
                        : null;

                    foreach (WordEquationContentSegment segment in WordEquation.GetVisibleContentSegments(container, coveringEquations)) {
                        WordParagraph sourceRun = segment.CreateSourceParagraph(
                            para._document,
                            para._paragraph,
                            fallbackRun);
                        if (segment.Equation != null) {
                            IElement? mathNode = CreateEquationNode(htmlDoc, parent, segment.Equation, options);
                            if (mathNode != null &&
                                hyperlinkNode == null &&
                                sourceRun.IsHyperLink &&
                                sourceRun.Hyperlink != null) {
                                IElement? sourceAnchor = CreateEquationHyperlinkNode(htmlDoc, sourceRun.Hyperlink);
                                if (sourceAnchor != null) {
                                    sourceAnchor.AppendChild(mathNode);
                                    expandedNodes.Add(sourceAnchor);
                                    continue;
                                }
                            }
                            if (mathNode != null) expandedNodes.Add(mathNode);
                            continue;
                        }

                        if (HtmlSemanticMetadata.IsTimeDateTimeMetadataRun(sourceRun)) {
                            continue;
                        }
                        if (segment.IsRunArtifact) {
                            AppendRunArtifacts(sourceRun, expandedNodes, segment.ArtifactElement);
                            continue;
                        }
                        if (string.IsNullOrEmpty(segment.Text)) continue;
                        expandedNodes.Add(CreateEquationAdjacentTextNode(
                            htmlDoc,
                            sourceRun,
                            segment.Text!,
                            options,
                            document.Settings.Language,
                            runStyles,
                            includeHyperlink: hyperlinkNode == null));
                    }

                    if (hyperlinkNode == null) return expandedNodes;
                    foreach (INode expandedNode in expandedNodes) {
                        hyperlinkNode.AppendChild(expandedNode);
                    }
                    return new List<INode> { hyperlinkNode };
                }

                INode? CreatePositionedEquationNode(WordEquationOccurrence occurrence) {
                    IElement? mathNode = CreateEquationNode(htmlDoc, parent, occurrence.Equation, options);
                    if (mathNode == null) return null;

                    if (occurrence.StartChildIndex >= 0 &&
                        occurrence.StartChildIndex < paragraphChildren.Count &&
                        paragraphChildren[occurrence.StartChildIndex] is Hyperlink hyperlink) {
                        IElement? anchor = CreateEquationHyperlinkNode(
                            htmlDoc,
                            new WordHyperLink(para._document, para._paragraph, hyperlink));
                        if (anchor != null) {
                            anchor.AppendChild(mathNode);
                            return anchor;
                        }
                    }

                    return mathNode;
                }

                void AppendEquationNodesBefore(int childIndex) {
                    while (nextEquation < equations.Count && equations[nextEquation].StartChildIndex < childIndex) {
                        int equationChildIndex = equations[nextEquation].StartChildIndex;
                        if (equationChildIndex >= 0 &&
                            equationChildIndex < paragraphChildren.Count &&
                            paragraphChildren[equationChildIndex] is DocumentFormat.OpenXml.OpenXmlElement container &&
                            WordEquation.IsVisibleEquationContentContainer(container)) {
                            List<WordEquationOccurrence> coveringEquations = equations
                                .Where(equation => equation.ContainsChildIndex(equationChildIndex))
                                .ToList();
                            if (expandedEquationContainers.Add(container)) {
                                foreach (INode expandedNode in CreateExpandedEquationContainerNodes(container, coveringEquations, para)) {
                                    AppendNode(expandedNode);
                                }
                            }
                            while (nextEquation < equations.Count && equations[nextEquation].StartChildIndex == equationChildIndex) {
                                nextEquation++;
                            }
                            continue;
                        }

                        INode? mathNode = CreatePositionedEquationNode(equations[nextEquation++]);
                        if (mathNode != null) AppendNode(mathNode);
                    }
                }

                for (int i = 0; i < runs.Count; i++) {
                    var run = runs[i];
                    DocumentFormat.OpenXml.OpenXmlElement? runContentContainer = run._hyperlink
                        ?? (DocumentFormat.OpenXml.OpenXmlElement?)run._stdRun
                        ?? run._run;
                    DocumentFormat.OpenXml.OpenXmlElement? runContainer =
                        WordEquation.GetDirectParagraphChild(run._paragraph, runContentContainer);
                    int runIndex = runContainer == null ? int.MaxValue : paragraphChildren.IndexOf(runContainer);
                    AppendEquationNodesBefore(runIndex < 0 ? int.MaxValue : runIndex);
                    List<WordEquationOccurrence> coveringEquations = equations
                        .Where(equation => equation.ContainsChildIndex(runIndex))
                        .ToList();
                    if (runContainer != null &&
                        coveringEquations.Any(equation => equation.StartChildIndex == runIndex) &&
                        expandedEquationContainers.Add(runContainer)) {
                        foreach (INode expandedNode in CreateExpandedEquationContainerNodes(runContainer, coveringEquations, run)) {
                            AppendNode(expandedNode);
                        }
                        while (nextEquation < equations.Count && equations[nextEquation].StartChildIndex == runIndex) {
                            nextEquation++;
                        }
                        continue;
                    }
                    if (coveringEquations.Count > 0) {
                        continue;
                    }
                    if (HtmlSemanticMetadata.IsTimeDateTimeMetadataRun(run)) {
                        continue;
                    }

                    if (AppendRunArtifacts(run, nodes)) continue;
                    if (string.IsNullOrEmpty(run.Text)) {
                        continue;
                    }

                    bool isHtmlDeletedText = string.Equals(run.CharacterStyleId, HtmlSemanticStyleIds.DeletedText, StringComparison.OrdinalIgnoreCase);
                    bool isHtmlInsertedText = string.Equals(run.CharacterStyleId, HtmlSemanticStyleIds.InsertedText, StringComparison.OrdinalIgnoreCase);
                    bool isHtmlMarkedText = string.Equals(run.CharacterStyleId, HtmlSemanticStyleIds.MarkedText, StringComparison.OrdinalIgnoreCase);

                    if (string.Equals(run.CharacterStyleId, "HtmlQuote", StringComparison.OrdinalIgnoreCase)) {
                        if (!inQuote) {
                            quote = CreateOutputElement(htmlDoc, "q");
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
                        var strong = CreateOutputElement(htmlDoc, "strong");
                        strong.AppendChild(node);
                        node = strong;
                    }

                    if (run.Italic) {
                        var em = CreateOutputElement(htmlDoc, "em");
                        em.AppendChild(node);
                        node = em;
                    }

                    if ((run.Strike || run.DoubleStrike) && !isHtmlDeletedText) {
                        var s = CreateOutputElement(htmlDoc, "s");
                        s.AppendChild(node);
                        node = s;
                    }

                    if (run.Underline != null && !isHtmlInsertedText) {
                        var u = CreateOutputElement(htmlDoc, "u");
                        u.AppendChild(node);
                        node = u;
                    }

                    if (run.VerticalTextAlignment == VerticalPositionValues.Superscript) {
                        var sup = CreateOutputElement(htmlDoc, "sup");
                        sup.AppendChild(node);
                        node = sup;
                    }

                    if (run.VerticalTextAlignment == VerticalPositionValues.Subscript) {
                        var sub = CreateOutputElement(htmlDoc, "sub");
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
                            var a = CreateOutputElement(htmlDoc, "a");
                            SetOutputAttribute(htmlDoc, a, "href", href!, "Hyperlink:href");
                            a.AppendChild(node);
                            node = a;
                        }
                        // if href is null/empty, fall back to plain text       
                    }

                    node = ApplyHtmlSemanticCharacterStyle(
                        htmlDoc,
                        run,
                        run.Text ?? string.Empty,
                        node,
                        options.IncludeRunHighlightStyles,
                        out bool handledHtmlStyle);

                    if (options.IncludeFontStyles) {
                        var font = run.FontFamily ?? options.FontFamily;
                        if (!string.IsNullOrEmpty(font)) {
                            var span = CreateOutputElement(htmlDoc, "span");
                            SetOutputAttribute(
                                htmlDoc,
                                span,
                                "style",
                                $"font-family:{QuoteCssString(font!)}",
                                "RunFontStyle");
                            span.AppendChild(node);
                            node = span;
                        }
                    }

                    if (run.FontSize != null) {
                        var span = CreateOutputElement(htmlDoc, "span");
                        span.SetAttribute("style", $"font-size:{run.FontSize.Value}pt");
                        span.AppendChild(node);
                        node = span;
                    }

                    // Caps / SmallCaps
                    if (run.CapsStyle == CapsStyle.SmallCaps) {
                        var span = CreateOutputElement(htmlDoc, "span");
                        span.SetAttribute("style", "font-variant:small-caps");
                        span.AppendChild(node);
                        node = span;
                    } else if (run.CapsStyle == CapsStyle.Caps) {
                        var span = CreateOutputElement(htmlDoc, "span");
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
                                string? normalized = NormalizeSixDigitHexColor(colorHex);
                                if (normalized != null) {
                                    inlineStyles.Add($"color:#{normalized}");
                                }
                            }
                        }
                        if (options.IncludeRunHighlightStyles && !isHtmlMarkedText) {
                            string? normalizedRunBackground = NormalizeSixDigitHexColor(
                                WordDocumentImageRenderer.ResolveRunShadingFillColorHex(run));
                            string? highlightCss = GetHighlightCss(
                                WordDocumentImageRenderer.ResolveRunHighlight(run));
                            if (!string.IsNullOrEmpty(highlightCss) &&
                                (!isHtmlMarkedText || normalizedRunBackground != null)) {
                                inlineStyles.Add($"background-color:{highlightCss}");
                            } else if (normalizedRunBackground != null) {
                                inlineStyles.Add($"background-color:#{normalizedRunBackground}");
                            }
                        }
                        if (inlineStyles.Count > 0) {
                            var span = CreateOutputElement(htmlDoc, "span");
                            span.SetAttribute("style", string.Join(";", inlineStyles));
                            span.AppendChild(node);
                            node = span;
                        }
                    }

                    if (options.IncludeRunClasses && !string.IsNullOrEmpty(run.CharacterStyleId) && !handledHtmlStyle) {
                        var spanClass = CreateOutputElement(htmlDoc, "span");
                        spanClass.SetAttribute("class", GetSafeStyleClassName(run.CharacterStyleId));
                        spanClass.AppendChild(node);
                        node = spanClass;
                        runStyles.Add(run.CharacterStyleId!);
                    }

                    var runLanguage = NormalizeRunLanguage(run.Language, document.Settings.Language);
                    if (!string.IsNullOrEmpty(runLanguage)) {
                        var spanLanguage = CreateOutputElement(htmlDoc, "span");
                        spanLanguage.SetAttribute("lang", runLanguage);
                        spanLanguage.AppendChild(node);
                        node = spanLanguage;
                    }

                    if (inQuote && quote != null) {
                        quote.AppendChild(node);
                    } else {
                        nodes.Add(node);
                    }
                }

                AppendEquationNodesBefore(int.MaxValue);
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
                var runs = para.GetFormattedRuns().ToList();
                return runs.Count > 0 && runs.All(r => r.Code);
            }

            bool TryAppendPlainParagraph(IElement parent, WordParagraph para) {
                if (options.IncludeFontStyles && !string.IsNullOrEmpty(options.FontFamily)) {
                    return false;
                }

                Paragraph paragraph = para._paragraph;
                if (paragraph.ParagraphProperties?.HasChildren == true) {
                    return false;
                }

                StringBuilder? text = null;
                foreach (var child in paragraph.ChildElements) {
                    if (child is ParagraphProperties) {
                        continue;
                    }
                    if (child is not Run run || run.RunProperties != null) {
                        return false;
                    }

                    foreach (var runChild in run.ChildElements) {
                        if (runChild is not Text wordText) {
                            return false;
                        }

                        text ??= new StringBuilder();
                        text.Append(wordText.Text);
                    }
                }

                var element = CreateOutputElement(htmlDoc, "p");
                if (text != null) {
                    element.TextContent = text.ToString();
                }
                parent.AppendChild(element);
                return true;
            }


            void AppendParagraph(IElement parent, WordParagraph para, bool suppressStructuralBookmark = false) {
                if (!suppressStructuralBookmark && para.IsBookmark && para.Bookmark != null) {
                    var name = para.Bookmark.Name ?? string.Empty;
                    var parts = name.Split(new[] { ':' }, 2);
                    if (parts.Length == 2 && IsStructuralTag(parts[0])) {
                        var structEl = CreateOutputElement(htmlDoc, parts[0]);
                        structEl.SetAttribute("id", parts[1]);
                        AppendParagraph(structEl, para, suppressStructuralBookmark: true);
                        parent.AppendChild(structEl);
                        return;
                    }
                }

                if (TryAppendPlainParagraph(parent, para)) {
                    return;
                }

                if (para.Borders.BottomStyle != null && string.IsNullOrWhiteSpace(para.Text)) {
                    var hr = CreateOutputElement(htmlDoc, "hr");
                    ApplyBookmarkId(hr, para);
                    parent.AppendChild(hr);
                    return;
                }

                if (IsCodeParagraph(para)) {
                    var pre = CreateOutputElement(htmlDoc, "pre");
                    ApplyBookmarkId(pre, para);
                    var code = CreateOutputElement(htmlDoc, "code");
                    code.TextContent = para.Text ?? string.Empty;
                    pre.AppendChild(code);
                    parent.AppendChild(pre);
                    return;
                }

                int level = para.Style.HasValue ? HeadingStyleMapper.GetLevelForHeadingStyle(para.Style.Value) : 0;
                bool isBlockQuote = (!string.IsNullOrEmpty(para.StyleId) && (string.Equals(para.StyleId, "Quote", StringComparison.OrdinalIgnoreCase) || string.Equals(para.StyleId, "IntenseQuote", StringComparison.OrdinalIgnoreCase)))
                    || (para.IndentationBefore.HasValue && para.IndentationBefore.Value > 0);
                var element = CreateOutputElement(htmlDoc, isBlockQuote ? "blockquote" : (level > 0 ? $"h{level}" : "p"));
                if (isBlockQuote && TryGetBlockquoteCiteAttribute(para, out var blockquoteCite)) {
                    element.SetAttribute("cite", blockquoteCite);
                }
                ApplyBookmarkId(element, para);
                if (options.IncludeParagraphClasses && !string.IsNullOrEmpty(para.StyleId)) {
                    element.SetAttribute("class", GetSafeStyleClassName(para.StyleId));
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
                string? normalizedParagraphBackground = NormalizeSixDigitHexColor(pBg);
                if (normalizedParagraphBackground != null) {
                    pStyles.Add($"background-color:#{normalizedParagraphBackground}");
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

            void AppendDefinitionListItem(IElement definitionList, WordParagraph para) {
                var item = CreateOutputElement(htmlDoc, GetDefinitionListTagName(para));
                ApplyBookmarkId(item, para);
                if (para.BiDi) {
                    item.SetAttribute("dir", "rtl");
                }
                AppendRuns(item, para);
                definitionList.AppendChild(item);
            }

            bool IsCaptionParagraph(WordParagraph para) =>
                string.Equals(para.StyleId, "Caption", StringComparison.OrdinalIgnoreCase);

            void AppendTableCaption(IElement tableElement, WordParagraph captionParagraph) {
                var caption = CreateOutputElement(htmlDoc, "caption");
                ApplyBookmarkId(caption, captionParagraph);
                if (captionParagraph.BiDi) {
                    caption.SetAttribute("dir", "rtl");
                }
                if (options.IncludeParagraphClasses && !string.IsNullOrEmpty(captionParagraph.StyleId)) {
                    caption.SetAttribute("class", GetSafeStyleClassName(captionParagraph.StyleId));
                    paragraphStyles.Add(captionParagraph.StyleId!);
                }
                AppendRuns(caption, captionParagraph);
                tableElement.AppendChild(caption);
            }

            void AppendTable(IElement parent, WordTable table, WordParagraph? captionParagraph = null, int nestingDepth = 0) {
                if (options.MaxTableNestingDepth <= 0) {
                    throw new ArgumentOutOfRangeException(nameof(options.MaxTableNestingDepth));
                }
                if (nestingDepth >= options.MaxTableNestingDepth) {
                    throw new InvalidDataException($"The Word table nesting exceeds the {options.MaxTableNestingDepth}-level HTML conversion limit.");
                }
                var tableEl = CreateOutputElement(htmlDoc, "table");
                var tableStyles = new List<string>();
                var tableWidth = GetWidthCss(table.WidthType, table.Width);
                if (!string.IsNullOrEmpty(tableWidth)) {
                    tableStyles.Add($"width:{tableWidth}");
                }
                var tableCellSpacing = GetTableCellSpacingCss(table);
                if (!string.IsNullOrEmpty(tableCellSpacing)) {
                    tableStyles.Add($"border-spacing:{tableCellSpacing}");
                }
                if (TableHasBorder(table)) {
                    tableStyles.Add("border:1px solid black");
                    tableStyles.Add(!string.IsNullOrEmpty(tableCellSpacing) ? "border-collapse:separate" : "border-collapse:collapse");
                }
                if (tableStyles.Count > 0) {
                    tableEl.SetAttribute("style", string.Join(";", tableStyles));
                }
                if (captionParagraph != null) {
                    AppendTableCaption(tableEl, captionParagraph);
                }
                if (options.IncludeTableColumnGroups) {
                    AppendColumnGroup(htmlDoc, tableEl, table);
                }

                int headerRowCount = 0;
                while (headerRowCount < table.Rows.Count && table.Rows[headerRowCount].RepeatHeaderRowAtTheTopOfEachPage) {
                    headerRowCount++;
                }
                bool hasFooterRow = table.ConditionalFormattingLastRow == true && table.Rows.Count > headerRowCount;
                IElement? thead = null;
                IElement? tbody = null;
                IElement? tfoot = null;

                for (int r = 0; r < table.Rows.Count; r++) {
                    var row = table.Rows[r];
                    var tr = CreateOutputElement(htmlDoc, "tr");
                    bool isHeaderRow = headerRowCount > 0 && r < headerRowCount;
                    bool isFooterRow = hasFooterRow && r == table.Rows.Count - 1;
                    for (int c = 0; c < row.Cells.Count; c++) {
                        var cell = row.Cells[c];
                        if (cell.HorizontalMerge == MergedCellValues.Continue || cell.VerticalMerge == MergedCellValues.Continue) {
                            continue;
                        }
                        var cellElement = CreateOutputElement(htmlDoc, isHeaderRow ? "th" : "td");
                        if (isHeaderRow) {
                            cellElement.SetAttribute("scope", "col");
                        }
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
                            string vAlign = "top";
                            if (cell.VerticalAlignment.Value == TableVerticalAlignmentValues.Center) {
                                vAlign = "middle";
                            } else if (cell.VerticalAlignment.Value == TableVerticalAlignmentValues.Bottom) {
                                vAlign = "bottom";
                            }
                            cellStyles.Add($"vertical-align:{vAlign}");
                        }
                        var bg = cell.ShadingFillColorHex;
                        string? normalizedCellBackground = NormalizeSixDigitHexColor(bg);
                        if (normalizedCellBackground != null) {
                            cellStyles.Add($"background-color:#{normalizedCellBackground}");
                        }
                        var borderCss = GetBorderCss(cell);
                        if (borderCss.Count > 0) {
                            cellStyles.AddRange(borderCss);
                        }
                        if (cellStyles.Count > 0) {
                            cellElement.SetAttribute("style", string.Join(";", cellStyles));
                        }

                        IElement? cellDefinitionList = null;
                        var cellParagraphs = cell.Paragraphs;
                        var processedCellParagraphs = new HashSet<WordParagraph>(ParagraphElementComparer.Instance);
                        var cellListStack = new Stack<IElement>();
                        var cellItemStack = new Stack<IElement>();
                        var cellListNumberStack = new Stack<int>();
                        for (int pIdx = 0; pIdx < cellParagraphs.Count; pIdx++) {
                            var p = cellParagraphs[pIdx];
                            if (processedCellParagraphs.Contains(p)) {
                                continue;
                            }
                            for (int j = pIdx + 1; j < cellParagraphs.Count && SameParagraphElement(cellParagraphs[j], p); j++) {
                                var candidate = cellParagraphs[j];
                                if ((!p.IsBookmark && candidate.IsBookmark) || candidate.Text.Length > p.Text.Length) {
                                    p = candidate;
                                }
                            }
                            if (IsDefinitionListParagraph(p) && IsEmptyDefinitionListParagraph(p)) {
                                for (int j = pIdx + 1; j < cellParagraphs.Count; j++) {
                                    if (!SameParagraphElement(cellParagraphs[j], p)) {
                                        break;
                                    }
                                    if (!IsEmptyDefinitionListParagraph(cellParagraphs[j])) {
                                        p = cellParagraphs[j];
                                        break;
                                    }
                                }
                            }
                            processedCellParagraphs.Add(p);
                            var cellListInfo = DocumentTraversal.GetListInfo(p);
                            if (cellListInfo != null) {
                                cellDefinitionList = null;
                                AppendListParagraph(cellElement, p, cellListInfo.Value, cellListStack, cellItemStack, cellListNumberStack);
                            } else if (IsCodeParagraph(p)) {
                                ClearListStacks(cellListStack, cellItemStack, cellListNumberStack);
                                cellDefinitionList = null;
                                List<string> lines = new();
                                lines.Add(p.Text);
                                while (pIdx + 1 < cellParagraphs.Count && IsCodeParagraph(cellParagraphs[pIdx + 1])) {
                                    lines.Add(cellParagraphs[pIdx + 1].Text);
                                    pIdx++;
                                }
                                var pre = CreateOutputElement(htmlDoc, "pre");
                                var code = CreateOutputElement(htmlDoc, "code");
                                code.TextContent = string.Join("\n", lines);
                                pre.AppendChild(code);
                                cellElement.AppendChild(pre);
                            } else if (IsDefinitionListParagraph(p)) {
                                ClearListStacks(cellListStack, cellItemStack, cellListNumberStack);
                                if (IsEmptyDefinitionListParagraph(p)) {
                                    continue;
                                }
                                if (cellDefinitionList == null) {
                                    cellDefinitionList = CreateOutputElement(htmlDoc, "dl");
                                    cellElement.AppendChild(cellDefinitionList);
                                }
                                AppendDefinitionListItem(cellDefinitionList, p);
                            } else {
                                ClearListStacks(cellListStack, cellItemStack, cellListNumberStack);
                                cellDefinitionList = null;
                                AppendParagraph(cellElement, p);
                            }
                        }

                        if (cell.HasNestedTables) {
                            foreach (var nested in cell.DirectNestedTables) {
                                cancellationToken.ThrowIfCancellationRequested();
                                AppendTable(cellElement, nested, nestingDepth: nestingDepth + 1);
                            }
                        }

                        tr.AppendChild(cellElement);
                    }
                    if (headerRowCount == 0 && !hasFooterRow) {
                        tableEl.AppendChild(tr);
                    } else if (isHeaderRow) {
                        if (thead == null) {
                            thead = CreateOutputElement(htmlDoc, "thead");
                            tableEl.AppendChild(thead);
                        }
                        thead.AppendChild(tr);
                    } else if (isFooterRow) {
                        if (tfoot == null) {
                            tfoot = CreateOutputElement(htmlDoc, "tfoot");
                            tableEl.AppendChild(tfoot);
                        }
                        tfoot.AppendChild(tr);
                    } else {
                        if (tbody == null) {
                            tbody = CreateOutputElement(htmlDoc, "tbody");
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
                        "*" => "'*'",
                        "+" => "'+'",
                        "•" or "·" or "●" or "∙" or "" or null or "" => "disc",
                        _ => QuoteCssListMarker(info.LevelText),
                    };
                }
                if (format != null && formatMap.TryGetValue(format.Value, out var map)) {
                    return map.Css;
                }
                return null;
            }

            string QuoteCssListMarker(string marker) {
                var escaped = marker
                    .Replace("\\", "\\\\")
                    .Replace("'", "\\'")
                    .Replace("\r", "\\d ")
                    .Replace("\n", "\\a ")
                    .Replace("\t", "\\9 ");
                return $"'{escaped}'";
            }

            string? GetListType(DocumentTraversal.ListInfo info) {
                var format = info.NumberFormat;
                if (format == NumberFormatValues.Bullet) {
                    return info.LevelText switch {
                        "o" or "◦" => "circle",
                        "■" or "§" => "square",
                        "-" or "\u2013" or "\u2014" or "*" or "+" => null,
                        "•" or "·" or "●" or "∙" or "" or null or "" => "disc",
                        _ => null,
                    };
                }
                if (format != null && formatMap.TryGetValue(format.Value, out var map)) {
                    return map.Type;
                }
                return null;
            }

            var listIndices = DocumentTraversal.BuildListIndices(document);

            void ClearListStacks(Stack<IElement> lists, Stack<IElement> items, Stack<int> numberIds) {
                lists.Clear();
                items.Clear();
                numberIds.Clear();
            }

            void AppendListParagraph(
                IElement parent,
                WordParagraph paragraph,
                DocumentTraversal.ListInfo listInfo,
                Stack<IElement> lists,
                Stack<IElement> items,
                Stack<int> numberIds) {
                int level = listInfo.Level;
                if (options.MaxListNestingDepth <= 0) {
                    throw new ArgumentOutOfRangeException(nameof(options.MaxListNestingDepth));
                }
                if (level < 0 || level >= options.MaxListNestingDepth) {
                    throw new InvalidDataException($"The Word list level {level} exceeds the {options.MaxListNestingDepth}-level HTML conversion limit.");
                }
                int desiredListDepth = level + 1;

                while (lists.Count > desiredListDepth) {
                    lists.Pop();
                    numberIds.Pop();
                }
                while (items.Count > level) {
                    items.Pop();
                }

                bool ordered = listInfo.Ordered;
                string listTag = ordered ? "ol" : "ul";
                int numberId = paragraph._listNumberId.GetValueOrDefault();
                if (lists.Count == desiredListDepth
                    && (numberIds.Peek() != numberId
                        || !string.Equals(lists.Peek().TagName, listTag, StringComparison.OrdinalIgnoreCase))) {
                    lists.Pop();
                    numberIds.Pop();
                }

                while (lists.Count < desiredListDepth) {
                    var listEl = CreateOutputElement(htmlDoc, listTag);
                    if (ordered) {
                        if (listIndices.TryGetValue(paragraph, out var indexInfo)) {
                            listEl.SetAttribute("start", indexInfo.Index.ToString());
                        } else {
                            listEl.SetAttribute("start", listInfo.Start.ToString());
                        }
                    }
                    var typeAttr = GetListType(listInfo);
                    if (!string.IsNullOrEmpty(typeAttr)) {
                        listEl.SetAttribute("type", typeAttr);
                    }
                    var listStyle = GetListStyle(listInfo);
                    if (options.IncludeListStyles && !string.IsNullOrEmpty(listStyle)) {
                        listEl.SetAttribute("style", $"list-style-type:{listStyle}");
                    }
                    if (options.IncludeListDefinitions) {
                        ApplyListDefinition(listEl, listInfo, listStyle, listDefinitions);
                    }
                    if (items.Count > 0) {
                        items.Peek().AppendChild(listEl);
                    } else {
                        parent.AppendChild(listEl);
                    }
                    lists.Push(listEl);
                    numberIds.Push(numberId);
                }

                var li = CreateOutputElement(htmlDoc, "li");
                ApplyBookmarkId(li, paragraph);
                lists.Peek().AppendChild(li);
                items.Push(li);
                AppendRuns(li, paragraph);
            }

            var processedParagraphs = new HashSet<WordParagraph>(ParagraphElementComparer.Instance);
            int sectionIndex = 0;
            foreach (var section in DocumentTraversal.EnumerateSections(document)) {
                cancellationToken.ThrowIfCancellationRequested();
                IElement sectionParent = body;
                if (options.IncludeSectionMetadata) {
                    sectionParent = CreateSectionElement(htmlDoc, section, sectionIndex, sectionIndex == 0);
                    body.AppendChild(sectionParent);
                }
                AppendHeaderFooterRegions(htmlDoc, sectionParent, section, sectionIndex, true, (parent, paragraph) => AppendParagraph(parent, paragraph), (parent, table) => AppendTable(parent, table), options, cancellationToken);

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
                IElement? activeDefinitionList = null;
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
                                if (elements[j] is WordParagraph sibling && SameParagraphElement(sibling, paragraph)) {
                                    if (sibling.IsBookmark) { paragraph = sibling; }
                                    continue;
                                }
                                break;
                            }
                        }
                        processedParagraphs.Add(paragraph);
                        if (TryAppendPlainParagraph(sectionParent, paragraph)) {
                            CloseLists();
                            activeDefinitionList = null;
                            continue;
                        }
                        if (IsCaptionParagraph(paragraph) && idx + 1 < elements.Count && elements[idx + 1] is WordTable) {
                            activeDefinitionList = null;
                            continue;
                        }
                        var listInfo = DocumentTraversal.GetListInfo(paragraph);
                        if (listInfo != null) {
                            activeDefinitionList = null;
                            AppendListParagraph(sectionParent, paragraph, listInfo.Value, listStack, itemStack, listNumberStack);
                        } else {
                            CloseLists();
                            if (IsDefinitionListParagraph(paragraph)) {
                                if (IsEmptyDefinitionListParagraph(paragraph)) {
                                    continue;
                                }
                                if (activeDefinitionList == null) {
                                    activeDefinitionList = CreateOutputElement(htmlDoc, "dl");
                                    sectionParent.AppendChild(activeDefinitionList);
                                }
                                AppendDefinitionListItem(activeDefinitionList, paragraph);
                            } else if (paragraph.IsImage && idx + 1 < elements.Count && elements[idx + 1] is WordParagraph captionPara && string.Equals(captionPara.StyleId, "Caption", StringComparison.OrdinalIgnoreCase)) {
                                activeDefinitionList = null;
                                var figure = CreateOutputElement(htmlDoc, "figure");
                                ApplyBookmarkId(figure, paragraph);
                                AppendRuns(figure, paragraph);
                                var figCap = CreateOutputElement(htmlDoc, "figcaption");
                                if (options.IncludeParagraphClasses && !string.IsNullOrEmpty(captionPara.StyleId)) {
                                    figCap.SetAttribute("class", GetSafeStyleClassName(captionPara.StyleId));
                                    paragraphStyles.Add(captionPara.StyleId!);
                                }
                                AppendRuns(figCap, captionPara);
                                figure.AppendChild(figCap);
                                sectionParent.AppendChild(figure);
                                idx++;
                            } else if (IsCaptionParagraph(paragraph) && idx + 1 < elements.Count && elements[idx + 1] is WordParagraph imagePara && imagePara.IsImage) {
                                activeDefinitionList = null;
                                var figure = CreateOutputElement(htmlDoc, "figure");
                                ApplyBookmarkId(figure, imagePara);
                                var figCap = CreateOutputElement(htmlDoc, "figcaption");
                                if (options.IncludeParagraphClasses && !string.IsNullOrEmpty(paragraph.StyleId)) {
                                    figCap.SetAttribute("class", GetSafeStyleClassName(paragraph.StyleId));
                                    paragraphStyles.Add(paragraph.StyleId!);
                                }
                                AppendRuns(figCap, paragraph);
                                figure.AppendChild(figCap);
                                AppendRuns(figure, imagePara);
                                sectionParent.AppendChild(figure);
                                idx++;
                            } else if (IsCodeParagraph(paragraph)) {
                                activeDefinitionList = null;
                                List<string> lines = new();
                                lines.Add(paragraph.Text);
                                while (idx + 1 < elements.Count && elements[idx + 1] is WordParagraph nextPara && DocumentTraversal.GetListInfo(nextPara) == null && IsCodeParagraph(nextPara)) {
                                    lines.Add(nextPara.Text);
                                    idx++;
                                }
                                var pre = CreateOutputElement(htmlDoc, "pre");
                                ApplyBookmarkId(pre, paragraph);
                                var code = CreateOutputElement(htmlDoc, "code");
                                code.TextContent = string.Join("\n", lines);
                                pre.AppendChild(code);
                                sectionParent.AppendChild(pre);
                            } else {
                                activeDefinitionList = null;
                                AppendParagraph(sectionParent, paragraph);
                            }
                        }
                    } else if (element is WordTable table) {
                        CloseLists();
                        activeDefinitionList = null;
                        WordParagraph? captionParagraph = null;
                        if (idx > 0 && elements[idx - 1] is WordParagraph previousCaption && IsCaptionParagraph(previousCaption)) {
                            captionParagraph = previousCaption;
                        } else if (idx + 1 < elements.Count && elements[idx + 1] is WordParagraph nextCaption && IsCaptionParagraph(nextCaption)) {
                            captionParagraph = nextCaption;
                            processedParagraphs.Add(nextCaption);
                            idx++;
                        }
                        AppendTable(sectionParent, table, captionParagraph);
                    }
                }
                if (options.ExportHeadersAndFooters) {
                    CloseLists();
                    AppendHeaderFooterRegions(htmlDoc, sectionParent, section, sectionIndex, false, (parent, paragraph) => AppendParagraph(parent, paragraph), (parent, table) => AppendTable(parent, table), options, cancellationToken);
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

            using var outputWriter = new BoundedHtmlWriter(
                options.MaxOutputCharacters,
                actual => ThrowExportLimitExceeded(options, "WordHtmlOutputLimitExceeded", "Generated HTML exceeds the configured output-character limit.", "MaxOutputCharacters", actual, options.MaxOutputCharacters));
            htmlDoc.DocumentElement.ToHtml(outputWriter, HtmlMarkupFormatter.Instance);
            OutputConstructionBudgets.Remove(htmlDoc);
            return outputWriter.ToString();
        }

        private static bool SameParagraphElement(WordParagraph left, WordParagraph right) =>
            ReferenceEquals(left._paragraph, right._paragraph);

        private sealed class ParagraphElementComparer : IEqualityComparer<WordParagraph> {
            internal static readonly ParagraphElementComparer Instance = new();

            public bool Equals(WordParagraph? left, WordParagraph? right) =>
                ReferenceEquals(left, right) ||
                (left != null && right != null && SameParagraphElement(left, right));

            public int GetHashCode(WordParagraph paragraph) {
                object identity = paragraph._paragraph != null ? paragraph._paragraph : paragraph;
                return System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(identity);
            }
        }

        private static string? NormalizeRunLanguage(string? language, string? documentLanguage) {
            var normalized = language?.Trim();
            if (string.IsNullOrEmpty(normalized)) {
                return null;
            }

            var normalizedDocumentLanguage = documentLanguage?.Trim();
            if (!string.IsNullOrEmpty(normalizedDocumentLanguage) &&
                string.Equals(normalized, normalizedDocumentLanguage, StringComparison.OrdinalIgnoreCase)) {
                return null;
            }

            return normalized;
        }
    }
}
