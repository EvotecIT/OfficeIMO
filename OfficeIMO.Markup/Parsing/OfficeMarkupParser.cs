using OfficeIMO.Markdown;

namespace OfficeIMO.Markup;

/// <summary>
/// Parses Markdown plus explicit OfficeIMO semantic fenced blocks into a profile-aware AST.
/// </summary>
public static class OfficeMarkupParser {
    private static readonly string[] OfficeLanguages = {
        "officeimo",
        "officeimo-presentation",
        "officeimo-document",
        "officeimo-workbook",
        "officeimo-slide",
        "officeimo-page-break",
        "officeimo-pagebreak",
        "officeimo-section",
        "officeimo-header",
        "officeimo-footer",
        "officeimo-toc",
        "officeimo-sheet",
        "officeimo-range",
        "officeimo-formula",
        "officeimo-table",
        "officeimo-chart",
        "officeimo-format"
    };

    public static OfficeMarkupParseResult Parse(string markup, OfficeMarkupParserOptions? options = null) {
        options ??= new OfficeMarkupParserOptions();
        var source = (markup ?? string.Empty).Replace("\r\n", "\n").Replace('\r', '\n');
        var diagnostics = new List<OfficeMarkupDiagnostic>();
        var metadata = ExtractFrontMatter(ref source);
        var profile = ResolveProfile(options.Profile, metadata);
        var document = new OfficeMarkupDocument(profile);
        CopyAttributes(metadata, document.Metadata);

        if (!TryMapOfficeSyntax(source, document, profile, diagnostics)) {
            var markdownOptions = CreateMarkdownOptions(options);
            var markdownDocument = MarkdownReader.Parse(source, markdownOptions);
            MapMarkdownBlocks(markdownDocument.Blocks, document.Blocks, profile, diagnostics);
        }

        if (options.Validate) {
            diagnostics.AddRange(OfficeMarkupValidator.Validate(document));
        }

        return new OfficeMarkupParseResult(document, diagnostics);
    }

    private static Dictionary<string, string> ExtractFrontMatter(ref string markup) {
        var metadata = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var lines = markup.Split('\n');
        if (lines.Length < 3 || !string.Equals(lines[0].Trim(), "---", StringComparison.Ordinal)) {
            return metadata;
        }

        var end = -1;
        for (int i = 1; i < lines.Length; i++) {
            if (string.Equals(lines[i].Trim(), "---", StringComparison.Ordinal)) {
                end = i;
                break;
            }
        }

        if (end < 0) {
            return metadata;
        }

        for (int i = 1; i < end; i++) {
            TryParseAttributeLine(lines[i], metadata);
        }

        markup = string.Join("\n", lines.Skip(end + 1));
        return metadata;
    }

    private static OfficeMarkupProfile ResolveProfile(OfficeMarkupProfile defaultProfile, IDictionary<string, string> metadata) {
        if (metadata.TryGetValue("profile", out var profile) && Enum.TryParse<OfficeMarkupProfile>(profile, true, out var parsed)) {
            return parsed;
        }

        return defaultProfile;
    }

    private static bool TryMapOfficeSyntax(
        string markup,
        OfficeMarkupDocument document,
        OfficeMarkupProfile profile,
        IList<OfficeMarkupDiagnostic> diagnostics) {
        if (!ContainsOfficeSyntax(markup)) {
            return false;
        }

        if (profile == OfficeMarkupProfile.Presentation && (ContainsAtDirective(markup, "slide") || HasSlideSeparators(markup))) {
            MapPresentationSyntax(markup, document.Blocks, profile, diagnostics);
            return true;
        }

        MapOfficeAwareText(markup, document.Blocks, profile, diagnostics, null);
        return true;
    }

    private static bool ContainsOfficeSyntax(string markup) =>
        markup.IndexOf("\n@", StringComparison.Ordinal) >= 0
        || markup.StartsWith("@", StringComparison.Ordinal)
        || markup.IndexOf("\n::", StringComparison.Ordinal) >= 0
        || markup.StartsWith("::", StringComparison.Ordinal)
        || HasSlideSeparators(markup);

    private static bool ContainsAtDirective(string markup, string name) {
        foreach (var line in markup.Split('\n')) {
            var trimmed = line.TrimStart();
            if (trimmed.StartsWith("@" + name, StringComparison.OrdinalIgnoreCase)) {
                return true;
            }
        }

        return false;
    }

    private static bool HasSlideSeparators(string markup) {
        var inFence = false;
        char fenceMarker = default;
        var fenceLength = 0;
        foreach (var line in markup.Split('\n')) {
            var trimmed = line.TrimStart();
            if (TryGetFenceInfo(trimmed, out var currentFenceMarker, out var currentFenceLength)) {
                if (inFence && currentFenceMarker == fenceMarker && currentFenceLength >= fenceLength) {
                    inFence = false;
                    fenceMarker = default;
                    fenceLength = 0;
                } else if (!inFence) {
                    inFence = true;
                    fenceMarker = currentFenceMarker;
                    fenceLength = currentFenceLength;
                }

                continue;
            }

            if (!inFence && string.Equals(trimmed, "---", StringComparison.Ordinal)) {
                return true;
            }
        }

        return false;
    }

    private static void MapPresentationSyntax(
        string markup,
        IList<OfficeMarkupBlock> target,
        OfficeMarkupProfile profile,
        IList<OfficeMarkupDiagnostic> diagnostics) {
        foreach (var segment in SplitSlideSegments(markup)) {
            if (string.IsNullOrWhiteSpace(segment)) {
                continue;
            }

            var slideSource = segment;
            var attributes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            ExtractAtDirective(ref slideSource, "slide", attributes);

            var slide = new OfficeMarkupSlideBlock(GetAttribute(attributes, "title"));
            ApplySlideAttributes(slide, attributes);
            MapOfficeAwareText(slideSource, slide.Blocks, profile, diagnostics, slide);
            PromoteLeadingHeadingToSlideTitle(slide);
            CopyAttributes(attributes, slide.Attributes);
            target.Add(slide);
        }
    }

    private static IEnumerable<string> SplitSlideSegments(string markup) {
        var builder = new StringBuilder();
        var inFence = false;
        char fenceMarker = default;
        var fenceLength = 0;
        foreach (var line in markup.Split('\n')) {
            var trimmed = line.TrimStart();
            if (TryGetFenceInfo(trimmed, out var currentFenceMarker, out var currentFenceLength)) {
                if (inFence && currentFenceMarker == fenceMarker && currentFenceLength >= fenceLength) {
                    inFence = false;
                    fenceMarker = default;
                    fenceLength = 0;
                } else if (!inFence) {
                    inFence = true;
                    fenceMarker = currentFenceMarker;
                    fenceLength = currentFenceLength;
                }

                builder.AppendLine(line);
                continue;
            }

            if (!inFence && string.Equals(trimmed, "---", StringComparison.Ordinal)) {
                yield return builder.ToString();
                builder.Clear();
                continue;
            }

            builder.AppendLine(line);
        }

        yield return builder.ToString();
    }

    private static void PromoteLeadingHeadingToSlideTitle(OfficeMarkupSlideBlock slide) {
        if (slide.Blocks.Count == 0 || slide.Blocks[0] is not OfficeMarkupHeadingBlock heading || heading.Level != 1) {
            return;
        }

        if (string.IsNullOrWhiteSpace(slide.Title)) {
            slide.Title = heading.Text;
        }

        slide.Blocks.RemoveAt(0);
    }

    private static void ApplySlideAttributes(OfficeMarkupSlideBlock slide, IDictionary<string, string> attributes) {
        slide.Layout = GetAttribute(attributes, "layout");
        slide.Section = GetAttribute(attributes, "section");
        slide.Transition = GetAttribute(attributes, "transition");
        slide.Background = GetAttribute(attributes, "background");
        slide.Notes = GetAttribute(attributes, "notes");
        slide.Placement = GetAttribute(attributes, "placement");
        if (TryGetInt32(attributes, "columns", out var columns)) {
            slide.Columns = columns;
        }
    }

    private static void MapOfficeAwareText(
        string markup,
        IList<OfficeMarkupBlock> target,
        OfficeMarkupProfile profile,
        IList<OfficeMarkupDiagnostic> diagnostics,
        OfficeMarkupSlideBlock? slideContext) {
        var lines = markup.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
        var markdownLines = new List<string>();
        var inFence = false;
        char fenceMarker = default;
        var fenceLength = 0;

        void FlushMarkdown() {
            var markdown = string.Join("\n", markdownLines).Trim('\n');
            markdownLines.Clear();
            if (string.IsNullOrWhiteSpace(markdown)) {
                return;
            }

            var nested = MarkdownReader.Parse(markdown, CreateNestedMarkdownOptions());
            MapMarkdownBlocks(nested.Blocks, target, profile, diagnostics);
        }

        for (int i = 0; i < lines.Length; i++) {
            var trimmed = lines[i].TrimStart();
            if (TryGetFenceInfo(trimmed, out var currentFenceMarker, out var currentFenceLength)) {
                markdownLines.Add(lines[i]);
                if (inFence) {
                    if (currentFenceMarker == fenceMarker && currentFenceLength >= fenceLength) {
                        inFence = false;
                        fenceMarker = default;
                        fenceLength = 0;
                    }
                } else {
                    inFence = true;
                    fenceMarker = currentFenceMarker;
                    fenceLength = currentFenceLength;
                }

                continue;
            }

            if (inFence) {
                markdownLines.Add(lines[i]);
                continue;
            }

            if (trimmed.StartsWith("::", StringComparison.Ordinal)) {
                FlushMarkdown();
                var directive = ReadColonDirective(lines, ref i);
                var block = CreateBlockFromDirective(directive.Command, directive.Attributes, directive.Body, directive.SourceText, diagnostics);
                if (block is OfficeMarkupExtensionBlock extension
                    && string.Equals(extension.Command, "notes", StringComparison.OrdinalIgnoreCase)
                    && slideContext != null) {
                    slideContext.Notes = AppendBlockText(slideContext.Notes, extension.Body);
                } else if (block != null) {
                    target.Add(block);
                }

                continue;
            }

            if (trimmed.StartsWith("@", StringComparison.Ordinal)) {
                FlushMarkdown();
                var directive = ReadAtDirective(lines, ref i);
                if (!string.Equals(directive.Command, "slide", StringComparison.OrdinalIgnoreCase)) {
                    var block = CreateBlockFromDirective(directive.Command, directive.Attributes, directive.Body, directive.SourceText, diagnostics);
                    if (block != null) {
                        target.Add(block);
                    }
                }

                continue;
            }

            markdownLines.Add(lines[i]);
        }

        FlushMarkdown();
    }

    private static bool TryGetFenceInfo(string line, out char marker, out int length) {
        marker = default;
        length = 0;
        if (string.IsNullOrWhiteSpace(line)) {
            return false;
        }

        char candidate = line[0];
        if (candidate != '`' && candidate != '~') {
            return false;
        }

        var index = 0;
        while (index < line.Length && line[index] == candidate) {
            index++;
        }

        if (index < 3) {
            return false;
        }

        marker = candidate;
        length = index;
        return true;
    }

    private static string AppendBlockText(string? current, string value) {
        if (string.IsNullOrWhiteSpace(current)) {
            return value.Trim();
        }

        return current!.TrimEnd() + "\n\n" + value.Trim();
    }

    private static bool ExtractAtDirective(ref string markup, string command, IDictionary<string, string> attributes) {
        var lines = markup.Split('\n').ToList();
        var inFence = false;
        char fenceMarker = default;
        var fenceLength = 0;

        for (int i = 0; i < lines.Count; i++) {
            var trimmed = lines[i].TrimStart();
            if (TryGetFenceInfo(trimmed, out var currentFenceMarker, out var currentFenceLength)) {
                if (inFence && currentFenceMarker == fenceMarker && currentFenceLength >= fenceLength) {
                    inFence = false;
                    fenceMarker = default;
                    fenceLength = 0;
                } else if (!inFence) {
                    inFence = true;
                    fenceMarker = currentFenceMarker;
                    fenceLength = currentFenceLength;
                }

                continue;
            }

            if (inFence || !trimmed.StartsWith("@" + command, StringComparison.OrdinalIgnoreCase)) {
                continue;
            }

            var index = i;
            var directive = ReadAtDirective(lines.ToArray(), ref index);
            CopyAttributes(directive.Attributes, attributes);
            lines.RemoveRange(i, index - i + 1);
            markup = string.Join("\n", lines);
            return true;
        }

        return false;
    }

    private static OfficeMarkupBlock? CreateBlockFromDirective(
        string command,
        IDictionary<string, string> attributes,
        string body,
        string sourceText,
        IList<OfficeMarkupDiagnostic> diagnostics) {
        var normalized = NormalizeCommand(command);
        OfficeMarkupBlock? block;
        switch (normalized) {
            case "notes":
                block = new OfficeMarkupExtensionBlock("notes", attributes, body);
                break;
            case "mermaid":
                block = new OfficeMarkupDiagramBlock("mermaid", body);
                break;
            case "image":
                block = new OfficeMarkupImageBlock(
                    GetAttribute(attributes, "src") ?? GetAttribute(attributes, "source") ?? body.Trim(),
                    GetAttribute(attributes, "alt"),
                    GetAttribute(attributes, "title"));
                break;
            case "chart":
                var chart = new OfficeMarkupChartBlock(GetAttribute(attributes, "type") ?? GetAttribute(attributes, "chartType") ?? "column") {
                    Title = GetAttribute(attributes, "title"),
                    Source = GetAttribute(attributes, "source") ?? GetAttribute(attributes, "range")
                };
                foreach (var row in ParseDelimitedRows(body)) {
                    chart.Data.Add(row);
                }

                block = chart;
                break;
            case "pagebreak":
            case "page-break":
                block = new OfficeMarkupPageBreakBlock();
                break;
            case "section":
                block = new OfficeMarkupSectionBlock(GetAttribute(attributes, "name") ?? GetAttribute(attributes, "title")) {
                    PageSize = GetAttribute(attributes, "pageSize") ?? GetAttribute(attributes, "size"),
                    Orientation = GetAttribute(attributes, "orientation")
                };
                break;
            case "header":
            case "footer":
                block = new OfficeMarkupHeaderFooterBlock(normalized, GetAttribute(attributes, "text") ?? body.Trim());
                break;
            case "toc":
            case "tableofcontents":
            case "table-of-contents":
                var toc = new OfficeMarkupTableOfContentsBlock {
                    Title = GetAttribute(attributes, "title")
                };
                if (TryGetInt32(attributes, "min", out var min) || TryGetInt32(attributes, "minLevel", out min)) {
                    toc.MinLevel = min;
                }

                if (TryGetInt32(attributes, "max", out var max) || TryGetInt32(attributes, "maxLevel", out max)) {
                    toc.MaxLevel = max;
                }

                block = toc;
                break;
            case "sheet":
                block = new OfficeMarkupSheetBlock(GetAttribute(attributes, "name") ?? body.Trim());
                break;
            case "range":
                var range = new OfficeMarkupRangeBlock(GetAttribute(attributes, "address") ?? GetAttribute(attributes, "range") ?? string.Empty) {
                    Sheet = GetAttribute(attributes, "sheet")
                };
                foreach (var row in ParseDelimitedRows(body)) {
                    range.Values.Add(row);
                }

                block = range;
                break;
            case "formula":
                block = new OfficeMarkupFormulaBlock(
                    GetAttribute(attributes, "cell") ?? string.Empty,
                    GetAttribute(attributes, "value") ?? GetAttribute(attributes, "formula") ?? body.Trim()) {
                    Sheet = GetAttribute(attributes, "sheet")
                };
                break;
            case "table":
            case "namedtable":
            case "named-table":
                var namedTable = new OfficeMarkupNamedTableBlock(
                    GetAttribute(attributes, "name") ?? "Table1",
                    GetAttribute(attributes, "range") ?? GetAttribute(attributes, "address") ?? string.Empty);
                if (TryGetBoolean(attributes, "header", out var hasHeader) || TryGetBoolean(attributes, "hasHeader", out hasHeader)) {
                    namedTable.HasHeader = hasHeader;
                }

                block = namedTable;
                break;
            case "format":
            case "formatting":
                block = new OfficeMarkupFormattingBlock(GetAttribute(attributes, "target") ?? GetAttribute(attributes, "range") ?? string.Empty) {
                    Style = GetAttribute(attributes, "style"),
                    NumberFormat = GetAttribute(attributes, "numberFormat") ?? GetAttribute(attributes, "format")
                };
                break;
            case "textbox":
                block = new OfficeMarkupTextBoxBlock(body.Trim()) {
                    Style = GetAttribute(attributes, "style")
                };
                break;
            case "columns":
                block = new OfficeMarkupColumnsBlock {
                    Gap = GetAttribute(attributes, "gap")
                };
                break;
            case "column":
            case "left":
            case "right":
                block = new OfficeMarkupColumnBlock(normalized, body.Trim()) {
                    Width = GetAttribute(attributes, "width")
                };
                break;
            case "card":
                block = new OfficeMarkupCardBlock(body.Trim()) {
                    Title = GetAttribute(attributes, "title"),
                    Style = GetAttribute(attributes, "style")
                };
                break;
            default:
                block = new OfficeMarkupExtensionBlock(command, attributes, body);
                diagnostics.Add(new OfficeMarkupDiagnostic(
                    OfficeMarkupDiagnosticSeverity.Info,
                    $"OfficeIMO markup directive '{command}' was preserved as an extension node.",
                    block));
                break;
        }

        block.SourceText = sourceText;
        ApplyPlacement(block, attributes);
        CopyAttributes(attributes, block.Attributes);
        return block;
    }

    private static OfficeSyntaxDirective ReadColonDirective(string[] lines, ref int index) {
        var startIndex = index;
        var header = lines[index].Trim();
        var tokens = Tokenize(header.Substring(2));
        var command = tokens.Count > 0 ? tokens[0] : string.Empty;
        var attributes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        AddInlineAttributes(tokens.Skip(1), attributes);

        var body = new StringBuilder();
        if (ShouldReadColonDirectiveBody(command, attributes)) {
            while (index + 1 < lines.Length) {
                var next = lines[index + 1].TrimStart();
                if (next.StartsWith("::", StringComparison.Ordinal)
                    || next.StartsWith("@", StringComparison.Ordinal)
                    || string.Equals(next.Trim(), "---", StringComparison.Ordinal)) {
                    break;
                }

                index++;
                body.AppendLine(lines[index]);
            }
        }

        var sourceText = string.Join("\n", lines.Skip(startIndex).Take(index - startIndex + 1));
        return new OfficeSyntaxDirective(command, attributes, body.ToString().Trim('\r', '\n'), sourceText);
    }

    private static bool ShouldReadColonDirectiveBody(string command, IDictionary<string, string> attributes) {
        switch (NormalizeCommand(command)) {
            case "notes":
            case "mermaid":
            case "textbox":
            case "card":
            case "column":
            case "left":
            case "right":
            case "range":
                return true;
            case "chart":
                return string.IsNullOrWhiteSpace(GetAttribute(attributes, "source"))
                    && string.IsNullOrWhiteSpace(GetAttribute(attributes, "range"));
            case "formula":
                return string.IsNullOrWhiteSpace(GetAttribute(attributes, "value"))
                    && string.IsNullOrWhiteSpace(GetAttribute(attributes, "formula"));
            case "image":
                return string.IsNullOrWhiteSpace(GetAttribute(attributes, "src"))
                    && string.IsNullOrWhiteSpace(GetAttribute(attributes, "source"));
            case "sheet":
                return string.IsNullOrWhiteSpace(GetAttribute(attributes, "name"));
            default:
                return false;
        }
    }

    private static OfficeSyntaxDirective ReadAtDirective(string[] lines, ref int index) {
        var startIndex = index;
        var header = lines[index].Trim();
        var afterAt = header.Substring(1).Trim();
        var commandTokens = Tokenize(afterAt.Split('{')[0]);
        var command = commandTokens.Count > 0 ? commandTokens[0] : string.Empty;
        var attributes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        AddInlineAttributes(commandTokens.Skip(1), attributes);

        var bodyLines = new List<string>();
        var openBrace = header.IndexOf('{');
        if (openBrace >= 0) {
            var afterBrace = header.Substring(openBrace + 1);
            var closeBrace = afterBrace.IndexOf('}');
            if (closeBrace >= 0) {
                bodyLines.Add(afterBrace.Substring(0, closeBrace));
            } else {
                bodyLines.Add(afterBrace);
                while (index + 1 < lines.Length) {
                    index++;
                    var line = lines[index];
                    var close = line.IndexOf('}');
                    if (close >= 0) {
                        bodyLines.Add(line.Substring(0, close));
                        break;
                    }

                    bodyLines.Add(line);
                }
            }
        }

        ParseAttributeText(string.Join("\n", bodyLines), attributes);
        var sourceText = string.Join("\n", lines.Skip(startIndex).Take(index - startIndex + 1));
        return new OfficeSyntaxDirective(command, attributes, string.Empty, sourceText);
    }

    private static void ParseAttributeText(string text, IDictionary<string, string> attributes) {
        foreach (var rawLine in (text ?? string.Empty).Split('\n')) {
            var line = rawLine.Trim().TrimEnd(',');
            if (line.Length == 0) {
                continue;
            }

            if (TryParseAttributeLine(line, attributes)) {
                continue;
            }

            AddInlineAttributes(Tokenize(line), attributes);
        }
    }

    private static bool TryParseAttributeLine(string line, IDictionary<string, string> attributes) {
        var trimmed = line.Trim();
        if (trimmed.Length == 0 || trimmed.StartsWith("#", StringComparison.Ordinal)) {
            return false;
        }

        var separator = trimmed.IndexOf(':');
        if (separator <= 0) {
            return false;
        }

        var key = trimmed.Substring(0, separator).Trim();
        if (key.Length == 0 || key.Any(char.IsWhiteSpace)) {
            return false;
        }

        attributes[key] = trimmed.Substring(separator + 1).Trim().Trim('"');
        return true;
    }

    private static List<string> Tokenize(string line) {
        var tokens = new List<string>();
        if (string.IsNullOrWhiteSpace(line)) {
            return tokens;
        }

        var builder = new StringBuilder();
        var quote = '\0';
        for (int i = 0; i < line.Length; i++) {
            var ch = line[i];
            if (quote != '\0') {
                if (ch == quote) {
                    quote = '\0';
                } else {
                    builder.Append(ch);
                }

                continue;
            }

            if (ch == '"' || ch == '\'') {
                quote = ch;
                continue;
            }

            if (char.IsWhiteSpace(ch)) {
                if (builder.Length > 0) {
                    tokens.Add(builder.ToString());
                    builder.Clear();
                }

                continue;
            }

            builder.Append(ch);
        }

        if (builder.Length > 0) {
            tokens.Add(builder.ToString());
        }

        return tokens;
    }

    private static void AddInlineAttributes(IEnumerable<string> tokens, IDictionary<string, string> attributes) {
        foreach (var token in tokens) {
            var index = token.IndexOf('=');
            if (index < 0) {
                attributes[token] = "true";
                continue;
            }

            var key = token.Substring(0, index).Trim();
            var value = token.Substring(index + 1).Trim().Trim('"');
            if (key.Length > 0) {
                attributes[key] = value;
            }
        }
    }

    private static string? GetAttribute(IDictionary<string, string> attributes, string name) =>
        attributes.TryGetValue(name, out var value) ? value : null;

    private static bool TryGetInt32(IDictionary<string, string> attributes, string name, out int value) {
        value = 0;
        return attributes.TryGetValue(name, out var text) && !string.IsNullOrWhiteSpace(text) && int.TryParse(text, out value);
    }

    private static bool TryGetBoolean(IDictionary<string, string> attributes, string name, out bool value) {
        value = false;
        if (!attributes.TryGetValue(name, out var text) || string.IsNullOrWhiteSpace(text)) {
            return false;
        }

        if (bool.TryParse(text, out value)) {
            return true;
        }

        if (string.Equals(text, "yes", StringComparison.OrdinalIgnoreCase) || string.Equals(text, "1", StringComparison.OrdinalIgnoreCase)) {
            value = true;
            return true;
        }

        if (string.Equals(text, "no", StringComparison.OrdinalIgnoreCase) || string.Equals(text, "0", StringComparison.OrdinalIgnoreCase)) {
            value = false;
            return true;
        }

        return false;
    }

    private sealed class OfficeSyntaxDirective {
        public OfficeSyntaxDirective(string command, Dictionary<string, string> attributes, string body, string sourceText) {
            Command = command;
            Attributes = attributes;
            Body = body;
            SourceText = sourceText;
        }

        public string Command { get; }
        public Dictionary<string, string> Attributes { get; }
        public string Body { get; }
        public string SourceText { get; }
    }

    private static MarkdownReaderOptions CreateMarkdownOptions(OfficeMarkupParserOptions options) {
        var markdownOptions = options.MarkdownOptions ?? MarkdownReaderOptions.CreateOfficeIMOProfile();
        RegisterOfficeFences(markdownOptions);
        return markdownOptions;
    }

    private static void RegisterOfficeFences(MarkdownReaderOptions options) {
        AddFence(options, "OfficeIMO Markup", OfficeLanguages, "officeimo");
        AddFence(options, "Mermaid diagram", new[] { "mermaid" }, "diagram");
    }

    private static void AddFence(MarkdownReaderOptions options, string name, IEnumerable<string> languages, string semanticKind) {
        var languageList = languages.ToArray();
        bool alreadyRegistered = options.FencedBlockExtensions.Any(extension =>
            extension.Languages.Any(language => languageList.Any(candidate => string.Equals(candidate, language, StringComparison.OrdinalIgnoreCase))));
        if (alreadyRegistered) {
            return;
        }

        options.FencedBlockExtensions.Add(new MarkdownFencedBlockExtension(
            name,
            languageList,
            context => new SemanticFencedBlock(semanticKind, context.InfoString, context.Content, context.Caption)));
    }

    private static void MapMarkdownBlocks(
        IEnumerable<IMarkdownBlock> markdownBlocks,
        IList<OfficeMarkupBlock> target,
        OfficeMarkupProfile profile,
        IList<OfficeMarkupDiagnostic> diagnostics) {
        foreach (var markdownBlock in markdownBlocks) {
            var mapped = MapMarkdownBlock(markdownBlock, profile, diagnostics);
            if (mapped != null) {
                target.Add(mapped);
            }
        }
    }

    private static OfficeMarkupBlock? MapMarkdownBlock(
        IMarkdownBlock markdownBlock,
        OfficeMarkupProfile profile,
        IList<OfficeMarkupDiagnostic> diagnostics) {
        switch (markdownBlock) {
            case HeadingBlock heading:
                return new OfficeMarkupHeadingBlock(heading.Level, heading.Text) {
                    SourceText = markdownBlock.RenderMarkdown()
                };

            case ParagraphBlock paragraph:
                return new OfficeMarkupParagraphBlock(ToPlainText(paragraph.Inlines)) {
                    SourceText = markdownBlock.RenderMarkdown()
                };

            case UnorderedListBlock unordered:
                return MapList(unordered.ListItems, ordered: false, start: 1, profile, diagnostics, markdownBlock.RenderMarkdown());

            case OrderedListBlock ordered:
                return MapList(ordered.ListItems, ordered: true, start: ordered.Start, profile, diagnostics, markdownBlock.RenderMarkdown());

            case CodeBlock code when IsMermaid(code.Language):
                return new OfficeMarkupDiagramBlock("mermaid", code.Content) {
                    SourceText = markdownBlock.RenderMarkdown()
                };

            case CodeBlock code:
                return new OfficeMarkupCodeBlock(code.Language, code.Content) {
                    SourceText = markdownBlock.RenderMarkdown()
                };

            case ImageBlock image:
                return new OfficeMarkupImageBlock(image.Path, image.PlainAlt ?? image.Alt, image.Title, image.Width, image.Height) {
                    SourceText = markdownBlock.RenderMarkdown()
                };

            case TableBlock table:
                return MapTable(table, markdownBlock.RenderMarkdown());

            case SemanticFencedBlock semantic when string.Equals(semantic.SemanticKind, "diagram", StringComparison.OrdinalIgnoreCase):
                return new OfficeMarkupDiagramBlock(semantic.Language, semantic.Content) {
                    SourceText = markdownBlock.RenderMarkdown()
                };

            case SemanticFencedBlock semantic:
                return MapOfficeExtension(semantic, profile, diagnostics, markdownBlock.RenderMarkdown());

            default:
                return new OfficeMarkupRawMarkdownBlock(markdownBlock.RenderMarkdown()) {
                    SourceText = markdownBlock.RenderMarkdown()
                };
        }
    }

    private static OfficeMarkupListBlock MapList(
        IReadOnlyList<ListItem> source,
        bool ordered,
        int start,
        OfficeMarkupProfile profile,
        IList<OfficeMarkupDiagnostic> diagnostics,
        string sourceText) {
        var list = new OfficeMarkupListBlock(ordered, start) {
            SourceText = sourceText
        };

        foreach (var item in source) {
            var astItem = new OfficeMarkupListItem(ToPlainText(item.Content), item.IsTask, item.Checked);
            MapMarkdownBlocks(item.ChildBlocks, astItem.Blocks, profile, diagnostics);
            list.Items.Add(astItem);
        }

        return list;
    }

    private static OfficeMarkupTableBlock MapTable(TableBlock source, string sourceText) {
        var table = new OfficeMarkupTableBlock {
            SourceText = sourceText
        };

        foreach (var header in source.Headers) {
            table.Headers.Add(header ?? string.Empty);
        }

        foreach (var row in source.Rows) {
            table.Rows.Add((row ?? Array.Empty<string>()).Select(cell => cell ?? string.Empty).ToArray());
        }

        return table;
    }

    private static OfficeMarkupBlock MapOfficeExtension(
        SemanticFencedBlock semantic,
        OfficeMarkupProfile profile,
        IList<OfficeMarkupDiagnostic> diagnostics,
        string sourceText) {
        var directive = OfficeMarkupDirective.Parse(semantic.Language, semantic.Content);
        OfficeMarkupBlock block;
        switch (NormalizeCommand(directive.Command)) {
            case "slide":
                block = CreateSlide(directive, profile, diagnostics);
                break;
            case "pagebreak":
            case "page-break":
                block = new OfficeMarkupPageBreakBlock();
                break;
            case "section":
                block = CreateSection(directive, profile, diagnostics);
                break;
            case "header":
            case "footer":
                block = new OfficeMarkupHeaderFooterBlock(
                    NormalizeCommand(directive.Command),
                    GetAttribute(directive, "text") ?? directive.Body);
                break;
            case "toc":
            case "tableofcontents":
            case "table-of-contents":
                block = CreateToc(directive);
                break;
            case "sheet":
                block = new OfficeMarkupSheetBlock(GetAttribute(directive, "name") ?? directive.Body.Trim());
                break;
            case "range":
                block = CreateRange(directive);
                break;
            case "formula":
                block = CreateFormula(directive);
                break;
            case "table":
            case "namedtable":
            case "named-table":
                block = CreateNamedTable(directive);
                break;
            case "chart":
                block = CreateChart(directive);
                break;
            case "format":
            case "formatting":
                block = CreateFormatting(directive);
                break;
            case "textbox":
                block = new OfficeMarkupTextBoxBlock(directive.Body.Trim()) {
                    Style = GetAttribute(directive, "style")
                };
                break;
            case "columns":
                block = new OfficeMarkupColumnsBlock {
                    Gap = GetAttribute(directive, "gap")
                };
                break;
            case "column":
            case "left":
            case "right":
                block = new OfficeMarkupColumnBlock(NormalizeCommand(directive.Command), directive.Body.Trim()) {
                    Width = GetAttribute(directive, "width")
                };
                break;
            case "card":
                block = new OfficeMarkupCardBlock(directive.Body.Trim()) {
                    Title = GetAttribute(directive, "title"),
                    Style = GetAttribute(directive, "style")
                };
                break;
            default:
                block = new OfficeMarkupExtensionBlock(directive.Command, directive.Attributes, directive.Body);
                diagnostics.Add(new OfficeMarkupDiagnostic(
                    OfficeMarkupDiagnosticSeverity.Warning,
                    $"Unknown OfficeIMO markup directive '{directive.Command}' was preserved as an extension node.",
                    block));
                break;
        }

        block.SourceText = sourceText;
        ApplyPlacement(block, directive.Attributes);
        CopyAttributes(directive.Attributes, block.Attributes);
        return block;
    }

    private static OfficeMarkupSlideBlock CreateSlide(
        OfficeMarkupDirective directive,
        OfficeMarkupProfile profile,
        IList<OfficeMarkupDiagnostic> diagnostics) {
        var slide = new OfficeMarkupSlideBlock(GetAttribute(directive, "title")) {
            Layout = GetAttribute(directive, "layout"),
            Section = GetAttribute(directive, "section"),
            Transition = GetAttribute(directive, "transition"),
            Background = GetAttribute(directive, "background"),
            Notes = GetAttribute(directive, "notes"),
            Placement = GetAttribute(directive, "placement")
        };
        if (TryGetInt32(directive, "columns", out var columns)) {
            slide.Columns = columns;
        }

        if (!string.IsNullOrWhiteSpace(directive.Body)) {
            var nested = MarkdownReader.Parse(directive.Body, CreateNestedMarkdownOptions());
            MapMarkdownBlocks(nested.Blocks, slide.Blocks, profile, diagnostics);
        }

        return slide;
    }

    private static OfficeMarkupSectionBlock CreateSection(
        OfficeMarkupDirective directive,
        OfficeMarkupProfile profile,
        IList<OfficeMarkupDiagnostic> diagnostics) {
        var section = new OfficeMarkupSectionBlock(GetAttribute(directive, "name") ?? GetAttribute(directive, "title")) {
            PageSize = GetAttribute(directive, "pageSize") ?? GetAttribute(directive, "size"),
            Orientation = GetAttribute(directive, "orientation")
        };

        if (!string.IsNullOrWhiteSpace(directive.Body)) {
            var nested = MarkdownReader.Parse(directive.Body, CreateNestedMarkdownOptions());
            MapMarkdownBlocks(nested.Blocks, section.Blocks, profile, diagnostics);
        }

        return section;
    }

    private static OfficeMarkupTableOfContentsBlock CreateToc(OfficeMarkupDirective directive) {
        var toc = new OfficeMarkupTableOfContentsBlock {
            Title = GetAttribute(directive, "title")
        };
        if (TryGetInt32(directive, "min", out var min) || TryGetInt32(directive, "minLevel", out min)) {
            toc.MinLevel = min;
        }

        if (TryGetInt32(directive, "max", out var max) || TryGetInt32(directive, "maxLevel", out max)) {
            toc.MaxLevel = max;
        }

        return toc;
    }

    private static OfficeMarkupRangeBlock CreateRange(OfficeMarkupDirective directive) {
        var range = new OfficeMarkupRangeBlock(GetAttribute(directive, "address") ?? GetAttribute(directive, "range") ?? string.Empty) {
            Sheet = GetAttribute(directive, "sheet")
        };

        foreach (var row in ParseDelimitedRows(directive.Body)) {
            range.Values.Add(row);
        }

        return range;
    }

    private static OfficeMarkupFormulaBlock CreateFormula(OfficeMarkupDirective directive) {
        return new OfficeMarkupFormulaBlock(
            GetAttribute(directive, "cell") ?? string.Empty,
            GetAttribute(directive, "value") ?? GetAttribute(directive, "formula") ?? directive.Body.Trim()) {
            Sheet = GetAttribute(directive, "sheet")
        };
    }

    private static OfficeMarkupNamedTableBlock CreateNamedTable(OfficeMarkupDirective directive) {
        var table = new OfficeMarkupNamedTableBlock(
            GetAttribute(directive, "name") ?? "Table1",
            GetAttribute(directive, "range") ?? GetAttribute(directive, "address") ?? string.Empty);
        if (TryGetBoolean(directive, "header", out var hasHeader) || TryGetBoolean(directive, "hasHeader", out hasHeader)) {
            table.HasHeader = hasHeader;
        }

        return table;
    }

    private static OfficeMarkupChartBlock CreateChart(OfficeMarkupDirective directive) {
        var chart = new OfficeMarkupChartBlock(GetAttribute(directive, "type") ?? GetAttribute(directive, "chartType") ?? "column") {
            Title = GetAttribute(directive, "title"),
            Source = GetAttribute(directive, "source") ?? GetAttribute(directive, "range"),
            Sheet = GetAttribute(directive, "sheet")
        };

        foreach (var row in ParseDelimitedRows(directive.Body)) {
            chart.Data.Add(row);
        }

        return chart;
    }

    private static OfficeMarkupFormattingBlock CreateFormatting(OfficeMarkupDirective directive) {
        return new OfficeMarkupFormattingBlock(GetAttribute(directive, "target") ?? GetAttribute(directive, "range") ?? string.Empty) {
            Style = GetAttribute(directive, "style"),
            NumberFormat = GetAttribute(directive, "numberFormat") ?? GetAttribute(directive, "format")
        };
    }

    private static MarkdownReaderOptions CreateNestedMarkdownOptions() {
        var options = MarkdownReaderOptions.CreateOfficeIMOProfile();
        RegisterOfficeFences(options);
        return options;
    }

    private static string ToPlainText(InlineSequence sequence) {
        var sb = new StringBuilder();
        ((IPlainTextMarkdownInline)sequence).AppendPlainText(sb);
        return sb.ToString();
    }

    private static bool IsMermaid(string language) =>
        string.Equals(language, "mermaid", StringComparison.OrdinalIgnoreCase);

    private static string NormalizeCommand(string value) =>
        (value ?? string.Empty).Trim().Replace("_", "-").ToLowerInvariant();

    private static void CopyAttributes(IDictionary<string, string> source, IDictionary<string, string> target) {
        foreach (var pair in source) {
            target[pair.Key] = pair.Value;
        }
    }

    private static void ApplyPlacement(OfficeMarkupBlock block, IDictionary<string, string> attributes) {
        var placement = CreatePlacement(attributes);
        if (placement == null) {
            return;
        }

        switch (block) {
            case OfficeMarkupImageBlock image:
                image.Placement = placement;
                break;
            case OfficeMarkupDiagramBlock diagram:
                diagram.Placement = placement;
                break;
            case OfficeMarkupChartBlock chart:
                chart.Placement = placement;
                break;
            case OfficeMarkupTextBoxBlock textBox:
                textBox.Placement = placement;
                break;
            case OfficeMarkupColumnsBlock columns:
                columns.Placement = placement;
                break;
            case OfficeMarkupCardBlock card:
                card.Placement = placement;
                break;
        }
    }

    private static OfficeMarkupPlacement? CreatePlacement(IDictionary<string, string> attributes) {
        var placement = new OfficeMarkupPlacement {
            X = GetAttribute(attributes, "x"),
            Y = GetAttribute(attributes, "y"),
            Width = GetAttribute(attributes, "w") ?? GetAttribute(attributes, "width"),
            Height = GetAttribute(attributes, "h") ?? GetAttribute(attributes, "height")
        };

        return placement.HasValue ? placement : null;
    }

    private static string? GetAttribute(OfficeMarkupDirective directive, string name) {
        return directive.Attributes.TryGetValue(name, out var value) ? value : null;
    }

    private static bool TryGetInt32(OfficeMarkupDirective directive, string name, out int value) {
        value = 0;
        var text = GetAttribute(directive, name);
        return !string.IsNullOrWhiteSpace(text) && int.TryParse(text, out value);
    }

    private static bool TryGetBoolean(OfficeMarkupDirective directive, string name, out bool value) {
        value = false;
        var text = GetAttribute(directive, name);
        if (string.IsNullOrWhiteSpace(text)) {
            return false;
        }

        if (bool.TryParse(text, out value)) {
            return true;
        }

        if (string.Equals(text, "yes", StringComparison.OrdinalIgnoreCase) || string.Equals(text, "1", StringComparison.OrdinalIgnoreCase)) {
            value = true;
            return true;
        }

        if (string.Equals(text, "no", StringComparison.OrdinalIgnoreCase) || string.Equals(text, "0", StringComparison.OrdinalIgnoreCase)) {
            value = false;
            return true;
        }

        return false;
    }

    private static IEnumerable<IReadOnlyList<string>> ParseDelimitedRows(string body) {
        if (string.IsNullOrWhiteSpace(body)) {
            yield break;
        }

        var lines = body.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
        foreach (var line in lines) {
            if (string.IsNullOrWhiteSpace(line)) {
                continue;
            }

            var separator = line.IndexOf('\t') >= 0 ? '\t' : ',';
            yield return line.Split(separator).Select(cell => cell.Trim()).ToArray();
        }
    }

    private sealed class OfficeMarkupDirective {
        private OfficeMarkupDirective(string command, Dictionary<string, string> attributes, string body) {
            Command = command;
            Attributes = attributes;
            Body = body;
        }

        public string Command { get; }
        public Dictionary<string, string> Attributes { get; }
        public string Body { get; }

        public static OfficeMarkupDirective Parse(string language, string content) {
            var attributes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var normalizedContent = (content ?? string.Empty).Replace("\r\n", "\n").Replace('\r', '\n');
            var lines = normalizedContent.Split('\n');
            var command = CommandFromLanguage(language);
            var index = 0;

            while (index < lines.Length && string.IsNullOrWhiteSpace(lines[index])) {
                index++;
            }

            if (string.IsNullOrEmpty(command) && index < lines.Length) {
                var firstTokens = Tokenize(lines[index]);
                if (firstTokens.Count > 0) {
                    command = firstTokens[0];
                    AddInlineAttributes(firstTokens.Skip(1), attributes);
                    index++;
                }
            } else if (index < lines.Length) {
                var firstTokens = Tokenize(lines[index]);
                if (firstTokens.Count > 0 && LooksLikeAttributes(firstTokens)) {
                    AddInlineAttributes(firstTokens, attributes);
                    index++;
                }
            }

            while (index < lines.Length) {
                var line = lines[index];
                if (string.IsNullOrWhiteSpace(line)) {
                    index++;
                    break;
                }

                if (!TryParseKeyValueLine(line, attributes)) {
                    break;
                }

                index++;
            }

            var body = string.Join("\n", lines.Skip(index)).Trim('\n');
            return new OfficeMarkupDirective(command, attributes, body);
        }

        private static string CommandFromLanguage(string language) {
            var value = (language ?? string.Empty).Trim();
            if (!value.StartsWith("officeimo-", StringComparison.OrdinalIgnoreCase)) {
                return string.Empty;
            }

            var suffix = value.Substring("officeimo-".Length);
            if (string.Equals(suffix, "presentation", StringComparison.OrdinalIgnoreCase)
                || string.Equals(suffix, "document", StringComparison.OrdinalIgnoreCase)
                || string.Equals(suffix, "workbook", StringComparison.OrdinalIgnoreCase)) {
                return string.Empty;
            }

            return suffix;
        }

        private static bool LooksLikeAttributes(IReadOnlyList<string> tokens) {
            for (int i = 0; i < tokens.Count; i++) {
                if (tokens[i].IndexOf('=') < 0) {
                    return false;
                }
            }

            return tokens.Count > 0;
        }

        private static void AddInlineAttributes(IEnumerable<string> tokens, IDictionary<string, string> attributes) {
            foreach (var token in tokens) {
                var index = token.IndexOf('=');
                if (index < 0) {
                    attributes[token] = "true";
                    continue;
                }

                var key = token.Substring(0, index).Trim();
                var value = token.Substring(index + 1).Trim().Trim('"');
                if (key.Length > 0) {
                    attributes[key] = value;
                }
            }
        }

        private static bool TryParseKeyValueLine(string line, IDictionary<string, string> attributes) {
            var trimmed = line.Trim();
            var separator = trimmed.IndexOf(':');
            if (separator <= 0) {
                return false;
            }

            var key = trimmed.Substring(0, separator).Trim();
            if (key.Length == 0 || key.Any(char.IsWhiteSpace)) {
                return false;
            }

            attributes[key] = trimmed.Substring(separator + 1).Trim();
            return true;
        }

        private static List<string> Tokenize(string line) {
            var tokens = new List<string>();
            if (string.IsNullOrWhiteSpace(line)) {
                return tokens;
            }

            var builder = new StringBuilder();
            var quote = '\0';
            for (int i = 0; i < line.Length; i++) {
                var ch = line[i];
                if (quote != '\0') {
                    if (ch == quote) {
                        quote = '\0';
                    } else {
                        builder.Append(ch);
                    }
                    continue;
                }

                if (ch == '"' || ch == '\'') {
                    quote = ch;
                    continue;
                }

                if (char.IsWhiteSpace(ch)) {
                    if (builder.Length > 0) {
                        tokens.Add(builder.ToString());
                        builder.Clear();
                    }
                    continue;
                }

                builder.Append(ch);
            }

            if (builder.Length > 0) {
                tokens.Add(builder.ToString());
            }

            return tokens;
        }
    }
}
