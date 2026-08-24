using OfficeIMO.Markdown;

namespace OfficeIMO.Markup;

public static partial class OfficeMarkupParser {
    private static readonly MarkdownReaderOptions DefaultMarkdownOptions = CreateDefaultMarkdownOptions();

    private static MarkdownReaderOptions CreateDefaultMarkdownOptions() {
        var options = MarkdownReaderOptions.CreateOfficeIMOProfile();
        // OfficeMarkupParser owns front-matter extraction before the Markdown pass.
        // Keep nested Markdown parsing from cloning an option graph just to disable it again.
        options.FrontMatter = false;
        RegisterOfficeFences(options);
        return options;
    }

    private static MarkdownReaderOptions CreateMarkdownOptions(OfficeMarkupParserOptions options) {
        if (options.MarkdownOptions == null) {
            return DefaultMarkdownOptions;
        }

        var markdownOptions = options.MarkdownOptions;
        RegisterOfficeFences(markdownOptions);
        return markdownOptions;
    }

    private static void RegisterOfficeFences(MarkdownReaderOptions options) {
        AddFence(options, "OfficeIMO Markup", OfficeLanguages, "officeimo");
        AddFence(options, "Mermaid diagram", MermaidLanguages, "diagram");
    }

    private static void AddFence(MarkdownReaderOptions options, string name, string[] languages, string semanticKind) {
        for (int extensionIndex = 0; extensionIndex < options.FencedBlockExtensions.Count; extensionIndex++) {
            IReadOnlyList<string> registeredLanguages = options.FencedBlockExtensions[extensionIndex].Languages;
            for (int registeredIndex = 0; registeredIndex < registeredLanguages.Count; registeredIndex++) {
                for (int candidateIndex = 0; candidateIndex < languages.Length; candidateIndex++) {
                    if (string.Equals(
                        registeredLanguages[registeredIndex],
                        languages[candidateIndex],
                        StringComparison.OrdinalIgnoreCase)) {
                        return;
                    }
                }
            }
        }

        options.FencedBlockExtensions.Add(new MarkdownFencedBlockExtension(
            name,
            languages,
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
                return WithLazySourceText(new OfficeMarkupHeadingBlock(heading.Level, heading.Text), markdownBlock);

            case ParagraphBlock paragraph:
                return WithLazySourceText(new OfficeMarkupParagraphBlock(ToPlainText(paragraph.Inlines)), markdownBlock);

            case UnorderedListBlock unordered:
                return MapList(unordered.Items, ordered: false, start: 1, profile, diagnostics, markdownBlock);

            case OrderedListBlock ordered:
                return MapList(ordered.Items, ordered: true, start: ordered.Start, profile, diagnostics, markdownBlock);

            case CodeBlock code when IsMermaid(code.Language):
                return WithLazySourceText(new OfficeMarkupDiagramBlock("mermaid", code.Content), markdownBlock);

            case CodeBlock code:
                return WithLazySourceText(new OfficeMarkupCodeBlock(code.Language, code.Content), markdownBlock);

            case ImageBlock image:
                return WithLazySourceText(
                    new OfficeMarkupImageBlock(image.Path, image.PlainAlt ?? image.Alt, image.Title, image.Width, image.Height),
                    markdownBlock);

            case TableBlock table:
                return MapTable(table, markdownBlock);

            case SemanticFencedBlock semantic when string.Equals(semantic.SemanticKind, "diagram", StringComparison.OrdinalIgnoreCase):
                return WithLazySourceText(new OfficeMarkupDiagramBlock(semantic.Language, semantic.Content), markdownBlock);

            case SemanticFencedBlock semantic:
                return MapOfficeExtension(semantic, profile, diagnostics, markdownBlock.RenderMarkdown());

            default:
                string rendered = markdownBlock.RenderMarkdown();
                return new OfficeMarkupRawMarkdownBlock(rendered) { SourceText = rendered };
        }
    }

    private static OfficeMarkupListBlock MapList(
        IReadOnlyList<ListItem> source,
        bool ordered,
        int start,
        OfficeMarkupProfile profile,
        IList<OfficeMarkupDiagnostic> diagnostics,
        IMarkdownBlock sourceBlock) {
        var list = WithLazySourceText(new OfficeMarkupListBlock(ordered, start), sourceBlock);

        foreach (var item in source) {
            var astItem = new OfficeMarkupListItem(ToPlainText(item.Content), item.IsTask, item.Checked);
            MapMarkdownBlocks(item.ChildBlocks, astItem.Blocks, profile, diagnostics);
            list.Items.Add(astItem);
        }

        return list;
    }

    private static OfficeMarkupTableBlock MapTable(TableBlock source, IMarkdownBlock sourceBlock) {
        var table = WithLazySourceText(new OfficeMarkupTableBlock(), sourceBlock);

        foreach (var header in source.Headers) {
            table.Headers.Add(header ?? string.Empty);
        }

        foreach (var row in source.Rows) {
            table.Rows.Add((row ?? Array.Empty<string>()).Select(cell => cell ?? string.Empty).ToArray());
        }

        return table;
    }

    private static TBlock WithLazySourceText<TBlock>(TBlock target, IMarkdownBlock source)
        where TBlock : OfficeMarkupBlock {
        target.SetLazySourceText(source);
        return target;
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
            var nested = MarkdownReader.ParseSemanticProjection(directive.Body, CreateNestedMarkdownOptions());
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
            var nested = MarkdownReader.ParseSemanticProjection(directive.Body, CreateNestedMarkdownOptions());
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
}
