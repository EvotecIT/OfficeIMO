using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Markdown.Pdf;

/// <summary>
/// Markdown-aware visual profile used by the first-party Markdown to PDF adapter.
/// </summary>
public sealed partial class MarkdownPdfStyle {
    private PdfCore.PdfTheme? _documentTheme;
    private PdfCore.PanelStyle? _codeBlockPanelStyle;
    private PdfCore.PanelStyle? _semanticBlockPanelStyle;
    private PdfCore.PanelStyle? _quotePanelStyle;
    private PdfCore.PanelStyle? _calloutPanelStyle;
    private PdfCore.PanelStyle? _detailsPanelStyle;
    private PdfCore.PanelStyle? _tocPanelStyle;
    private PdfCore.PdfTableStyle? _tableStyle;
    private PdfCore.PdfTableStyle? _checklistTableStyle;
    private PdfCore.PdfTableStyle? _definitionListTableStyle;
    private PdfCore.PdfTableStyle? _frontMatterTableStyle;
    private double? _codeBlockFontSize;
    private double? _codeBlockLabelFontSize;
    private PdfCore.PdfColor? _codeBlockTextColor;
    private PdfCore.PdfColor? _codeBlockLabelColor;
    private double? _semanticBlockFontSize;
    private double? _semanticBlockLabelFontSize;
    private PdfCore.PdfColor? _semanticBlockTextColor;
    private PdfCore.PdfColor? _semanticBlockLabelColor;
    private PdfCore.PdfColor? _checklistCheckedIconColor;
    private PdfCore.PdfColor? _checklistUncheckedIconColor;
    private PdfCore.PdfColor? _checklistCheckedTextColor;
    private PdfCore.PdfColor? _checklistUncheckedTextColor;
    private PdfCore.PdfColor? _checklistCheckedFillColor;
    private PdfCore.PdfColor? _checklistUncheckedFillColor;
    private double? _documentHeaderTitleFontSize;
    private double? _documentHeaderSubtitleFontSize;
    private double? _documentHeaderMetadataFontSize;
    private PdfCore.PdfColor? _documentHeaderTitleColor;
    private PdfCore.PdfColor? _documentHeaderSubtitleColor;
    private PdfCore.PdfColor? _documentHeaderMetadataColor;
    private PdfCore.PdfColor? _documentHeaderRuleColor;
    private PdfCore.PdfColor? _tocTitleColor;
    private PdfCore.PdfColor? _tocLinkColor;
    private PdfCore.PdfColor? _tocTextColor;
    private PdfCore.PdfColor? _linkColor;
    private bool? _underlineLinks;
    private MarkdownPdfPageDecoration? _pageDecoration;
    private MarkdownPdfFigureStyle? _figureStyle;

    /// <summary>Name of the profile for diagnostics and documentation.</summary>
    public string Name { get; set; } = "Custom";

    /// <summary>Document-level PDF theme applied before Markdown blocks are rendered.</summary>
    public PdfCore.PdfTheme? DocumentTheme {
        get => _documentTheme?.Clone();
        set => _documentTheme = value?.Clone();
    }

    /// <summary>Panel style used for fenced code blocks.</summary>
    public PdfCore.PanelStyle? CodeBlockPanelStyle {
        get => _codeBlockPanelStyle?.Clone();
        set => _codeBlockPanelStyle = value?.Clone();
    }

    /// <summary>Panel style used for semantic fenced blocks such as diagrams or data blocks.</summary>
    public PdfCore.PanelStyle? SemanticBlockPanelStyle {
        get => _semanticBlockPanelStyle?.Clone();
        set => _semanticBlockPanelStyle = value?.Clone();
    }

    /// <summary>Panel style used for block quotes.</summary>
    public PdfCore.PanelStyle? QuotePanelStyle {
        get => _quotePanelStyle?.Clone();
        set => _quotePanelStyle = value?.Clone();
    }

    /// <summary>Base panel style used for callouts. The callout kind still controls the border color.</summary>
    public PdfCore.PanelStyle? CalloutPanelStyle {
        get => _calloutPanelStyle?.Clone();
        set => _calloutPanelStyle = value?.Clone();
    }

    /// <summary>Panel style used for details/summary blocks.</summary>
    public PdfCore.PanelStyle? DetailsPanelStyle {
        get => _detailsPanelStyle?.Clone();
        set => _detailsPanelStyle = value?.Clone();
    }

    /// <summary>Panel style used for generated Markdown table-of-contents blocks.</summary>
    public PdfCore.PanelStyle? TocPanelStyle {
        get => _tocPanelStyle?.Clone();
        set => _tocPanelStyle = value?.Clone();
    }

    /// <summary>Table style used for Markdown pipe tables.</summary>
    public PdfCore.PdfTableStyle? TableStyle {
        get => _tableStyle?.Clone();
        set => _tableStyle = value?.Clone();
    }

    /// <summary>Table style used for task-list checklists.</summary>
    public PdfCore.PdfTableStyle? ChecklistTableStyle {
        get => _checklistTableStyle?.Clone();
        set => _checklistTableStyle = value?.Clone();
    }

    /// <summary>Table style used for definition lists.</summary>
    public PdfCore.PdfTableStyle? DefinitionListTableStyle {
        get => _definitionListTableStyle?.Clone();
        set => _definitionListTableStyle = value?.Clone();
    }

    /// <summary>Table style used when front matter is rendered into the PDF body.</summary>
    public PdfCore.PdfTableStyle? FrontMatterTableStyle {
        get => _frontMatterTableStyle?.Clone();
        set => _frontMatterTableStyle = value?.Clone();
    }

    /// <summary>Font size for fenced code block content.</summary>
    public double? CodeBlockFontSize {
        get => _codeBlockFontSize;
        set => _codeBlockFontSize = ValidateOptionalPositive(value, nameof(CodeBlockFontSize));
    }

    /// <summary>Font size for the optional fenced code language label.</summary>
    public double? CodeBlockLabelFontSize {
        get => _codeBlockLabelFontSize;
        set => _codeBlockLabelFontSize = ValidateOptionalPositive(value, nameof(CodeBlockLabelFontSize));
    }

    /// <summary>Text color for fenced code block content.</summary>
    public PdfCore.PdfColor? CodeBlockTextColor {
        get => _codeBlockTextColor;
        set => _codeBlockTextColor = value;
    }

    /// <summary>Text color for the optional fenced code language label.</summary>
    public PdfCore.PdfColor? CodeBlockLabelColor {
        get => _codeBlockLabelColor;
        set => _codeBlockLabelColor = value;
    }

    /// <summary>Font size for semantic fenced block content.</summary>
    public double? SemanticBlockFontSize {
        get => _semanticBlockFontSize;
        set => _semanticBlockFontSize = ValidateOptionalPositive(value, nameof(SemanticBlockFontSize));
    }

    /// <summary>Font size for semantic fenced block labels.</summary>
    public double? SemanticBlockLabelFontSize {
        get => _semanticBlockLabelFontSize;
        set => _semanticBlockLabelFontSize = ValidateOptionalPositive(value, nameof(SemanticBlockLabelFontSize));
    }

    /// <summary>Text color for semantic fenced block content.</summary>
    public PdfCore.PdfColor? SemanticBlockTextColor {
        get => _semanticBlockTextColor;
        set => _semanticBlockTextColor = value;
    }

    /// <summary>Text color for semantic fenced block labels.</summary>
    public PdfCore.PdfColor? SemanticBlockLabelColor {
        get => _semanticBlockLabelColor;
        set => _semanticBlockLabelColor = value;
    }

    /// <summary>Icon color used for checked Markdown task-list items.</summary>
    public PdfCore.PdfColor? ChecklistCheckedIconColor {
        get => _checklistCheckedIconColor;
        set => _checklistCheckedIconColor = value;
    }

    /// <summary>Icon color used for unchecked Markdown task-list items.</summary>
    public PdfCore.PdfColor? ChecklistUncheckedIconColor {
        get => _checklistUncheckedIconColor;
        set => _checklistUncheckedIconColor = value;
    }

    /// <summary>Text color used for checked Markdown task-list items.</summary>
    public PdfCore.PdfColor? ChecklistCheckedTextColor {
        get => _checklistCheckedTextColor;
        set => _checklistCheckedTextColor = value;
    }

    /// <summary>Text color used for unchecked Markdown task-list items.</summary>
    public PdfCore.PdfColor? ChecklistUncheckedTextColor {
        get => _checklistUncheckedTextColor;
        set => _checklistUncheckedTextColor = value;
    }

    /// <summary>Optional row fill used for checked Markdown task-list items.</summary>
    public PdfCore.PdfColor? ChecklistCheckedFillColor {
        get => _checklistCheckedFillColor;
        set => _checklistCheckedFillColor = value;
    }

    /// <summary>Optional row fill used for unchecked Markdown task-list items.</summary>
    public PdfCore.PdfColor? ChecklistUncheckedFillColor {
        get => _checklistUncheckedFillColor;
        set => _checklistUncheckedFillColor = value;
    }

    /// <summary>Font size used for front matter title blocks.</summary>
    public double? DocumentHeaderTitleFontSize {
        get => _documentHeaderTitleFontSize;
        set => _documentHeaderTitleFontSize = ValidateOptionalPositive(value, nameof(DocumentHeaderTitleFontSize));
    }

    /// <summary>Font size used for front matter subtitles or descriptions.</summary>
    public double? DocumentHeaderSubtitleFontSize {
        get => _documentHeaderSubtitleFontSize;
        set => _documentHeaderSubtitleFontSize = ValidateOptionalPositive(value, nameof(DocumentHeaderSubtitleFontSize));
    }

    /// <summary>Font size used for front matter author/date/tags metadata lines.</summary>
    public double? DocumentHeaderMetadataFontSize {
        get => _documentHeaderMetadataFontSize;
        set => _documentHeaderMetadataFontSize = ValidateOptionalPositive(value, nameof(DocumentHeaderMetadataFontSize));
    }

    /// <summary>Text color used for front matter title blocks.</summary>
    public PdfCore.PdfColor? DocumentHeaderTitleColor {
        get => _documentHeaderTitleColor;
        set => _documentHeaderTitleColor = value;
    }

    /// <summary>Text color used for front matter subtitles or descriptions.</summary>
    public PdfCore.PdfColor? DocumentHeaderSubtitleColor {
        get => _documentHeaderSubtitleColor;
        set => _documentHeaderSubtitleColor = value;
    }

    /// <summary>Text color used for front matter author/date/tags metadata lines.</summary>
    public PdfCore.PdfColor? DocumentHeaderMetadataColor {
        get => _documentHeaderMetadataColor;
        set => _documentHeaderMetadataColor = value;
    }

    /// <summary>Rule color used below rendered front matter document headers.</summary>
    public PdfCore.PdfColor? DocumentHeaderRuleColor {
        get => _documentHeaderRuleColor;
        set => _documentHeaderRuleColor = value;
    }

    /// <summary>Title color used by generated Markdown table-of-contents blocks.</summary>
    public PdfCore.PdfColor? TocTitleColor {
        get => _tocTitleColor;
        set => _tocTitleColor = value;
    }

    /// <summary>Link color used by generated Markdown table-of-contents entries.</summary>
    public PdfCore.PdfColor? TocLinkColor {
        get => _tocLinkColor;
        set => _tocLinkColor = value;
    }

    /// <summary>Marker and secondary text color used by generated Markdown table-of-contents blocks.</summary>
    public PdfCore.PdfColor? TocTextColor {
        get => _tocTextColor;
        set => _tocTextColor = value;
    }

    /// <summary>Text color used for ordinary Markdown links in paragraphs, lists, tables, panels, and definitions.</summary>
    public PdfCore.PdfColor? LinkColor {
        get => _linkColor;
        set => _linkColor = value;
    }

    /// <summary>Whether ordinary Markdown links should be underlined.</summary>
    public bool? UnderlineLinks {
        get => _underlineLinks;
        set => _underlineLinks = value;
    }

    /// <summary>Optional page-level decoration profile applied through the shared PDF engine.</summary>
    public MarkdownPdfPageDecoration? PageDecoration {
        get => _pageDecoration?.Clone();
        set => _pageDecoration = value?.Clone();
    }

    /// <summary>Figure styling used for Markdown images and generated visual blocks.</summary>
    public MarkdownPdfFigureStyle? FigureStyle {
        get => _figureStyle?.Clone();
        set => _figureStyle = value?.Clone();
    }

    /// <summary>Creates a built-in visual profile.</summary>
    internal static MarkdownPdfStyle Create(OfficeVisualThemeKind kind) {
        switch (kind) {
            case OfficeVisualThemeKind.Plain:
                return Plain();
            case OfficeVisualThemeKind.WordLike:
                return WordLike();
            case OfficeVisualThemeKind.TechnicalDocument:
                return TechnicalDocument();
            case OfficeVisualThemeKind.GitHubLike:
                return GitHubLike();
            case OfficeVisualThemeKind.Compact:
                return Compact();
            case OfficeVisualThemeKind.Report:
                return Report();
            default:
                throw new ArgumentOutOfRangeException(nameof(kind), kind, "Unsupported Markdown PDF theme kind.");
        }
    }

    /// <summary>Plain profile with minimal additional styling.</summary>
    internal static MarkdownPdfStyle Plain() => new MarkdownPdfStyle {
        Name = "Plain",
        CodeBlockPanelStyle = Panel(248, 250, 252, 203, 213, 225, 0.5, 8, 7, 4, 8),
        SemanticBlockPanelStyle = Panel(248, 250, 252, 148, 163, 184, 0.5, 8, 7, 4, 8),
        QuotePanelStyle = Panel(249, 250, 251, 148, 163, 184, 0.75, 10, 8, 0, 8),
        CalloutPanelStyle = Panel(248, 250, 252, 37, 99, 235, 1, 10, 8, 4, 8),
        DetailsPanelStyle = Panel(249, 250, 251, 203, 213, 225, 0.5, 8, 6, 0, 6),
        TocPanelStyle = Panel(248, 250, 252, 203, 213, 225, 0.6, 10, 8, 4, 9),
        TableStyle = MarkdownTable(PdfCore.TableStyles.Light(), 10.25, 10.25, 1.22, 6, 5, 5, 10),
        ChecklistTableStyle = ChecklistTable(203, 213, 225, 31, 41, 55, 10, 4),
        DefinitionListTableStyle = DefinitionTable(226, 232, 240, 0.4, 5, 4),
        FrontMatterTableStyle = FrontMatterTable(),
        CodeBlockFontSize = 9.5,
        CodeBlockLabelFontSize = 8,
        CodeBlockTextColor = Color(31, 41, 55),
        CodeBlockLabelColor = Color(71, 85, 105),
        SemanticBlockFontSize = 9.5,
        SemanticBlockLabelFontSize = 8,
        SemanticBlockTextColor = Color(31, 41, 55),
        SemanticBlockLabelColor = Color(71, 85, 105),
        ChecklistCheckedIconColor = Color(22, 163, 74),
        ChecklistUncheckedIconColor = Color(100, 116, 139),
        ChecklistCheckedTextColor = Color(71, 85, 105),
        ChecklistUncheckedTextColor = Color(31, 41, 55),
        ChecklistCheckedFillColor = Color(240, 253, 244),
        DocumentHeaderTitleFontSize = 25,
        DocumentHeaderSubtitleFontSize = 12.5,
        DocumentHeaderMetadataFontSize = 9.5,
        DocumentHeaderTitleColor = Color(17, 24, 39),
        DocumentHeaderSubtitleColor = Color(71, 85, 105),
        DocumentHeaderMetadataColor = Color(100, 116, 139),
        DocumentHeaderRuleColor = Color(203, 213, 225),
        TocTitleColor = Color(17, 24, 39),
        TocLinkColor = Color(37, 99, 235),
        TocTextColor = Color(100, 116, 139),
        LinkColor = Color(37, 99, 235),
        UnderlineLinks = true,
        FigureStyle = MarkdownPdfFigureStyle.Plain()
    };

    /// <summary>Neutral Word-like profile for general Markdown documents.</summary>
    internal static MarkdownPdfStyle WordLike() {
        MarkdownPdfStyle theme = Plain();
        theme.Name = "WordLike";
        theme.DocumentTheme = PdfCore.PdfTheme.WordLike();
        return theme;
    }

    /// <summary>Polished profile for technical guides, specs, and READMEs.</summary>
    internal static MarkdownPdfStyle TechnicalDocument() => new MarkdownPdfStyle {
        Name = "TechnicalDocument",
        PageDecoration = MarkdownPdfPageDecoration.TechnicalDocument(),
        DocumentTheme = PdfCore.PdfTheme.TechnicalDocument(),
        CodeBlockPanelStyle = PanelWithLeftRule(246, 248, 250, 208, 215, 222, 9, 105, 218, 0.5, 3, 10, 8, 4, 9),
        SemanticBlockPanelStyle = PanelWithLeftRule(240, 249, 255, 186, 230, 253, 2, 132, 199, 0.5, 3, 10, 8, 4, 9),
        QuotePanelStyle = PanelWithLeftRule(248, 250, 252, 226, 232, 240, 100, 116, 139, 0.4, 3, 10, 8, 2, 9),
        CalloutPanelStyle = PanelWithLeftRule(248, 250, 252, 226, 232, 240, 37, 99, 235, 0.4, 3, 11, 8, 4, 9),
        DetailsPanelStyle = Panel(241, 245, 249, 203, 213, 225, 0.5, 9, 7, 2, 7),
        TocPanelStyle = PanelWithLeftRule(248, 250, 252, 226, 232, 240, 9, 105, 218, 0.4, 3, 11, 8, 4, 10),
        TableStyle = TechnicalTable(),
        ChecklistTableStyle = ChecklistTable(203, 213, 225, 15, 23, 42, 10.5, 5),
        DefinitionListTableStyle = DefinitionTable(203, 213, 225, 0.35, 6, 5),
        FrontMatterTableStyle = FrontMatterTable(),
        CodeBlockFontSize = 9.5,
        CodeBlockLabelFontSize = 8,
        CodeBlockTextColor = Color(15, 23, 42),
        CodeBlockLabelColor = Color(9, 105, 218),
        SemanticBlockFontSize = 9.25,
        SemanticBlockLabelFontSize = 8,
        SemanticBlockTextColor = Color(15, 23, 42),
        SemanticBlockLabelColor = Color(2, 132, 199),
        ChecklistCheckedIconColor = Color(22, 163, 74),
        ChecklistUncheckedIconColor = Color(100, 116, 139),
        ChecklistCheckedTextColor = Color(71, 85, 105),
        ChecklistUncheckedTextColor = Color(15, 23, 42),
        ChecklistCheckedFillColor = Color(240, 253, 244),
        DocumentHeaderTitleFontSize = 26,
        DocumentHeaderSubtitleFontSize = 12.5,
        DocumentHeaderMetadataFontSize = 9.5,
        DocumentHeaderTitleColor = Color(15, 23, 42),
        DocumentHeaderSubtitleColor = Color(51, 65, 85),
        DocumentHeaderMetadataColor = Color(100, 116, 139),
        DocumentHeaderRuleColor = Color(9, 105, 218),
        TocTitleColor = Color(15, 23, 42),
        TocLinkColor = Color(9, 105, 218),
        TocTextColor = Color(100, 116, 139),
        LinkColor = Color(9, 105, 218),
        UnderlineLinks = true,
        FigureStyle = MarkdownPdfFigureStyle.Framed(
            Color(248, 250, 252),
            Color(203, 213, 225),
            Color(71, 85, 105),
            Color(100, 116, 139),
            borderWidth: 0.45)
    };

    /// <summary>GitHub-inspired profile for README-style exports.</summary>
    internal static MarkdownPdfStyle GitHubLike() => new MarkdownPdfStyle {
        Name = "GitHubLike",
        DocumentTheme = PdfCore.PdfTheme.WordLike(),
        CodeBlockPanelStyle = Panel(246, 248, 250, 208, 215, 222, 0.5, 9, 8, 4, 9),
        SemanticBlockPanelStyle = Panel(246, 248, 250, 139, 148, 158, 0.6, 9, 8, 4, 9),
        QuotePanelStyle = PanelWithLeftRule(255, 255, 255, 208, 215, 222, 208, 215, 222, 0.0, 3, 10, 7, 2, 8),
        CalloutPanelStyle = Panel(246, 248, 250, 9, 105, 218, 0.8, 10, 8, 4, 9),
        DetailsPanelStyle = Panel(246, 248, 250, 208, 215, 222, 0.5, 8, 7, 2, 7),
        TocPanelStyle = Panel(246, 248, 250, 208, 215, 222, 0.5, 10, 8, 4, 9),
        TableStyle = GitHubTable(),
        ChecklistTableStyle = ChecklistTable(208, 215, 222, 36, 41, 47, 10, 5),
        DefinitionListTableStyle = DefinitionTable(208, 215, 222, 0.5, 6, 5),
        FrontMatterTableStyle = GitHubTable(),
        CodeBlockFontSize = 9.25,
        CodeBlockLabelFontSize = 8,
        CodeBlockTextColor = Color(36, 41, 47),
        CodeBlockLabelColor = Color(87, 96, 106),
        SemanticBlockFontSize = 9.25,
        SemanticBlockLabelFontSize = 8,
        SemanticBlockTextColor = Color(36, 41, 47),
        SemanticBlockLabelColor = Color(87, 96, 106),
        ChecklistCheckedIconColor = Color(26, 127, 55),
        ChecklistUncheckedIconColor = Color(87, 96, 106),
        ChecklistCheckedTextColor = Color(87, 96, 106),
        ChecklistUncheckedTextColor = Color(36, 41, 47),
        ChecklistCheckedFillColor = Color(246, 248, 250),
        DocumentHeaderTitleFontSize = 24,
        DocumentHeaderSubtitleFontSize = 12,
        DocumentHeaderMetadataFontSize = 9,
        DocumentHeaderTitleColor = Color(36, 41, 47),
        DocumentHeaderSubtitleColor = Color(87, 96, 106),
        DocumentHeaderMetadataColor = Color(87, 96, 106),
        DocumentHeaderRuleColor = Color(208, 215, 222),
        TocTitleColor = Color(36, 41, 47),
        TocLinkColor = Color(9, 105, 218),
        TocTextColor = Color(87, 96, 106),
        LinkColor = Color(9, 105, 218),
        UnderlineLinks = true,
        FigureStyle = MarkdownPdfFigureStyle.Framed(
            Color(246, 248, 250),
            Color(208, 215, 222),
            Color(87, 96, 106),
            Color(87, 96, 106),
            borderWidth: 0.45)
    };

    /// <summary>Compact profile for dense technical notes.</summary>
    internal static MarkdownPdfStyle Compact() => new MarkdownPdfStyle {
        Name = "Compact",
        DocumentTheme = PdfCore.PdfTheme.Compact(),
        CodeBlockPanelStyle = Panel(248, 250, 252, 203, 213, 225, 0.4, 7, 5, 2, 5),
        SemanticBlockPanelStyle = Panel(241, 245, 249, 148, 163, 184, 0.4, 7, 5, 2, 5),
        QuotePanelStyle = PanelWithLeftRule(248, 250, 252, 226, 232, 240, 148, 163, 184, 0.0, 2.5, 8, 5, 1, 5),
        CalloutPanelStyle = Panel(248, 250, 252, 37, 99, 235, 0.7, 8, 5, 2, 5),
        DetailsPanelStyle = Panel(249, 250, 251, 203, 213, 225, 0.4, 7, 5, 1, 5),
        TocPanelStyle = Panel(248, 250, 252, 226, 232, 240, 0.4, 8, 6, 2, 6),
        TableStyle = CompactTable(),
        ChecklistTableStyle = ChecklistTable(226, 232, 240, 31, 41, 55, 9.5, 3),
        DefinitionListTableStyle = DefinitionTable(226, 232, 240, 0.35, 4, 3),
        FrontMatterTableStyle = CompactTable(),
        CodeBlockFontSize = 8.75,
        CodeBlockLabelFontSize = 7.5,
        CodeBlockTextColor = Color(31, 41, 55),
        CodeBlockLabelColor = Color(71, 85, 105),
        SemanticBlockFontSize = 8.75,
        SemanticBlockLabelFontSize = 7.5,
        SemanticBlockTextColor = Color(31, 41, 55),
        SemanticBlockLabelColor = Color(71, 85, 105),
        ChecklistCheckedIconColor = Color(22, 163, 74),
        ChecklistUncheckedIconColor = Color(100, 116, 139),
        ChecklistCheckedTextColor = Color(71, 85, 105),
        ChecklistUncheckedTextColor = Color(31, 41, 55),
        ChecklistCheckedFillColor = Color(240, 253, 244),
        DocumentHeaderTitleFontSize = 20,
        DocumentHeaderSubtitleFontSize = 10.5,
        DocumentHeaderMetadataFontSize = 8.5,
        DocumentHeaderTitleColor = Color(31, 41, 55),
        DocumentHeaderSubtitleColor = Color(71, 85, 105),
        DocumentHeaderMetadataColor = Color(100, 116, 139),
        DocumentHeaderRuleColor = Color(203, 213, 225),
        TocTitleColor = Color(31, 41, 55),
        TocLinkColor = Color(37, 99, 235),
        TocTextColor = Color(100, 116, 139),
        LinkColor = Color(37, 99, 235),
        UnderlineLinks = true,
        FigureStyle = MarkdownPdfFigureStyle.Plain()
    };

    /// <summary>Report-oriented profile with stronger hierarchy and tables.</summary>
    internal static MarkdownPdfStyle Report() => new MarkdownPdfStyle {
        Name = "Report",
        PageDecoration = MarkdownPdfPageDecoration.Report(),
        DocumentTheme = PdfCore.PdfTheme.Report(),
        CodeBlockPanelStyle = Panel(248, 250, 252, 191, 219, 254, 0.6, 9, 7, 4, 8),
        SemanticBlockPanelStyle = Panel(239, 246, 255, 96, 165, 250, 0.7, 9, 7, 4, 8),
        QuotePanelStyle = PanelWithLeftRule(248, 250, 252, 226, 232, 240, 30, 64, 175, 0.0, 3, 10, 8, 2, 9),
        CalloutPanelStyle = Panel(239, 246, 255, 37, 99, 235, 0.9, 10, 8, 4, 9),
        DetailsPanelStyle = Panel(241, 245, 249, 148, 163, 184, 0.5, 9, 7, 2, 7),
        TocPanelStyle = PanelWithLeftRule(239, 246, 255, 191, 219, 254, 30, 64, 175, 0.5, 3, 11, 8, 4, 10),
        TableStyle = ReportTable(),
        ChecklistTableStyle = ChecklistTable(191, 219, 254, 30, 41, 59, 10.5, 5),
        DefinitionListTableStyle = DefinitionTable(191, 219, 254, 0.4, 6, 5),
        FrontMatterTableStyle = ReportTable(),
        CodeBlockFontSize = 9.5,
        CodeBlockLabelFontSize = 8,
        CodeBlockTextColor = Color(30, 41, 59),
        CodeBlockLabelColor = Color(30, 64, 175),
        SemanticBlockFontSize = 9.25,
        SemanticBlockLabelFontSize = 8,
        SemanticBlockTextColor = Color(30, 41, 59),
        SemanticBlockLabelColor = Color(30, 64, 175),
        ChecklistCheckedIconColor = Color(22, 163, 74),
        ChecklistUncheckedIconColor = Color(71, 85, 105),
        ChecklistCheckedTextColor = Color(71, 85, 105),
        ChecklistUncheckedTextColor = Color(30, 41, 59),
        ChecklistCheckedFillColor = Color(239, 246, 255),
        DocumentHeaderTitleFontSize = 27,
        DocumentHeaderSubtitleFontSize = 13,
        DocumentHeaderMetadataFontSize = 9.5,
        DocumentHeaderTitleColor = Color(30, 41, 59),
        DocumentHeaderSubtitleColor = Color(30, 64, 175),
        DocumentHeaderMetadataColor = Color(71, 85, 105),
        DocumentHeaderRuleColor = Color(30, 64, 175),
        TocTitleColor = Color(30, 41, 59),
        TocLinkColor = Color(30, 64, 175),
        TocTextColor = Color(71, 85, 105),
        LinkColor = Color(30, 64, 175),
        UnderlineLinks = true,
        FigureStyle = MarkdownPdfFigureStyle.Framed(
            Color(239, 246, 255),
            Color(191, 219, 254),
            Color(30, 64, 175),
            Color(71, 85, 105),
            borderWidth: 0.5)
    };

    /// <summary>Creates a copy of this visual theme.</summary>
    public MarkdownPdfStyle Clone() => new MarkdownPdfStyle {
        Name = Name,
        DocumentTheme = _documentTheme,
        CodeBlockPanelStyle = _codeBlockPanelStyle,
        SemanticBlockPanelStyle = _semanticBlockPanelStyle,
        QuotePanelStyle = _quotePanelStyle,
        CalloutPanelStyle = _calloutPanelStyle,
        DetailsPanelStyle = _detailsPanelStyle,
        TocPanelStyle = _tocPanelStyle,
        TableStyle = _tableStyle,
        ChecklistTableStyle = _checklistTableStyle,
        DefinitionListTableStyle = _definitionListTableStyle,
        FrontMatterTableStyle = _frontMatterTableStyle,
        CodeBlockFontSize = _codeBlockFontSize,
        CodeBlockLabelFontSize = _codeBlockLabelFontSize,
        CodeBlockTextColor = _codeBlockTextColor,
        CodeBlockLabelColor = _codeBlockLabelColor,
        SemanticBlockFontSize = _semanticBlockFontSize,
        SemanticBlockLabelFontSize = _semanticBlockLabelFontSize,
        SemanticBlockTextColor = _semanticBlockTextColor,
        SemanticBlockLabelColor = _semanticBlockLabelColor,
        ChecklistCheckedIconColor = _checklistCheckedIconColor,
        ChecklistUncheckedIconColor = _checklistUncheckedIconColor,
        ChecklistCheckedTextColor = _checklistCheckedTextColor,
        ChecklistUncheckedTextColor = _checklistUncheckedTextColor,
        ChecklistCheckedFillColor = _checklistCheckedFillColor,
        ChecklistUncheckedFillColor = _checklistUncheckedFillColor,
        DocumentHeaderTitleFontSize = _documentHeaderTitleFontSize,
        DocumentHeaderSubtitleFontSize = _documentHeaderSubtitleFontSize,
        DocumentHeaderMetadataFontSize = _documentHeaderMetadataFontSize,
        DocumentHeaderTitleColor = _documentHeaderTitleColor,
        DocumentHeaderSubtitleColor = _documentHeaderSubtitleColor,
        DocumentHeaderMetadataColor = _documentHeaderMetadataColor,
        DocumentHeaderRuleColor = _documentHeaderRuleColor,
        TocTitleColor = _tocTitleColor,
        TocLinkColor = _tocLinkColor,
        TocTextColor = _tocTextColor,
        LinkColor = _linkColor,
        UnderlineLinks = _underlineLinks,
        PageDecoration = _pageDecoration,
        FigureStyle = _figureStyle
    };

    internal PdfCore.PdfTheme? DocumentThemeSnapshot => _documentTheme?.Clone();
    internal PdfCore.PanelStyle CodeBlockPanelStyleSnapshot => (_codeBlockPanelStyle ?? Plain()._codeBlockPanelStyle)!.Clone();
    internal PdfCore.PanelStyle SemanticBlockPanelStyleSnapshot => (_semanticBlockPanelStyle ?? Plain()._semanticBlockPanelStyle)!.Clone();
    internal PdfCore.PanelStyle QuotePanelStyleSnapshot => (_quotePanelStyle ?? Plain()._quotePanelStyle)!.Clone();
    internal PdfCore.PanelStyle DetailsPanelStyleSnapshot => (_detailsPanelStyle ?? Plain()._detailsPanelStyle)!.Clone();
    internal PdfCore.PanelStyle TocPanelStyleSnapshot => (_tocPanelStyle ?? Plain()._tocPanelStyle)!.Clone();
    internal PdfCore.PdfTableStyle TableStyleSnapshot => (_tableStyle ?? PdfCore.TableStyles.Light()).Clone();
    internal PdfCore.PdfTableStyle ChecklistTableStyleSnapshot => (_checklistTableStyle ?? Plain()._checklistTableStyle)!.Clone();
    internal PdfCore.PdfTableStyle DefinitionListTableStyleSnapshot => (_definitionListTableStyle ?? Plain()._definitionListTableStyle)!.Clone();
    internal PdfCore.PdfTableStyle FrontMatterTableStyleSnapshot => (_frontMatterTableStyle ?? PdfCore.TableStyles.Minimal()).Clone();
    internal double CodeBlockFontSizeSnapshot => _codeBlockFontSize ?? 9.5;
    internal double CodeBlockLabelFontSizeSnapshot => _codeBlockLabelFontSize ?? 8;
    internal PdfCore.PdfColor CodeBlockTextColorSnapshot => _codeBlockTextColor ?? Color(31, 41, 55);
    internal PdfCore.PdfColor CodeBlockLabelColorSnapshot => _codeBlockLabelColor ?? Color(71, 85, 105);
    internal double SemanticBlockFontSizeSnapshot => _semanticBlockFontSize ?? CodeBlockFontSizeSnapshot;
    internal double SemanticBlockLabelFontSizeSnapshot => _semanticBlockLabelFontSize ?? CodeBlockLabelFontSizeSnapshot;
    internal PdfCore.PdfColor SemanticBlockTextColorSnapshot => _semanticBlockTextColor ?? CodeBlockTextColorSnapshot;
    internal PdfCore.PdfColor SemanticBlockLabelColorSnapshot => _semanticBlockLabelColor ?? CodeBlockLabelColorSnapshot;
    internal PdfCore.PdfColor ChecklistCheckedIconColorSnapshot => _checklistCheckedIconColor ?? Plain()._checklistCheckedIconColor ?? Color(22, 163, 74);
    internal PdfCore.PdfColor ChecklistUncheckedIconColorSnapshot => _checklistUncheckedIconColor ?? Plain()._checklistUncheckedIconColor ?? Color(100, 116, 139);
    internal PdfCore.PdfColor ChecklistCheckedTextColorSnapshot => _checklistCheckedTextColor ?? Plain()._checklistCheckedTextColor ?? Color(71, 85, 105);
    internal PdfCore.PdfColor ChecklistUncheckedTextColorSnapshot => _checklistUncheckedTextColor ?? Plain()._checklistUncheckedTextColor ?? Color(31, 41, 55);
    internal PdfCore.PdfColor? ChecklistCheckedFillColorSnapshot => _checklistCheckedFillColor ?? Plain()._checklistCheckedFillColor;
    internal PdfCore.PdfColor? ChecklistUncheckedFillColorSnapshot => _checklistUncheckedFillColor ?? Plain()._checklistUncheckedFillColor;
    internal double DocumentHeaderTitleFontSizeSnapshot => _documentHeaderTitleFontSize ?? Plain()._documentHeaderTitleFontSize ?? 25;
    internal double DocumentHeaderSubtitleFontSizeSnapshot => _documentHeaderSubtitleFontSize ?? Plain()._documentHeaderSubtitleFontSize ?? 12.5;
    internal double DocumentHeaderMetadataFontSizeSnapshot => _documentHeaderMetadataFontSize ?? Plain()._documentHeaderMetadataFontSize ?? 9.5;
    internal PdfCore.PdfColor DocumentHeaderTitleColorSnapshot => _documentHeaderTitleColor ?? Plain()._documentHeaderTitleColor ?? Color(17, 24, 39);
    internal PdfCore.PdfColor DocumentHeaderSubtitleColorSnapshot => _documentHeaderSubtitleColor ?? Plain()._documentHeaderSubtitleColor ?? Color(71, 85, 105);
    internal PdfCore.PdfColor DocumentHeaderMetadataColorSnapshot => _documentHeaderMetadataColor ?? Plain()._documentHeaderMetadataColor ?? Color(100, 116, 139);
    internal PdfCore.PdfColor DocumentHeaderRuleColorSnapshot => _documentHeaderRuleColor ?? Plain()._documentHeaderRuleColor ?? Color(203, 213, 225);
    internal PdfCore.PdfColor TocTitleColorSnapshot => _tocTitleColor ?? Plain()._tocTitleColor ?? Color(17, 24, 39);
    internal PdfCore.PdfColor TocLinkColorSnapshot => _tocLinkColor ?? Plain()._tocLinkColor ?? Color(37, 99, 235);
    internal PdfCore.PdfColor TocTextColorSnapshot => _tocTextColor ?? Plain()._tocTextColor ?? Color(100, 116, 139);
    internal PdfCore.PdfColor LinkColorSnapshot => _linkColor ?? _tocLinkColor ?? Plain()._linkColor ?? Color(37, 99, 235);
    internal bool UnderlineLinksSnapshot => _underlineLinks ?? Plain()._underlineLinks ?? true;
    internal MarkdownPdfFigureStyle FigureStyleSnapshot => (_figureStyle ?? Plain()._figureStyle ?? MarkdownPdfFigureStyle.Plain()).Clone();

    internal void ApplyPageDecorations(PdfCore.PdfDocument pdf, PdfCore.PdfOptions options) {
        if (pdf == null) {
            throw new ArgumentNullException(nameof(pdf));
        }

        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        _pageDecoration?.Apply(pdf, options);
    }

    internal PdfCore.PanelStyle CreateCalloutPanelStyle(string? kind) {
        PdfCore.PanelStyle style = (_calloutPanelStyle ?? Plain()._calloutPanelStyle)!.Clone();
        style.BorderColor = GetCalloutColor(kind);
        PdfCore.PdfPanelBorder? leftBorder = style.LeftBorder;
        if (leftBorder != null) {
            style.LeftBorder = new PdfCore.PdfPanelBorder {
                Color = GetCalloutColor(kind),
                Width = leftBorder.Width
            };
        }

        return style;
    }

}
