using System.Reflection;
using System.Xml.Linq;
using Markdig;

namespace OfficeIMO.Tests.MarkdownSuite;

internal static class MarkdigExtensionInventory {
    public static MarkdigExtensionInventoryReport Build(string repositoryRoot) {
        var rows = CreateRows();
        var reflectedMethodNames = GetReflectedPipelineMethodNames();
        return new MarkdigExtensionInventoryReport(
            GetPackageReferenceVersion(repositoryRoot),
            reflectedMethodNames,
            rows);
    }

    private static IReadOnlyList<MarkdigExtensionInventoryRow> CreateRows() => [
        Row("UseAbbreviations", "Abbreviations", "Inline abbreviation definitions and expansion", MarkdigExtensionInventoryStatus.Gap, "No OfficeIMO abbreviation parser or renderer contract exists yet.", "Decide whether abbreviation expansion belongs in core or an optional inline extension."),
        Row("UseAdvancedExtensions", "Advanced extension bundle", "Markdig convenience bundle over multiple extension families", MarkdigExtensionInventoryStatus.Intentional, "OfficeIMO should track individual feature families instead of claiming bundle parity.", "Keep this row as a roll-up guard; do not implement as a broad on switch."),
        Row("UseAlertBlocks", "Alert blocks", "GitHub-style alert/admonition blocks", MarkdigExtensionInventoryStatus.Partial, "OfficeIMO has callout blocks and GitHub-style callout parsing, but not Markdig's alert rendering callback shape.", "Align callout/alert syntax, AST fields, source spans, and renderer customization explicitly."),
        Row("UseAutoIdentifiers", "Auto identifiers", "Heading id generation and slug options", MarkdigExtensionInventoryStatus.Covered, "OfficeIMO has automatic heading ids with duplicate-slug tracking, an opt-out HTML switch, Markdig default and GitHub-compatible slug styles, GFM HTML profile wiring, heading traversal APIs, and source-backed heading syntax/native metadata.", "Keep slug-style and heading-source fixtures aligned as broader renderer profiles evolve."),
        Row("UseAutoLinks", "Extended autolinks", "Bare URL, www, and email autolinks", MarkdigExtensionInventoryStatus.Partial, "OfficeIMO has profile-sensitive bare URL/email autolinks with Markdig-style previous-character and domain-without-period options, GFM coverage, source-backed target and angle-marker metadata, and Markdown writer preservation for parsed bare and angle autolink spelling, but broader Markdig/GFM edge breadth is not complete.", "Broaden remaining GFM/Markdig autolink edge cases before promotion."),
        Row("UseBootstrap", "Bootstrap renderer helpers", "Bootstrap-oriented HTML rendering conventions", MarkdigExtensionInventoryStatus.Intentional, "This is renderer-theme behavior rather than a core Markdown syntax family for OfficeIMO.", "Keep theme/rendering presets separate from parser parity."),
        Row("UseCjkFriendlyEmphasis", "CJK-friendly emphasis", "CJK-aware emphasis delimiter behavior", MarkdigExtensionInventoryStatus.Partial, "OfficeIMO has selected CJK-adjacent emphasis regression coverage, but not a Markdig-compatible CJK emphasis option.", "Fold into the CommonMark emphasis delimiter rewrite and keep CJK-specific fixtures explicit."),
        Row("UseCitations", "Citations", "Citation inline/block syntax", MarkdigExtensionInventoryStatus.Gap, "No citation AST or renderer contract exists.", "Decide whether citations are in scope after core CommonMark/GFM closure."),
        Row("UseCustomContainers", "Custom containers", "Colon-fenced custom container blocks", MarkdigExtensionInventoryStatus.Gap, "OfficeIMO has semantic block extension seams, but not Markdig custom container syntax parity.", "Route to block parser extensions plus renderer/writer source-slice contracts."),
        Row("UseDefinitionLists", "Definition lists", "Definition list block syntax", MarkdigExtensionInventoryStatus.Partial, "OfficeIMO has structured definition-list AST, Markdig-style colon-marker term grouping, multiple-definition parsing, source/native projection, profile-correct HTML comparison coverage, grouped Markdown writer preservation for reparsing, Markdig lazy paragraph, nested block, loose-definition, edge-continuation, and empty-marker first-continuation coverage, loose-definition writer preservation, and blank-separated marker-group writer preservation, but full source-map and writer edge breadth is not closed.", "Broaden remaining Markdig definition-list source-map and writer edge cases before promotion."),
        Row("UseDiagrams", "Diagrams", "Diagram fenced-code rendering helpers", MarkdigExtensionInventoryStatus.Partial, "OfficeIMO has semantic fenced blocks and visual renderer hooks, but not Markdig diagram extension parity.", "Compare Mermaid/Nomnom-style cases and decide renderer-package ownership."),
        Row("UseEmojiAndSmiley", "Emoji and smiley", "Emoji shortcode and smiley replacement", MarkdigExtensionInventoryStatus.Gap, "OfficeIMO has emoji word-join normalization only, not shortcode/smiley expansion.", "Keep normalization separate from an optional inline replacement extension."),
        Row("UseEmphasisExtras", "Emphasis extras", "Strikethrough, inserted, marked, superscript, and subscript-style extras", MarkdigExtensionInventoryStatus.Partial, "OfficeIMO has strikethrough and highlight/mark-style inlines, but not the full Markdig emphasis-extra set.", "Inventory exact delimiter options before changing inline parsing."),
        Row("UseFigures", "Figures", "Figure block/rendering support", MarkdigExtensionInventoryStatus.Partial, "OfficeIMO has image/figure import and publisher figure rendering paths, but not Markdig figure syntax parity.", "Separate HTML-import figure recovery from Markdown parser extension support."),
        Row("UseFooters", "Footers", "Footer block syntax", MarkdigExtensionInventoryStatus.Gap, "No footer block parser or semantic node exists.", "Leave out of scope unless document footer semantics become a Markdown requirement."),
        Row("UseFootnotes", "Footnotes", "Footnote definitions and references", MarkdigExtensionInventoryStatus.Covered, "OfficeIMO has GFM footnote parsing and GitHub HTML rendering for first-reference ordering, repeated-reference backrefs, missing/unused definitions, nested block bodies, source/native label and marker spans, and structured Markdown writer roundtrip proof.", "Keep the GFM footnote fixture corpus and structured-body writer coverage current."),
        Row("UseGenericAttributes", "Generic attributes", "Attribute blocks/spans on Markdown elements", MarkdigExtensionInventoryStatus.Partial, "OfficeIMO captures fenced-code brace metadata, but not generic attributes on arbitrary blocks/inlines.", "Design attribute storage on semantic and syntax nodes before broad support."),
        Row("UseGlobalization", "Globalization", "Culture-aware extensions", MarkdigExtensionInventoryStatus.Gap, "No Markdig globalization extension equivalent is documented for OfficeIMO.", "Revisit only if a real consumer needs culture-sensitive Markdown behavior."),
        Row("UseGridTables", "Grid tables", "Pandoc-style grid tables", MarkdigExtensionInventoryStatus.Gap, "OfficeIMO has pipe tables only; grid table parsing is absent.", "Decide if grid tables belong in core or an optional extension package."),
        Row("UseJiraLinks", "Jira links", "Jira issue-link shortcuts", MarkdigExtensionInventoryStatus.Gap, "No Jira-link shortcut parser exists.", "Treat as optional link extension after core link/source mapping is stable."),
        Row("UseListExtras", "List extras", "Additional list syntaxes", MarkdigExtensionInventoryStatus.Gap, "OfficeIMO list work is focused on CommonMark/GFM task behavior, not Markdig list extras.", "Inventory Markdig list-extra syntax before choosing scope."),
        Row("UseMathematics", "Mathematics", "Inline/block math syntax and rendering", MarkdigExtensionInventoryStatus.Partial, "OfficeIMO has math-oriented semantic/rendering paths through host options, but not Markdig math delimiter parity.", "Define math parser ownership and compare inline/block math fixtures."),
        Row("UseMediaLinks", "Media links", "Shortcut media embedding links", MarkdigExtensionInventoryStatus.Partial, "OfficeIMO has image/media document semantics, but not Markdig media-link provider parity.", "Route shortcut media providers through renderer/host extension seams if in scope."),
        Row("UseNonAsciiNoEscape", "Non-ASCII no-escape rendering", "Renderer escaping policy", MarkdigExtensionInventoryStatus.Intentional, "OfficeIMO keeps escaping behavior profile/renderer-owned instead of mirroring this Markdig switch.", "Document any renderer escaping profile differences when output claims broaden."),
        Row("UsePipeTables", "Pipe tables", "GFM pipe table syntax", MarkdigExtensionInventoryStatus.Covered, "OfficeIMO has GFM pipe-table parsing with delimiter-row validation, escaped/code-span pipe handling, body-row padding/truncation, container ownership, semantic table/cell AST, syntax/native source spans, GitHub HTML rendering, and aligned Markdown writer roundtrip proof.", "Keep the GFM table fixture corpus and table-cell source-edit coverage current."),
        Row("UsePragmaLines", "Pragma lines", "Pragma line syntax", MarkdigExtensionInventoryStatus.Gap, "No pragma-line parser or semantic contract exists.", "Leave out of core unless a concrete document workflow needs it."),
        Row("UsePreciseSourceLocation", "Precise source location", "Line/column source location precision", MarkdigExtensionInventoryStatus.Partial, "OfficeIMO has syntax/source/native spans and source slices, but full lossless trivia/original mapping is still partial.", "Continue Phase 3 source-map and trivia work before claiming parity."),
        Row("UseReferralLinks", "Referral links", "HTML link rel referral policy", MarkdigExtensionInventoryStatus.Gap, "No Markdig-compatible referral-link renderer policy exists.", "Treat as renderer policy work if requested."),
        Row("UseSelfPipeline", "Self pipeline", "Pipeline-aware rendering helper", MarkdigExtensionInventoryStatus.Intentional, "This is a Markdig pipeline composition helper, not a Markdown feature OfficeIMO should mirror directly.", "Keep extension composition in OfficeIMO reader/render/write options."),
        Row("UseSmartyPants", "SmartyPants", "Typographic quote/dash replacements", MarkdigExtensionInventoryStatus.Gap, "No SmartyPants inline transform exists.", "Consider as an optional inline transform after delimiter parsing stabilizes."),
        Row("UseSoftlineBreakAsHardlineBreak", "Soft line break as hard line break", "Renderer/parser option for softbreak output", MarkdigExtensionInventoryStatus.Covered, "OfficeIMO exposes an explicit reader option that parses ordinary paragraph soft breaks as hard breaks while keeping CommonMark/GFM defaults unchanged, rendering HTML breaks, writing normalized hard-break markdown, and avoiding fake source marker metadata.", "Keep the option covered alongside paragraph/list source-map and writer fixtures."),
        Row("UseTaskLists", "Task lists", "GFM task-list markers", MarkdigExtensionInventoryStatus.Covered, "OfficeIMO has GFM task-list parsing for checked, unchecked, uppercase, nested, and invalid tight-marker cases; semantic AST flags; exact marker source spans; native snapshots/source edits; GitHub HTML rendering; and Markdown writer roundtrip proof.", "Keep the GFM fixture corpus and marker source-edit coverage current."),
        Row("UseYamlFrontMatter", "YAML front matter", "YAML front matter block parsing", MarkdigExtensionInventoryStatus.Covered, "OfficeIMO preserves YAML front matter as a top-of-document raw YAML AST payload with body and fence source spans, structured key/value helpers for simple entries, native source fields and snapshots, HTML omission, and Markdown writer roundtrip behavior.", "Keep raw YAML, parsed-entry helpers, and front-matter source-edit fixtures aligned as lossless trivia work expands.")
    ];

    private static MarkdigExtensionInventoryRow Row(string methodName, string family, string markdigScope, MarkdigExtensionInventoryStatus status, string officeImoState, string nextAction) =>
        new(methodName, family, markdigScope, status, officeImoState, nextAction);

    private static IReadOnlyList<string> GetReflectedPipelineMethodNames() {
        var builderType = typeof(MarkdownPipelineBuilder);
        return typeof(Markdig.Markdown).Assembly
            .GetExportedTypes()
            .Where(static type => type.IsAbstract && type.IsSealed)
            .SelectMany(static type => type.GetMethods(BindingFlags.Public | BindingFlags.Static))
            .Where(method => {
                var parameters = method.GetParameters();
                return parameters.Length > 0 && parameters[0].ParameterType == builderType;
            })
            .Select(static method => method.Name)
            .Distinct(StringComparer.Ordinal)
            .OrderBy(static name => name, StringComparer.Ordinal)
            .ToArray();
    }

    private static string GetPackageReferenceVersion(string repositoryRoot) {
        string projectPath = Path.Combine(repositoryRoot, "OfficeIMO.Tests", "OfficeIMO.Tests.csproj");
        var document = XDocument.Load(projectPath);
        var ns = document.Root?.Name.Namespace ?? XNamespace.None;
        return document
            .Descendants(ns + "PackageReference")
            .Where(static element => string.Equals((string?)element.Attribute("Include"), "Markdig", StringComparison.OrdinalIgnoreCase))
            .Select(static element => (string?)element.Attribute("Version"))
            .SingleOrDefault() ?? throw new InvalidOperationException("Markdig package reference was not found in " + projectPath + ".");
    }
}

internal sealed class MarkdigExtensionInventoryReport {
    public MarkdigExtensionInventoryReport(string markdigVersion, IReadOnlyList<string> reflectedMethodNames, IReadOnlyList<MarkdigExtensionInventoryRow> rows) {
        MarkdigVersion = markdigVersion;
        ReflectedMethodNames = reflectedMethodNames;
        Rows = rows;
    }

    public string MarkdigVersion { get; }
    public IReadOnlyList<string> ReflectedMethodNames { get; }
    public IReadOnlyList<MarkdigExtensionInventoryRow> Rows { get; }

    public int Total => Rows.Count;
    public int Covered => Count(MarkdigExtensionInventoryStatus.Covered);
    public int Partial => Count(MarkdigExtensionInventoryStatus.Partial);
    public int Intentional => Count(MarkdigExtensionInventoryStatus.Intentional);
    public int Gap => Count(MarkdigExtensionInventoryStatus.Gap);

    public IReadOnlyList<string> MissingTrackedUseMethods =>
        ReflectedMethodNames
            .Where(static name => name.StartsWith("Use", StringComparison.Ordinal))
            .Where(static name => name != "Use")
            .Except(Rows.Select(static row => row.MethodName), StringComparer.Ordinal)
            .OrderBy(static name => name, StringComparer.Ordinal)
            .ToArray();

    public IReadOnlyList<string> ObsoleteTrackedUseMethods =>
        Rows
            .Select(static row => row.MethodName)
            .Except(ReflectedMethodNames, StringComparer.Ordinal)
            .OrderBy(static name => name, StringComparer.Ordinal)
            .ToArray();

    private int Count(MarkdigExtensionInventoryStatus status) =>
        Rows.Count(row => row.Status == status);
}

internal sealed class MarkdigExtensionInventoryRow {
    public MarkdigExtensionInventoryRow(string methodName, string family, string markdigScope, MarkdigExtensionInventoryStatus status, string officeImoState, string nextAction) {
        MethodName = methodName;
        Family = family;
        MarkdigScope = markdigScope;
        Status = status;
        OfficeImoState = officeImoState;
        NextAction = nextAction;
    }

    public string MethodName { get; }
    public string Family { get; }
    public string MarkdigScope { get; }
    public MarkdigExtensionInventoryStatus Status { get; }
    public string OfficeImoState { get; }
    public string NextAction { get; }
}

internal enum MarkdigExtensionInventoryStatus {
    Covered,
    Partial,
    Intentional,
    Gap
}
