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
        Row("UseAbbreviations", "Abbreviations", "Inline abbreviation definitions and expansion", MarkdigExtensionInventoryStatus.Covered, "Core parser, opt-in", "Keep Markdig comparison, syntax/native/source-edit, and writer fixtures current.", "OfficeIMO has opt-in abbreviation definitions through MarkdownReaderOptions.Abbreviations, case-sensitive/later-wins document-wide definition collection, consumed definition syntax nodes including empty-title and list-item-contained definitions, AbbreviationInline semantic nodes, HTML <abbr> rendering, Markdig comparison cases for top-level text, Unicode text, unresolved bracket text, emphasis, link labels, blockquotes, lists, list-item definitions, dash/opening-punctuation boundaries, and pipe-table cells when UsePipeTables is also enabled, syntax/native metadata for visible text plus definition title source edits, nested container/table-cell AST propagation, definition-preserving Markdown writing for parse-owned definitions with front matter, empty-title, list-contained definitions, and reparse-stability coverage, list-contained definition source-token navigation and native title edits, and Markdig-style literal non-ASCII text rendering through HtmlOptions.EscapeNonAsciiText = false.", "Keep abbreviation comparison, list-contained source-token, native source-edit, and writer fixtures aligned as lossless trivia work expands."),
        Row("UseAdvancedExtensions", "Advanced extension bundle", "Markdig convenience bundle over multiple extension families", MarkdigExtensionInventoryStatus.Intentional, "Intentional bundle guard", "Keep individual feature rows authoritative; do not add a broad bundle switch.", "OfficeIMO should track individual feature families instead of claiming bundle parity.", "Keep this row as a roll-up guard; do not implement as a broad on switch."),
        Row("UseAlertBlocks", "Alert blocks", "GitHub-style alert/admonition blocks", MarkdigExtensionInventoryStatus.Covered, "Core parser plus renderer policy", "Keep upstream-style GitHub alert, syntax/native/source-edit, and writer reparse fixtures current.", "OfficeIMO has structured callout blocks with GitHub-style no-title alert parsing, Markdig-style lazy continuation and empty-body HTML behavior, titled callouts as a richer OfficeIMO extension, source/native fields for the opening marker, kind, closing marker, title, and body including mixed lazy-continuation and quoted body source spans plus source edits, Markdown writing with curated writer/reparse proof, portable fallbacks, and an opt-in Markdig-style alert HTML fallback. MarkdownReaderOptions.CalloutTitleMode now makes that boundary explicit: OfficeIMO mode keeps rich titled callouts, while MarkdigCompatible mode treats titled alert markers as ordinary blockquotes like Markdig. Expanded comparison proof covers no-title note, tip, important, warning, caution, list, nested quote, fenced code, inline-rich body, empty body, lowercase kind, malformed marker fallback, lazy-continuation body, custom alert cases, all five GitHub Docs alert examples, separated multi-alert documents, paragraph boundaries, nested-list blockquote boundaries, plus titled alert blockquote fallback.", "Keep alert comparison, source/native, and writer fixtures aligned as broader GFM and lossless trivia work expands."),
        Row("UseAutoIdentifiers", "Auto identifiers", "Heading id generation and slug options", MarkdigExtensionInventoryStatus.Covered, "Core renderer option", "Keep slug-style and source metadata fixtures current.", "OfficeIMO has automatic heading ids with duplicate-slug tracking, an opt-out HTML switch, Markdig default and GitHub-compatible slug styles, GFM HTML profile wiring, heading traversal APIs, and source-backed heading syntax/native metadata.", "Keep slug-style and heading-source fixtures aligned as broader renderer profiles evolve."),
        Row("UseAutoLinks", "Extended autolinks", "Bare URL, www, email, and selected scheme autolinks", MarkdigExtensionInventoryStatus.Covered, "Core parser, profile-gated", "Keep Markdig/GFM autolink fixtures, source metadata, writer preservation, and URL/text-rendering profile evidence current.", "OfficeIMO has profile-sensitive bare URL/email autolinks with Markdig-style previous-character, domain-without-period, query/fragment special-character, balanced-parenthesis trailing-punctuation, punctuation-before-closing-parenthesis preservation, single trailing punctuation/underscore trimming, optional trailing semicolon retention, optional trailing quote retention with paired single-quote literal fallback, lowercase www. prefix matching, optional www-host and url-host underscore rejection, optional user-info authority rejection for Markdig-compatible http/www/ftp literals, optional closing-bracket URL consumption, lowercase bare scheme matching, profile-selectable bare scheme prefixes for Markdig-compatible mailto:, ftp://, and tel: behavior while OfficeIMO/GFM can keep xmpp:, apostrophe-started bare scheme literal fallback, bare mailto path/query/fragment targets with address-only display, optional Markdig-compatible mailto semicolon and address-only colon/dash handling, Markdig-compatible href tilde and quote percent-encoding, default/Markdig-style IDNA host rendering, cmark-gfm-style percent-encoded Unicode hosts through the GFM HTML profile, literal non-ASCII display text through the explicit HTML text policy, GFM/table-cell coverage, source-backed target and angle-marker metadata, and Markdown writer preservation for parsed bare and angle autolink spelling. The focused Markdig AutoLinks and AutoLinks+PipeTables comparison lanes pass across the covered option matrix.", "Keep broader GFM fixture breadth separate from the Markdig UseAutoLinks row."),
        Row("UseBootstrap", "Bootstrap renderer helpers", "Bootstrap-oriented HTML rendering conventions", MarkdigExtensionInventoryStatus.Intentional, "Renderer theme policy", "Keep parser parity separate from optional theme presets.", "This is renderer-theme behavior rather than a core Markdown syntax family for OfficeIMO.", "Keep theme/rendering presets separate from parser parity."),
        Row("UseCjkFriendlyEmphasis", "CJK-friendly emphasis", "CJK-aware emphasis delimiter behavior", MarkdigExtensionInventoryStatus.Covered, "Core delimiter parser option", "Keep CJK delimiter comparison and source-token fixtures current as the broader emphasis parser evolves.", "OfficeIMO exposes MarkdownReaderOptions.CjkFriendlyEmphasis as an opt-in Markdig-compatible star-emphasis delimiter mode. Focused Markdig comparison covers Japanese punctuation, Chinese text with code spans, CJK-adjacent Latin punctuation, single-star emphasis, and the underscore boundary that remains literal. Native/source metadata and Markdown writer proof cover the parsed strong marker spans.", "Keep the CJK-friendly option aligned with future emphasis delimiter rewrites and do not enable it by default in CommonMark/GFM profiles."),
        Row("UseCitations", "Citations", "Citation inline/block syntax", MarkdigExtensionInventoryStatus.Gap, "Optional parser extension, deferred", "Citation AST, renderer/writer contract, and real consumer need after core/GFM closure.", "No citation AST or renderer contract exists.", "Decide whether citations are in scope after core CommonMark/GFM closure."),
        Row("UseCustomContainers", "Custom containers", "Colon-fenced custom container blocks", MarkdigExtensionInventoryStatus.Gap, "Core extension seam plus optional built-in parser", "Container parser contract, child-block source mapping, renderer/writer source-slice APIs, and Markdig fixtures.", "OfficeIMO has semantic block extension seams, but not Markdig custom container syntax parity.", "Route to block parser extensions plus renderer/writer source-slice contracts."),
        Row("UseDefinitionLists", "Definition lists", "Definition list block syntax", MarkdigExtensionInventoryStatus.Covered, "Core parser, opt-in/profile-gated", "Keep focused Markdig comparison, syntax/native/source-edit, no generated-child diagnostic, and writer reparse fixtures current.", "OfficeIMO has structured definition-list AST, Markdig-style colon-marker term grouping, multiple-definition parsing, source/native projection, profile-correct HTML comparison coverage, grouped Markdown writer preservation for reparsing, Markdig lazy paragraph soft-break preservation for single and multiple continuation lines, nested block, loose-definition, edge-continuation, setext-continuation, lazy setext-heading continuation preservation, setext-following and thematic-break-following lazy-continuation boundaries, nested-list and nested-blockquote heading/thematic interruption boundaries, nested-list and nested-blockquote lazy thematic-break boundary preservation, nested-list and nested-blockquote lazy equals-setext literal preservation, nested-blockquote lazy paragraph soft-break preservation, multi-line nested-blockquote lazy paragraph boundaries before unindented lists with source/native/writer proof, blank-separated nested-list and nested-blockquote lazy soft-break preservation with source/native/writer proof, blank-separated nested-list lazy reference-definition-looking text preservation with source/native/writer proof, nested-blockquote table-shaped lazy continuation ownership with and without pipe tables, empty-marker first-continuation coverage, same-type unordered and ordered lazy list-tail merging inside definition bodies, multi-line table-shaped lazy paragraph tails after nested list bodies, nested-list table-shaped lazy table ownership with pipe tables enabled, non-1 ordered lazy tails after nested list bodies plus active nested-blockquote ordered-list boundaries that close the definition list, ordered-parenthesis lazy delimiter splits after ordered-list bodies, escaped-pipe table-shaped lazy text rendering, pipe-table delimiter-count mismatch padding with pipe tables enabled, unindented blockquote boundaries after nested list bodies with source/native/writer proof, unindented list boundaries after nested blockquote bodies with source/native/writer proof, unindented fenced-code and raw-HTML boundaries after nested bodies with source/native/writer proof, unindented blockquote continuations inside active nested blockquotes, lazy reference-definition-looking text preservation inside nested blockquotes, mixed unordered-to-ordered list tails preserved as separate definition body children with source/native/writer proof, blank-separated indented paragraph body dedent with lazy-tail source/writer proof, and different unordered-marker split preservation through parsed marker Markdown writing, parsed and generated definition marker syntax tokens, native source-backed marker fields/source edits, loose-definition writer preservation, blank-separated marker-group writer preservation, blank-separated pre-marker term boundary proof, table-shaped continuation profile proof with literal paragraphs when tables are off and nested tables when pipe tables are on, paragraph-plus-unindented-table-shaped lazy continuation proof that preserves table-looking text when tables are off, raw HTML block lazy-continuation body rendering and reparse proof for marker-line and empty-marker forms, closed and unclosed fenced-code body lazy-continuation boundaries plus definition-body indentation normalization for marker-line and empty-marker forms, tight nested-list writer preservation, setext-continuation writer reparse proof, blank-separated setext-heading writer/reparse preservation, paragraph-plus-thematic-break writer reparse preservation, empty-marker blank-separated body source/writer preservation, typed plus source-field multiline definition-body edits that keep continuation indentation valid for simple and marker forms, and compact no generated-definition-child diagnostic proof for the final tail and nested-body probes.", "Keep definition-list comparison, native source, no generated-child diagnostic, and writer reparse fixtures aligned as lossless trivia work expands."),
        Row("UseDiagrams", "Diagrams", "Diagram fenced-code rendering helpers", MarkdigExtensionInventoryStatus.Partial, "Renderer/host policy over semantic fences", "Named diagram language mapping, renderer package ownership, source/writer behavior, and comparison fixtures.", "OfficeIMO has semantic fenced blocks and visual renderer hooks, but not Markdig diagram extension parity.", "Compare Mermaid/Nomnom-style cases and decide renderer-package ownership."),
        Row("UseEmojiAndSmiley", "Emoji and smiley", "Emoji shortcode and smiley replacement", MarkdigExtensionInventoryStatus.Gap, "Optional inline transform", "Shortcode/smiley tables, opt-in profile behavior, source metadata, writer rules, and no conflict with Unicode normalization.", "OfficeIMO has emoji word-join normalization only, not shortcode/smiley expansion.", "Keep normalization separate from an optional inline replacement extension."),
        Row("UseEmphasisExtras", "Emphasis extras", "Strikethrough, inserted, marked, superscript, and subscript-style extras", MarkdigExtensionInventoryStatus.Covered, "Core inline parser, profile-gated", "Keep delimiter fixtures aligned with GFM and lossless source work.", "OfficeIMO has strikethrough, inserted-text, highlight/mark, superscript, and subscript inline nodes with Markdig comparison cases, parser-owned source marker metadata, native projection, HTML rendering, Markdown writing, and explicit GFM single-tilde strikethrough profile coverage.", "Keep emphasis-extra delimiter cases aligned as broader GFM and lossless trivia coverage expands."),
        Row("UseFigures", "Figures", "Figure block/rendering support", MarkdigExtensionInventoryStatus.Partial, "Core image AST plus optional parser syntax", "Separate HTML-import figure recovery from Markdown figure syntax, then prove renderer/writer/source behavior.", "OfficeIMO has image/figure import and publisher figure rendering paths, but not Markdig figure syntax parity.", "Separate HTML-import figure recovery from Markdown parser extension support."),
        Row("UseFooters", "Footers", "Footer block syntax", MarkdigExtensionInventoryStatus.Gap, "Deferred document semantics", "Only implement if Markdown-authored footer semantics become a real document requirement.", "No footer block parser or semantic node exists.", "Leave out of scope unless document footer semantics become a Markdown requirement."),
        Row("UseFootnotes", "Footnotes", "Footnote definitions and references", MarkdigExtensionInventoryStatus.Covered, "Core parser, GFM profile", "Keep GFM footnote fixture corpus and structured writer proof current.", "OfficeIMO has GFM footnote parsing and GitHub HTML rendering for first-reference ordering, repeated-reference backrefs, missing/unused definitions, nested block bodies, source/native label and marker spans, and structured Markdown writer roundtrip proof.", "Keep the GFM footnote fixture corpus and structured-body writer coverage current."),
        Row(
            "UseGenericAttributes",
            "Generic attributes",
            "Attribute blocks/spans on Markdown elements",
            MarkdigExtensionInventoryStatus.Partial,
            "Core AST/source architecture",
            "Remaining arbitrary block-family parsing, complete inline-family breadth, and broader Markdown writer/source preservation across arbitrary shapes.",
            "OfficeIMO now has generic attribute storage on semantic MarkdownObject nodes and MarkdownSyntaxNode nodes, with fenced-code id/classes/attributes projected from MarkdownCodeFenceInfo through ordinary CodeBlock and SemanticFencedBlock parser paths. Default fenced-code HTML rendering projects fence-info attributes onto the HTML pre wrapper, while source-backed standalone generic attributes before fenced code render on the code element to match Markdig's UseGenericAttributes behavior. Semantic fenced-block code fallback renderers receive the attributed CodeBlock. Opt-in MarkdownReaderOptions.GenericAttributes now parses Markdig-style trailing attribute blocks for ATX headings, Setext headings, paragraphs, standalone attribute blocks before fenced code, headings, paragraphs, inline-image paragraphs in portable/Markdig profiles, root ordered lists, root unordered lists, pipe tables, OfficeIMO-default typed image blocks, dash setext/thematic forms, and indented code, root ordered/unordered list items, nested list items, blockquote-contained list items, definition-list-looking text without UseDefinitionLists, definition-list terms, definition-list definition paragraphs, and Markdig-style pipe-table cells that promote attributes to the owning table, while standalone attributes before HTML blocks are consumed without metadata and blockquote paragraph, heading, and standalone-before-blockquote attribute blocks remain literal to match Markdig's container behavior. Standalone attributes before reference-link-definition-looking lines now produce attributed literal paragraphs without registering reference definitions, with source-backed native edit proof. Standalone fenced-code, list, pipe-table, typed image-block, dash setext/thematic, and indented-code-derived paragraph attributes are source-backed in syntax/native/source-edit APIs; list attributes project to the top-level `<ol>`/`<ul>` element, pipe-table attributes project to the `<table>` element, typed image-block attributes project to the `<img>` element, dash setext/thematic attributes produce an empty h2, and indented-code attributes produce an attributed paragraph. Paragraph attribute blocks preserve Markdig's consumed separator whitespace in HTML and Markdown writing, including thematic-break-like paragraph lines such as `--- {#id}`, `*** {#id}`, and `___ {#id}`, no-space bare-URL paragraph attribute blocks such as `https://example.com{#id}` keep literal URL text and no-space Markdown writing, no-space abbreviation-ending paragraph attribute blocks target the owning paragraph like Markdig when UseAbbreviations is combined with UseGenericAttributes, ordinary no-space plain-text paragraph attribute blocks such as `word{#id}` and `C++{#id}` target the owning paragraph without stealing paired inline delimiter targets, and standalone attribute continuation lines at the end of paragraphs are consumed without metadata or rendered output like Markdig, including soft and hard line-break forms. Paragraph-contained attributes embedded at the end of nested link labels, image alt text, linked-image alt text, emphasis content, and strong content now promote to the paragraph owner like Markdig, strip the literal attribute text from nested content, and remain source-backed in syntax/native projections. List-item attribute blocks are consumed for Markdig-compatible HTML without projecting attributes onto <li>, preserve Markdig's consumed separator whitespace, write normalized trailing attribute blocks, and expose semantic attributes plus syntax/native/source-edit proof; focused Markdig interaction proof covers the same consumption behavior when UseTaskLists is also enabled. Definition-list term attributes project to `<dt>` and remain source-backed on the semantic term, while definition-value paragraph attributes are consumed without projecting onto `<dd>`, matching Markdig's rendered behavior. Footnote definition body paragraphs consume and project generic attributes when UseFootnotes and UseGenericAttributes are combined, standalone generic attributes before footnote definitions are consumed without metadata to match Markdig's boundary, and footnote references consume following attribute blocks without rendering literal text or native `attributes` metadata to match Markdig's reference behavior. No-space inline attribute blocks attach to links, reference links, images, reference images, linked images, emphasis, strong, code spans, angle autolinks, superscript, and subscript nodes. Markdig leaves strikethrough, highlight, and inserted emphasis-extra attribute blocks literal, and OfficeIMO follows that boundary. Raw inline HTML consumes a following generic attribute block without projecting it into rendered HTML, matching Markdig's rendered output for that shape. Those attributes flow through semantic/syntax storage, default HTML rendering, Markdown writing, and reparse proof for the covered shapes. Generic attribute blocks on covered block and inline shapes are source-backed as dedicated GenericAttributeBlock syntax tokens and in native projections as `attributes` source fields/metadata, with syntax navigation, preserved-trivia source slicing, snapshot, and source-edit proof. It still does not parse generic attributes for arbitrary block families or every inline family.",
            "Extend the shared attribute parser/writer to more block and inline families, then promote once writer/source propagation and token-level coverage are proven across arbitrary shapes."),
        Row("UseGlobalization", "Globalization", "Culture-aware extensions", MarkdigExtensionInventoryStatus.Gap, "Deferred compatibility option", "Only implement with a concrete culture-sensitive behavior contract and fixtures.", "No Markdig globalization extension equivalent is documented for OfficeIMO.", "Revisit only if a real consumer needs culture-sensitive Markdown behavior."),
        Row("UseGridTables", "Grid tables", "Pandoc-style grid tables", MarkdigExtensionInventoryStatus.Gap, "Optional block parser extension", "Grid table AST/source model, HTML/Markdown writer behavior, malformed-table fallback, and Markdig/Pandoc-style fixtures.", "OfficeIMO has pipe tables only; grid table parsing is absent.", "Decide if grid tables belong in core or an optional extension package."),
        Row("UseJiraLinks", "Jira links", "Jira issue-link shortcuts", MarkdigExtensionInventoryStatus.Gap, "Optional link inline extension", "Configurable issue-key resolver, renderer policy, writer preservation, and source metadata without affecting ordinary text.", "No Jira-link shortcut parser exists.", "Treat as optional link extension after core link/source mapping is stable."),
        Row("UseListExtras", "List extras", "Additional list syntaxes", MarkdigExtensionInventoryStatus.Gap, "Optional parser work after list cleanup", "Inventory Markdig list-extra syntax, choose supported forms, and prove canonical ListItem/source behavior.", "OfficeIMO list work is focused on CommonMark/GFM task behavior, not Markdig list extras.", "Inventory Markdig list-extra syntax before choosing scope."),
        Row("UseMathematics", "Mathematics", "Inline/block math syntax and rendering", MarkdigExtensionInventoryStatus.Partial, "Optional parser plus renderer/host policy", "Inline/block math delimiters, AST/source/native metadata, writer preservation, and renderer handoff contract.", "OfficeIMO has math-oriented semantic/rendering paths through host options, but not Markdig math delimiter parity.", "Define math parser ownership and compare inline/block math fixtures."),
        Row("UseMediaLinks", "Media links", "Shortcut media embedding links", MarkdigExtensionInventoryStatus.Partial, "Renderer/host policy with optional link parser", "Provider model, safe renderer output, writer preservation, and source metadata for shortcut media links.", "OfficeIMO has image/media document semantics, but not Markdig media-link provider parity.", "Route shortcut media providers through renderer/host extension seams if in scope."),
        Row("UseNonAsciiNoEscape", "Non-ASCII no-escape rendering", "Renderer escaping policy", MarkdigExtensionInventoryStatus.Covered, "Renderer escaping policy", "Keep renderer escape-policy coverage aligned as new HTML output paths are introduced.", "OfficeIMO exposes HtmlOptions.EscapeNonAsciiText so Markdig/GFM-style HTML output can keep non-ASCII visible text literal while preserving the historical .NET encoder behavior by default. The GitHub-flavored HTML profile enables literal non-ASCII text rendering, and inline text, link display text, code block text, captions, simple quote text, abbreviation output, TOC labels/titles/anchors, heading helper text and generated attributes, page titles, body/head/asset metadata, footnote ids/backrefs, code and callout classes, link/image policy attributes, raw-HTML escape output, image-blocked placeholder text, portable fallback helper text, image/link title and alt attributes, picture-source descriptor attributes, sanitizer escape output, and custom HTML render-extension helper APIs use the explicit policy. URL-bearing attributes remain routed through the URL attribute encoder.", "Keep direct encoder audits and focused non-ASCII render-policy tests current when adding new HTML output paths."),
        Row("UsePipeTables", "Pipe tables", "GFM pipe table syntax", MarkdigExtensionInventoryStatus.Covered, "Core parser, GFM profile", "Keep GFM table corpus and table-cell source-edit coverage current.", "OfficeIMO has GFM pipe-table parsing with delimiter-row validation, escaped/code-span pipe handling, body-row padding/truncation, container ownership, semantic table/cell AST, syntax/native source spans, GitHub HTML rendering, and aligned Markdown writer roundtrip proof.", "Keep the GFM table fixture corpus and table-cell source-edit coverage current."),
        Row("UsePragmaLines", "Pragma lines", "Pragma line syntax", MarkdigExtensionInventoryStatus.Gap, "Deferred metadata parser", "Only implement if a concrete workflow needs pragma metadata with source-preserving writer behavior.", "No pragma-line parser or semantic contract exists.", "Leave out of core unless a concrete document workflow needs it."),
        Row("UsePreciseSourceLocation", "Precise source location", "Line/column source location precision", MarkdigExtensionInventoryStatus.Partial, "Cross-cutting core source architecture", "Complete lossless trivia/original mapping, generated-node diagnostics, and source-edit coverage before claiming parity.", "OfficeIMO has syntax/source/native spans, source slices, original-source slices when trivia is preserved, reason-aware original-source slice failure reporting, source edits, roundtrip diagnostics, addressable native block/snapshot source fields including repeated fields by occurrence index, native inline/inline-metadata source-slice APIs for source-backed link targets, titles, formatting content, and similar metadata, source-slice APIs aligned with native source-edit targets for blocks, list item content, table cells, definition-list objects, reference definitions, reference-definition fields, and document-level source trivia, paragraph-level native projections and source slices for list-item paragraphs, custom block parser context normalized source-slice APIs for parser-created spans and relative line ranges, document-transform context normalized/original source-slice APIs for parsed model objects and syntax spans, plus document-level source trivia, snapshots, source-order enumeration, position lookup, and normalized/original source slices for blank lines, whitespace-only lines, leading/trailing horizontal whitespace, tabs, and line endings. Document-level trivia columns and generic line/column source-slice fallback now expand tabs with the same tab-stop model as source maps. Original-source slices now map line-ending-equivalent normalized spans back to the original CRLF, LF, or standalone CR spelling through a shared mapper used by parse results and transform contexts, and line-ending trivia source edits preserve original source bytes around the changed trivia when possible. Full lossless trivia/original mapping is still partial.", "Continue Phase 3 source-map and trivia work before claiming parity."),
        Row("UseReferralLinks", "Referral links", "HTML link rel referral policy", MarkdigExtensionInventoryStatus.Gap, "Renderer policy", "Only implement as an opt-in link-rendering policy with safe defaults and tests.", "No Markdig-compatible referral-link renderer policy exists.", "Treat as renderer policy work if requested."),
        Row("UseSelfPipeline", "Self pipeline", "Pipeline-aware rendering helper", MarkdigExtensionInventoryStatus.Intentional, "Intentional composition difference", "Keep extension composition in OfficeIMO options rather than mirroring Markdig pipeline helpers.", "This is a Markdig pipeline composition helper, not a Markdown feature OfficeIMO should mirror directly.", "Keep extension composition in OfficeIMO reader/render/write options."),
        Row("UseSmartyPants", "SmartyPants", "Typographic quote/dash replacements", MarkdigExtensionInventoryStatus.Gap, "Optional inline transform", "Smart punctuation transform with opt-in profile, source/edit behavior, writer policy, and escaping rules.", "No SmartyPants inline transform exists.", "Consider as an optional inline transform after delimiter parsing stabilizes."),
        Row("UseSoftlineBreakAsHardlineBreak", "Soft line break as hard line break", "Renderer/parser option for softbreak output", MarkdigExtensionInventoryStatus.Covered, "Core parser option", "Keep option covered alongside paragraph/list source-map and writer fixtures.", "OfficeIMO exposes an explicit reader option that parses ordinary paragraph soft breaks as hard breaks while keeping CommonMark/GFM defaults unchanged, rendering HTML breaks, writing normalized hard-break markdown, and avoiding fake source marker metadata.", "Keep the option covered alongside paragraph/list source-map and writer fixtures."),
        Row("UseTaskLists", "Task lists", "GFM task-list markers", MarkdigExtensionInventoryStatus.Covered, "Core parser, GFM profile", "Keep GFM task marker source-edit coverage current.", "OfficeIMO has GFM task-list parsing for checked, unchecked, uppercase, nested, and invalid tight-marker cases; semantic AST flags; exact marker source spans; native snapshots/source edits; GitHub HTML rendering; and Markdown writer roundtrip proof.", "Keep the GFM fixture corpus and marker source-edit coverage current."),
        Row("UseYamlFrontMatter", "YAML front matter", "YAML front matter block parsing", MarkdigExtensionInventoryStatus.Covered, "Core parser, OfficeIMO profile", "Keep raw YAML helpers and front-matter source-edit fixtures aligned with lossless work.", "OfficeIMO preserves YAML front matter as a top-of-document raw YAML AST payload with body and fence source spans, structured key/value helpers for simple entries, native source fields and snapshots, HTML omission, and Markdown writer roundtrip behavior.", "Keep raw YAML, parsed-entry helpers, and front-matter source-edit fixtures aligned as lossless trivia work expands.")
    ];

    private static MarkdigExtensionInventoryRow Row(string methodName, string family, string markdigScope, MarkdigExtensionInventoryStatus status, string route, string promotionBar, string officeImoState, string nextAction) =>
        new(methodName, family, markdigScope, status, route, promotionBar, officeImoState, nextAction);

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

internal enum MarkdigExtensionScopeDecision {
    CoreEngine,
    OptionalExtension,
    RendererHostPolicy,
    Deferred,
    IntentionalDifference,
    Unknown
}

internal static class MarkdigExtensionScopeDecisionExtensions {
    public static string ToDisplayText(this MarkdigExtensionScopeDecision decision) =>
        decision switch {
            MarkdigExtensionScopeDecision.CoreEngine => "Core engine",
            MarkdigExtensionScopeDecision.OptionalExtension => "Optional extension",
            MarkdigExtensionScopeDecision.RendererHostPolicy => "Renderer/host policy",
            MarkdigExtensionScopeDecision.Deferred => "Deferred",
            MarkdigExtensionScopeDecision.IntentionalDifference => "Intentional difference",
            MarkdigExtensionScopeDecision.Unknown => "Unknown",
            _ => throw new ArgumentOutOfRangeException(nameof(decision), decision, null)
        };
}

internal sealed class MarkdigExtensionInventoryRow {
    public MarkdigExtensionInventoryRow(string methodName, string family, string markdigScope, MarkdigExtensionInventoryStatus status, string route, string promotionBar, string officeImoState, string nextAction) {
        MethodName = methodName;
        Family = family;
        MarkdigScope = markdigScope;
        Status = status;
        Route = route;
        PromotionBar = promotionBar;
        OfficeImoState = officeImoState;
        NextAction = nextAction;
    }

    public string MethodName { get; }
    public string Family { get; }
    public string MarkdigScope { get; }
    public MarkdigExtensionInventoryStatus Status { get; }
    public string Route { get; }
    public MarkdigExtensionScopeDecision ScopeDecision => ClassifyScopeDecision(Route);
    public string PromotionBar { get; }
    public string OfficeImoState { get; }
    public string NextAction { get; }

    private static MarkdigExtensionScopeDecision ClassifyScopeDecision(string route) {
        if (route.Contains("Intentional", StringComparison.OrdinalIgnoreCase)) {
            return MarkdigExtensionScopeDecision.IntentionalDifference;
        }

        if (route.Contains("Deferred", StringComparison.OrdinalIgnoreCase)) {
            return MarkdigExtensionScopeDecision.Deferred;
        }

        if (route.Contains("Renderer", StringComparison.OrdinalIgnoreCase)) {
            return MarkdigExtensionScopeDecision.RendererHostPolicy;
        }

        if (route.Contains("Core", StringComparison.OrdinalIgnoreCase)) {
            return MarkdigExtensionScopeDecision.CoreEngine;
        }

        if (route.Contains("Optional", StringComparison.OrdinalIgnoreCase)) {
            return MarkdigExtensionScopeDecision.OptionalExtension;
        }

        return MarkdigExtensionScopeDecision.Unknown;
    }
}

internal enum MarkdigExtensionInventoryStatus {
    Covered,
    Partial,
    Intentional,
    Gap
}
