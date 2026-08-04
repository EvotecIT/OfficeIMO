namespace OfficeIMO.Tests.MarkdownSuite;

internal static class MarkdigExtensionCompatibilityMatrixWriter {
    public static string Write(MarkdigExtensionInventoryReport report) {
        var sb = new StringBuilder();

        sb.AppendLine("# OfficeIMO.Markdown Markdig Extension Compatibility Matrix");
        sb.AppendLine();
        sb.AppendLine($"This matrix turns the Markdig `{report.MarkdigVersion}` extension inventory into execution lanes. It is generated from `MarkdigExtensionInventory`, so the open work stays tied to reflected Markdig pipeline entry points instead of drifting into ad hoc fixture lists.");
        sb.AppendLine();
        sb.AppendLine("Use it as the control board for parity slices:");
        sb.AppendLine();
        sb.AppendLine("- If the `Engine parser` lane is open, improve `OfficeIMO.Markdown` behavior before adding more proof.");
        sb.AppendLine("- If only `Proof` is open, the behavior already exists and needs focused Markdig/source/writer evidence.");
        sb.AppendLine("- If `Decision` says optional, deferred, renderer policy, or intentional, make the scope call before implementing parser behavior.");
        sb.AppendLine();
        sb.AppendLine("## Summary");
        sb.AppendLine();
        sb.AppendLine($"Current inventory: {report.Total} Markdig extension-family rows; {report.Covered} covered, {report.Partial} partial, {report.Intentional} intentional, {report.Gap} gap.");
        sb.AppendLine();
        sb.AppendLine("| Metric | Count |");
        sb.AppendLine("| --- | ---: |");
        sb.AppendLine($"| Markdig extension-family rows | {report.Total} |");
        sb.AppendLine($"| Covered | {report.Covered} |");
        sb.AppendLine($"| Partial | {report.Partial} |");
        sb.AppendLine($"| Intentional | {report.Intentional} |");
        sb.AppendLine($"| Gap | {report.Gap} |");
        sb.AppendLine();
        sb.AppendLine("## Execution Matrix");
        sb.AppendLine();
        sb.AppendLine("| Markdig entry point | Status | Decision | Engine parser | AST/source | Writer/render | Proof | Next non-looping action |");
        sb.AppendLine("| --- | --- | --- | --- | --- | --- | --- | --- |");

        foreach (var row in report.Rows) {
            MatrixCells cells = CreateCells(row);
            sb.AppendLine($"| `{row.MethodName}` | `{row.Status}` | {EscapeTable(cells.Decision)} | {EscapeTable(cells.EngineParser)} | {EscapeTable(cells.AstSource)} | {EscapeTable(cells.WriterRender)} | {EscapeTable(cells.Proof)} | {EscapeTable(cells.NextAction)} |");
        }

        sb.AppendLine();
        sb.AppendLine("## Work Checklist");
        sb.AppendLine();
        sb.AppendLine("- [ ] Pick one row and one open lane before implementation starts.");
        sb.AppendLine("- [ ] If the row needs engine work, change parser/AST/source/writer/render behavior in the owning layer first.");
        sb.AppendLine("- [ ] If the row only needs proof, add focused Markdig comparisons, source/native snapshots, writer checks, renderer checks, or generated inventory assertions.");
        sb.AppendLine("- [ ] If the row is optional, deferred, renderer-owned, or intentional, record that scope decision before adding syntax.");
        sb.AppendLine("- [ ] Promote a row to `Covered` only when the matrix lanes and the inventory promotion bar agree.");
        sb.AppendLine();
        sb.AppendLine("## Immediate Queue");
        sb.AppendLine();
        sb.AppendLine("- [x] Keep `UseGenericAttributes` opt-in and bounded to supported semantic owners. Callouts, details, and custom containers now share source-backed HTML/Markdown ownership; unsupported Markdig targets remain literal until adopted end to end.");
        var definitionLists = report.Rows.FirstOrDefault(static row => row.MethodName == "UseDefinitionLists");
        if (definitionLists?.Status == MarkdigExtensionInventoryStatus.Covered) {
            sb.AppendLine("- [x] Keep `UseDefinitionLists` covered while broader source/trivia work evolves; do not reopen it without new Markdig evidence.");
        } else {
            sb.AppendLine("- [ ] Promote or explicitly bound `UseDefinitionLists` after closing remaining source-map and writer edge breadth.");
        }
        sb.AppendLine("- [x] Keep the `UseAlertBlocks` titled-callout boundary explicit: OfficeIMO mode keeps rich titles; Markdig-compatible mode treats titled markers as ordinary blockquotes.");
        sb.AppendLine("- [x] Keep `UsePreciseSourceLocation` explicitly partial relative to the comparison surface but complete for OfficeIMO's bounded public contract: source-backed fields expose exact normalized spans and inspected original mappings; generated nodes remain spanless, unavailable mappings carry reasons, and arbitrary semantic mutation uses normalized writing.");

        return sb.ToString().Replace("\r\n", "\n");
    }

    private static MatrixCells CreateCells(MarkdigExtensionInventoryRow row) {
        if (row.Status == MarkdigExtensionInventoryStatus.Covered) {
            return new MatrixCells(row.ScopeDecision.ToDisplayText(), "Covered", "Covered", "Covered", "Covered", row.NextAction);
        }

        if (row.Status == MarkdigExtensionInventoryStatus.Intentional) {
            return new MatrixCells(row.ScopeDecision.ToDisplayText(), "Not planned", "Not planned", "Intentional difference", "Documented", row.NextAction);
        }

        switch (row.MethodName) {
            case "UseAlertBlocks":
                return new MatrixCells("Core plus renderer policy", "No-title, empty, lazy-continuation, lowercase, malformed-marker, and titled-boundary behavior covered", "Callout/alert marker, kind, title, body, lazy-continuation body, and quote source fields are exposed", "Has opt-in Markdig alert HTML fallback plus curated writer/reparse proof", "Has expanded alert comparison plus lazy-continuation syntax/native/source-edit proof; needs upstream-style GFM sweep", row.NextAction);
            case "UseCjkFriendlyEmphasis":
                return new MatrixCells("Core delimiter option", "Partial delimiter behavior", "Needs delimiter token proof", "Mostly shared with emphasis rendering", "Needs CJK comparison fixtures", row.NextAction);
            case "UseCustomContainers":
                return new MatrixCells("Optional/core block extension", "Root, nested, blockquote-contained, list-child, list-item-contained, unclosed, shorter-fence, and trailing-text fence cases support Markdig-style div/class HTML for scoped cases", "Source-backed syntax/native opening fence, info, body, closing fence, source slices, snapshots, caret lookup, source edits, child block ownership, and list-item custom-container token remapping exist for scoped cases", "Parsed-block Markdown writing and default HTML output exist for scoped cases; generated nested custom-container writing lengthens outer colon fences for stable reparse, including list-item-contained containers; tight-list custom-container HTML now matches Markdig and routes syntax/type renderer overrides through the shared dispatcher, while broader writer seams remain partial", "Has focused Markdig comparison plus syntax/native/source-edit/renderer-override and generated nested-writer reparse proof; needs broader blockquote/container ownership and writer breadth", row.NextAction);
            case "UseDefinitionLists":
                return new MatrixCells("Core opt-in parser", "Partial", "Needs remaining source-map breadth", "Needs writer/reparse edge breadth", "Needs focused edge proof before promotion", row.NextAction);
            case "UseDiagrams":
                return new MatrixCells("Renderer/host policy", "Semantic fences exist", "Needs language/source mapping decision", "Needs renderer package ownership", "Needs renderer comparison fixtures", row.NextAction);
            case "UseEmojiAndSmiley":
                return new MatrixCells("Optional inline transform", "Missing shortcode/smiley transform", "Needs source metadata policy", "Needs writer literal/normalized policy", "Needs opt-in transform fixtures", row.NextAction);
            case "UseFigures":
                return new MatrixCells("Core image plus optional syntax", "Partial image/figure behavior", "Needs Markdown figure syntax source model", "Needs renderer/writer contract", "Needs syntax-vs-import proof", row.NextAction);
            case "UseGenericAttributes":
                return new MatrixCells("Core engine", "Supported block/inline owners plus callout, details, and custom-container targets are covered; unsupported Markdig targets remain literal", "Supported owners retain dedicated attribute tokens, semantic associations, normalized/original source mappings, and explicit unavailable reasons", "Supported owners project attributes to HTML, write Markdown, and reparse to the same semantic owner", "Focused Markdig-boundary and OfficeIMO container source-edit/renderer/writer proof is present", row.NextAction);
            case "UseGridTables":
                return new MatrixCells("Optional block parser", "Missing grid-table parser", "Missing table source model", "Missing HTML/Markdown writer behavior", "Missing malformed fallback fixtures", row.NextAction);
            case "UseListExtras":
                return new MatrixCells("Core opt-in parser", "Alpha and roman ordered markers are parsed for the scoped Markdig syntax, including nested lower-roman lists after parent text and inside blockquotes", "Uses canonical OrderedListBlock/ListItem marker style, delimiter, marker text, syntax marker spans, and nested/container listMarker source edits; needs broader edge breadth", "HTML type/start and parsed-marker Markdown writing are covered for scoped cases", "Has nested lower-roman Markdig comparison, blockquote nested-list comparison, and blockquote/nested-container source-edit reparse proof; needs remaining breadth before promotion", row.NextAction);
            case "UseMathematics":
                return new MatrixCells("Optional parser plus renderer policy", "Missing delimiter parity", "Needs math node/source metadata", "Needs renderer handoff and writer policy", "Needs inline/block math fixtures", row.NextAction);
            case "UseMediaLinks":
                return new MatrixCells("Renderer policy plus optional parser", "Missing shortcut parser", "Needs source metadata for providers", "Needs safe renderer output policy", "Needs provider comparison fixtures", row.NextAction);
            case "UsePreciseSourceLocation":
                return new MatrixCells("Cross-cutting source architecture", "Complete for documented source-backed fields; deliberately unavailable elsewhere", "Documented fields expose normalized spans, exact or line-ending-equivalent original mappings, stable semantic associations, source slices, snapshots, and source edits. Generated and transformed nodes expose machine-readable unavailable reasons instead of invented spans", "Unchanged trivia-backed input and validated field edits preserve original text; arbitrary semantic mutation uses canonical Markdown with explicit fallback diagnostics", "Covered by source-field, delimiter, trivia, line-ending, tab, generated-node, transform, native projection, source-edit, and reparse fixtures within the bounded contract", row.NextAction);
            case "UseReferralLinks":
                return new MatrixCells("Renderer policy", "Not parser-owned", "Needs link metadata decision", "Missing opt-in rel policy", "Needs renderer-policy tests", row.NextAction);
            case "UseSmartyPants":
                return new MatrixCells("Optional inline transform", "Missing smart punctuation transform", "Needs source/edit behavior", "Needs writer/escaping policy", "Needs opt-in transform fixtures", row.NextAction);
            case "UseCitations":
            case "UseFooters":
            case "UseGlobalization":
            case "UsePragmaLines":
                return new MatrixCells(row.ScopeDecision.ToDisplayText(), "Deferred", "Deferred", "Deferred", "Needs real consumer requirement", row.NextAction);
            case "UseJiraLinks":
                return new MatrixCells("Optional link extension", "Missing issue-key parser", "Needs source metadata", "Needs resolver/render policy", "Needs opt-in fixtures", row.NextAction);
            default:
                return CreateFallbackCells(row);
        }
    }

    private static MatrixCells CreateFallbackCells(MarkdigExtensionInventoryRow row) {
        string decision = row.ScopeDecision.ToDisplayText();

        return row.ScopeDecision switch {
            MarkdigExtensionScopeDecision.CoreEngine => new MatrixCells(decision, "Needed", "Needed", "Needed", "Needed", row.NextAction),
            MarkdigExtensionScopeDecision.OptionalExtension => new MatrixCells(decision, "Optional/missing", "Needed if enabled", "Needed if enabled", "Needed if enabled", row.NextAction),
            MarkdigExtensionScopeDecision.RendererHostPolicy => new MatrixCells(decision, "Not parser-owned until scoped", "Needs metadata decision", "Needed", "Needed", row.NextAction),
            MarkdigExtensionScopeDecision.Deferred => new MatrixCells(decision, "Deferred", "Deferred", "Deferred", "Needs real consumer requirement", row.NextAction),
            MarkdigExtensionScopeDecision.IntentionalDifference => new MatrixCells(decision, "Not planned", "Not planned", "Intentional difference", "Documented", row.NextAction),
            _ => new MatrixCells(decision, "Unknown", "Unknown", "Unknown", "Unknown", row.NextAction)
        };
    }

    private static string EscapeTable(string value) => value.Replace("|", "\\|");

    private sealed class MatrixCells {
        public MatrixCells(string decision, string engineParser, string astSource, string writerRender, string proof, string nextAction) {
            Decision = decision;
            EngineParser = engineParser;
            AstSource = astSource;
            WriterRender = writerRender;
            Proof = proof;
            NextAction = nextAction;
        }

        public string Decision { get; }
        public string EngineParser { get; }
        public string AstSource { get; }
        public string WriterRender { get; }
        public string Proof { get; }
        public string NextAction { get; }
    }
}
