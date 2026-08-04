namespace OfficeIMO.Tests.MarkdownSuite;

internal static class MarkdigExtensionInventoryMarkdownWriter {
    internal const string PartialBoundariesStart = "<!-- extension-partial-boundaries:start -->";
    internal const string PartialBoundariesEnd = "<!-- extension-partial-boundaries:end -->";

    public static string Write(MarkdigExtensionInventoryReport report) {
        var sb = new StringBuilder();

        sb.AppendLine("# OfficeIMO.Markdown Markdig Extension Inventory");
        sb.AppendLine();
        sb.AppendLine($"This report compares the Markdig `{report.MarkdigVersion}` extension-family entry points reflected from the local comparison package with the current `OfficeIMO.Markdown` support story.");
        sb.AppendLine();
        sb.AppendLine("Status values:");
        sb.AppendLine();
        sb.AppendLine("- `Covered`: implemented and protected by focused evidence.");
        sb.AppendLine("- `Partial`: real OfficeIMO support exists, but Markdig breadth, options, source mapping, writer behavior, or renderer behavior is incomplete.");
        sb.AppendLine("- `Intentional`: the Markdig entry point is a bundle, helper, or renderer policy that OfficeIMO should model differently.");
        sb.AppendLine("- `Gap`: no meaningful OfficeIMO equivalent exists yet.");
        sb.AppendLine();
        sb.AppendLine("Route values name the candidate owner for an implementation. Scope decisions collapse those routes into execution buckets, so missing behavior belongs in the reusable engine, optional extension, renderer/host policy, deferred backlog, or intentionally documented difference instead of drifting into ad hoc tests.");
        sb.AppendLine();
        sb.AppendLine("Refresh command:");
        sb.AppendLine();
        sb.AppendLine("```powershell");
        sb.AppendLine("$env:OFFICEIMO_UPDATE_MARKDIG_INVENTORY = '1'");
        sb.AppendLine("dotnet test OfficeIMO.Markdown.Tests\\OfficeIMO.Markdown.Tests.csproj --framework net8.0 --filter \"FullyQualifiedName~Markdown_Markdig_Extension_Inventory_Tests\"");
        sb.AppendLine("Remove-Item Env:\\OFFICEIMO_UPDATE_MARKDIG_INVENTORY");
        sb.AppendLine("```");
        sb.AppendLine();
        sb.AppendLine("## Summary");
        sb.AppendLine();
        sb.AppendLine("| Metric | Count |");
        sb.AppendLine("| --- | ---: |");
        sb.AppendLine($"| Markdig extension-family rows | {report.Total} |");
        sb.AppendLine($"| Covered | {report.Covered} |");
        sb.AppendLine($"| Partial | {report.Partial} |");
        sb.AppendLine($"| Intentional | {report.Intentional} |");
        sb.AppendLine($"| Gap | {report.Gap} |");
        sb.AppendLine();
        sb.AppendLine("## Extension Families");
        sb.AppendLine();
        sb.AppendLine("| Markdig entry point | Family | Status | Scope decision | Route | Promotion bar | OfficeIMO state | Next action |");
        sb.AppendLine("| --- | --- | --- | --- | --- | --- | --- | --- |");

        foreach (var row in report.Rows) {
            sb.AppendLine($"| `{row.MethodName}` | {EscapeTable(row.Family)} | `{row.Status}` | {EscapeTable(row.ScopeDecision.ToDisplayText())} | {EscapeTable(row.Route)} | {EscapeTable(row.PromotionBar)} | {EscapeTable(row.OfficeImoState)} | {EscapeTable(row.NextAction)} |");
        }

        sb.AppendLine();
        sb.AppendLine("## Reflected Pipeline Entry Points");
        sb.AppendLine();
        sb.AppendLine("These public Markdig pipeline-builder methods are reflected from the local package so package upgrades cannot silently add a new `Use*` extension family without updating this report.");
        sb.AppendLine();
        sb.AppendLine("| Method | Tracked as extension family |");
        sb.AppendLine("| --- | --- |");

        var tracked = report.Rows.Select(static row => row.MethodName).ToHashSet(StringComparer.Ordinal);
        foreach (string methodName in report.ReflectedMethodNames) {
            sb.AppendLine($"| `{methodName}` | {(tracked.Contains(methodName) ? "Yes" : "No")} |");
        }

        sb.AppendLine();
        sb.AppendLine("## Next Use");
        sb.AppendLine();
        sb.AppendLine("- Use this inventory to decide whether an upcoming slice is parser grammar, AST/source mapping, renderer/writer behavior, extension seam work, or an intentional profile difference.");
        sb.AppendLine("- Keep `Partial` rows honest: promote them to `Covered` only when parser, AST/source, renderer, writer, and fixture evidence all match the claimed scope.");
        sb.AppendLine("- Use the `Scope decision`, `Route`, and `Promotion bar` columns before implementation so every slice moves the right owner instead of creating another local workaround.");
        sb.AppendLine("- Add fixtures or engine work by row, not by nearby test names.");

        return sb.ToString().Replace("\r\n", "\n");
    }

    public static string WritePublishedPartialBoundaries(MarkdigExtensionInventoryReport report) {
        var sb = new StringBuilder();
        sb.AppendLine(PartialBoundariesStart);
        sb.AppendLine("### Partial-family boundaries");
        sb.AppendLine();
        sb.AppendLine("These are the exact current implementation boundaries and promotion requirements for every `Partial` family in the structured extension inventory.");

        foreach (MarkdigExtensionInventoryRow row in report.Rows.Where(static row =>
                     row.Status == MarkdigExtensionInventoryStatus.Partial)) {
            sb.AppendLine();
            sb.Append("#### ").AppendLine(EscapePublishedText(row.Family));
            sb.AppendLine();
            sb.Append("- **OfficeIMO state:** ").AppendLine(EscapePublishedText(GetPublishedOfficeImoState(row)));
            sb.Append("- **Promotion bar:** ").AppendLine(EscapePublishedText(GetPublishedPromotionBar(row)));
        }

        sb.AppendLine(PartialBoundariesEnd);
        return sb.ToString().Replace("\r\n", "\n").TrimEnd();
    }

    public static string GetPublishedRoute(MarkdigExtensionInventoryRow row) =>
        row.Status == MarkdigExtensionInventoryStatus.Gap
            ? "Unavailable; candidate owner: " + row.Route
            : row.Route;

    private static string GetPublishedOfficeImoState(MarkdigExtensionInventoryRow row) =>
        row.Family switch {
            "Custom containers" => "Opt-in colon-fenced containers have complete ownership for the supported root, nested, blockquote-contained, and list-contained shapes: child parsing, HTML rendering, Markdown writing, syntax/native fields, source slices, source edits, and stable reparse. This row remains partial only relative to the broader optional extension surface.",
            "Diagrams" => "Semantic fenced blocks and visual renderer hooks exist; named diagram-language mapping and a complete renderer handoff contract remain open.",
            "Figures" => "Image and figure import plus publisher rendering paths exist; a dedicated Markdown figure syntax and its source/writer contract remain open.",
            "Generic attributes" => "Generic attributes have end-to-end ownership for every supported target family, including callouts, details blocks, and custom containers: semantic and syntax storage, exact source fields, HTML projection, Markdown writing, source edits, and stable reparse. Targets outside that declared set remain literal or deliberately consumed according to the documented profile boundary.",
            "List extras" => "Opt-in alphabetic and Roman ordered markers support nested parsing, marker-style HTML, source metadata and edits, and Markdown writer preservation. Remaining edge, source-edit, and reparse coverage keeps this family partial.",
            "Mathematics" => "Math-oriented semantic and rendering hooks exist, but inline and block delimiter parsing does not yet have a complete AST, source, writer, and renderer contract.",
            "Media links" => "Image and media semantics exist, but shortcut media providers do not yet have a complete parser, safe-renderer, source, and writer contract.",
            "Precise source location" => "The public source contract is complete and field-bounded: documented source-backed fields expose normalized spans, exact or line-ending-equivalent original mappings, stable semantic associations, and source edits; generated or transformed nodes are spanless with machine-readable unavailable reasons; arbitrary semantic edits use normalized writing. This row remains partial only because arbitrary-node locations and lossless arbitrary semantic edits are intentionally outside the contract.",
            _ => throw new InvalidOperationException("Published partial-family text is missing for " + row.Family + ".")
        };

    private static string GetPublishedPromotionBar(MarkdigExtensionInventoryRow row) =>
        row.Family switch {
            "Custom containers" => "Adopt any additional optional container shape only with parser, semantic owner, source fields, HTML output, Markdown writing, and reparse proof in the same change.",
            "Diagrams" => "Define named diagram-language mapping, renderer-package ownership, source/writer behavior, and focused fixtures.",
            "Figures" => "Separate HTML-import figure recovery from authored Markdown figure syntax, then prove renderer, writer, and source behavior.",
            "Generic attributes" => "Adopt any additional target family only when it has one semantic owner plus source mapping, HTML projection or deliberate consumption, Markdown writing, source editing, and stable reparse proof.",
            "List extras" => "Broaden remaining list-marker edges, native source edits, and writer reparse proof.",
            "Mathematics" => "Define inline and block delimiters, AST/source/native metadata, writer preservation, and renderer handoff.",
            "Media links" => "Define the provider model, safe renderer output, writer preservation, and source metadata for shortcut media links.",
            "Precise source location" => "Do not promote this comparison row unless OfficeIMO deliberately expands beyond its field-bounded contract; never infer spans for generated nodes or advertise arbitrary semantic edits as lossless.",
            _ => throw new InvalidOperationException("Published promotion text is missing for " + row.Family + ".")
        };

    private static string EscapeTable(string value) => value.Replace("|", "\\|");

    private static string EscapePublishedText(string value) => value
        .Replace("<", "&lt;")
        .Replace(">", "&gt;");
}
