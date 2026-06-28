namespace OfficeIMO.Tests.MarkdownSuite;

internal static class MarkdigExtensionInventoryMarkdownWriter {
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
        sb.AppendLine("Route values name the owning layer for future work, so missing behavior is fixed in the reusable engine, optional extension, renderer/host policy, or intentionally documented difference instead of drifting into ad hoc tests.");
        sb.AppendLine();
        sb.AppendLine("Refresh command:");
        sb.AppendLine();
        sb.AppendLine("```powershell");
        sb.AppendLine("$env:OFFICEIMO_UPDATE_MARKDIG_INVENTORY = '1'");
        sb.AppendLine("dotnet test OfficeIMO.Tests\\OfficeIMO.Tests.csproj --framework net8.0 --filter \"FullyQualifiedName~Markdown_Markdig_Extension_Inventory_Tests\"");
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
        sb.AppendLine("| Markdig entry point | Family | Status | Route | Promotion bar | OfficeIMO state | Next action |");
        sb.AppendLine("| --- | --- | --- | --- | --- | --- | --- |");

        foreach (var row in report.Rows) {
            sb.AppendLine($"| `{row.MethodName}` | {EscapeTable(row.Family)} | `{row.Status}` | {EscapeTable(row.Route)} | {EscapeTable(row.PromotionBar)} | {EscapeTable(row.OfficeImoState)} | {EscapeTable(row.NextAction)} |");
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
        sb.AppendLine("- Use the `Route` and `Promotion bar` columns before implementation so every slice moves the right owner instead of creating another local workaround.");
        sb.AppendLine("- Add fixtures or engine work by row, not by nearby test names.");

        return sb.ToString().Replace("\r\n", "\n");
    }

    private static string EscapeTable(string value) => value.Replace("|", "\\|");
}
