namespace OfficeIMO.Tests.MarkdownSuite;

internal static class CommonMarkInventoryMarkdownWriter {
    public static string Write(CommonMarkInventoryReport report) {
        var sb = new StringBuilder();

        sb.AppendLine("# OfficeIMO.Markdown CommonMark Inventory");
        sb.AppendLine();
        sb.AppendLine("This report is generated from the checked-in official CommonMark `0.31.2` spec JSON and the current `OfficeIMO.Markdown` CommonMark profile.");
        sb.AppendLine();
        sb.AppendLine("Refresh command:");
        sb.AppendLine();
        sb.AppendLine("```powershell");
        sb.AppendLine("$env:OFFICEIMO_UPDATE_COMMONMARK_INVENTORY = '1'");
        sb.AppendLine("dotnet test OfficeIMO.Markdown.Tests\\OfficeIMO.Markdown.Tests.csproj --framework net8.0 --filter \"FullyQualifiedName~Markdown_CommonMark_Inventory_Tests\"");
        sb.AppendLine("Remove-Item Env:\\OFFICEIMO_UPDATE_COMMONMARK_INVENTORY");
        sb.AppendLine("```");
        sb.AppendLine();
        sb.AppendLine("## Summary");
        sb.AppendLine();
        sb.AppendLine("| Metric | Count |");
        sb.AppendLine("| --- | ---: |");
        sb.AppendLine($"| Official examples | {report.Total} |");
        sb.AppendLine($"| Pinned smoke fixtures | {report.Pinned} |");
        sb.AppendLine($"| Passing pinned fixtures | {report.PassingPinned} |");
        sb.AppendLine($"| Passing unpinned examples | {report.PassingUnpinned} |");
        sb.AppendLine($"| Failing examples | {report.Failing} |");
        sb.AppendLine($"| Intentional deviations | {report.IntentionalDeviations} |");
        sb.AppendLine();
        sb.AppendLine("## Section Inventory");
        sb.AppendLine();
        sb.AppendLine("| Section | Official | Pinned | Passing pinned | Passing unpinned | Failing | Intentional |");
        sb.AppendLine("| --- | ---: | ---: | ---: | ---: | ---: | ---: |");

        foreach (var section in report.EnumerateSectionSummaries()) {
            sb.AppendLine($"| {EscapeTable(section.Section)} | {section.Total} | {section.Pinned} | {section.PassingPinned} | {section.PassingUnpinned} | {section.Failing} | {section.IntentionalDeviations} |");
        }

        sb.AppendLine();
        sb.AppendLine("## Failure Clusters");
        sb.AppendLine();
        sb.AppendLine("| Cluster | Failing | Sections | First examples |");
        sb.AppendLine("| --- | ---: | --- | --- |");

        foreach (var cluster in report.EnumerateFailureClusters()) {
            sb.AppendLine($"| {EscapeTable(cluster.Cluster)} | {cluster.Count} | {EscapeTable(cluster.Sections)} | {string.Join(", ", cluster.Examples.Select(static example => "#" + example.ToString(System.Globalization.CultureInfo.InvariantCulture)))} |");
        }

        sb.AppendLine();
        sb.AppendLine("## Next Use");
        sb.AppendLine();
        sb.AppendLine("- Use the failure clusters to pick parser work by root cause, not by nearby example number.");
        sb.AppendLine("- When a parser slice lands, refresh this report and promote newly passing examples into `commonmark-0.31.2-smoke.json` only when the engine contract is understood.");
        sb.AppendLine("- Keep intentional deviations at zero unless the compatibility matrix explains the profile difference.");

        return sb.ToString().Replace("\r\n", "\n");
    }

    private static string EscapeTable(string value) => value.Replace("|", "\\|");
}
