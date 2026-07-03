namespace OfficeIMO.Tests.MarkdownSuite;

internal static class GfmInventoryMarkdownWriter {
    public static string Write(GfmInventoryReport report) {
        var sb = new StringBuilder();

        sb.AppendLine("# OfficeIMO.Markdown GFM Inventory");
        sb.AppendLine();
        sb.AppendLine("This report is generated from the checked-in cmark-gfm extension smoke fixtures and the current `OfficeIMO.Markdown` GitHub Flavored Markdown profile.");
        sb.AppendLine();
        sb.AppendLine("Refresh command:");
        sb.AppendLine();
        sb.AppendLine("```powershell");
        sb.AppendLine("$env:OFFICEIMO_UPDATE_GFM_INVENTORY = '1'");
        sb.AppendLine("dotnet test OfficeIMO.Markdown.Tests\\OfficeIMO.Markdown.Tests.csproj --framework net8.0 --filter \"FullyQualifiedName~Markdown_GitHubFlavoredMarkdown_Inventory_Tests\"");
        sb.AppendLine("Remove-Item Env:\\OFFICEIMO_UPDATE_GFM_INVENTORY");
        sb.AppendLine("```");
        sb.AppendLine();
        sb.AppendLine("## Summary");
        sb.AppendLine();
        sb.AppendLine("| Metric | Count |");
        sb.AppendLine("| --- | ---: |");
        sb.AppendLine($"| Tracked fixtures | {report.Total} |");
        sb.AppendLine($"| Upstream cmark-gfm fixtures | {report.UpstreamTracked} |");
        sb.AppendLine($"| OfficeIMO supplement fixtures | {report.Supplements} |");
        sb.AppendLine($"| Passing fixtures | {report.Passing} |");
        sb.AppendLine($"| Failing fixtures | {report.Failing} |");
        sb.AppendLine($"| Intentional deviations | {report.IntentionalDeviations} |");
        sb.AppendLine();
        sb.AppendLine("## Section Inventory");
        sb.AppendLine();
        sb.AppendLine("| Section | Tracked | Upstream | Supplements | Passing | Failing | Intentional |");
        sb.AppendLine("| --- | ---: | ---: | ---: | ---: | ---: | ---: |");

        foreach (var section in report.EnumerateSectionSummaries()) {
            sb.AppendLine($"| {EscapeTable(section.Section)} | {section.Total} | {section.Upstream} | {section.Supplements} | {section.Passing} | {section.Failing} | {section.IntentionalDeviations} |");
        }

        sb.AppendLine();
        sb.AppendLine("## Source Inventory");
        sb.AppendLine();
        sb.AppendLine("| Source | Tracked | Passing | Failing |");
        sb.AppendLine("| --- | ---: | ---: | ---: |");

        foreach (var source in report.EnumerateSourceSummaries()) {
            sb.AppendLine($"| {EscapeTable(source.Source)} | {source.Total} | {source.Passing} | {source.Failing} |");
        }

        sb.AppendLine();
        sb.AppendLine("## Failure Clusters");
        sb.AppendLine();
        sb.AppendLine("| Cluster | Failing | Sections | First fixture indexes |");
        sb.AppendLine("| --- | ---: | --- | --- |");

        foreach (var cluster in report.EnumerateFailureClusters()) {
            sb.AppendLine($"| {EscapeTable(cluster.Cluster)} | {cluster.Count} | {EscapeTable(cluster.Sections)} | {string.Join(", ", cluster.Indexes.Select(static index => "#" + index.ToString(System.Globalization.CultureInfo.InvariantCulture)))} |");
        }

        sb.AppendLine();
        sb.AppendLine("## Next Use");
        sb.AppendLine();
        sb.AppendLine("- Use the section inventory to pick GFM expansion work by enabled extension family.");
        sb.AppendLine("- Keep upstream cmark-gfm fixtures and OfficeIMO supplement fixtures separated when adding new cases.");
        sb.AppendLine("- When a GFM parser or renderer slice lands, refresh this report and promote new upstream examples only after the behavior contract is understood.");

        return sb.ToString().Replace("\r\n", "\n");
    }

    private static string EscapeTable(string value) => value.Replace("|", "\\|");
}
