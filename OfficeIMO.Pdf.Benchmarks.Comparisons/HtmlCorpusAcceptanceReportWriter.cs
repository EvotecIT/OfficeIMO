using System.Globalization;
using System.Text;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

internal static class HtmlCorpusAcceptanceReportWriter {
    internal static async Task WriteAsync(
        string path,
        HtmlCorpusAcceptanceEvidence acceptance,
        IReadOnlyList<HtmlCorpusCaseEvidence> evidence) {
        var builder = new StringBuilder();
        builder.AppendLine("# H4 advanced held-out visual acceptance");
        builder.AppendLine();
        builder.Append("Status: **").Append(acceptance.Passed ? "Passed" : "Failed").AppendLine("**  ");
        builder.Append("Acceptance configuration SHA-256: `").Append(acceptance.ConfigurationSha256).AppendLine("`");
        builder.AppendLine();
        builder.AppendLine(acceptance.ReferencePolicy);
        builder.AppendLine();
        builder.AppendLine("## Case results");
        builder.AppendLine();
        builder.AppendLine("| Case | Screen | Print | Screen-to-page |");
        builder.AppendLine("| --- | --- | --- | --- |");
        foreach (HtmlCorpusCaseAcceptance item in acceptance.Cases) {
            builder.Append("| `").Append(item.Id).Append("` | ")
                .Append(Status(item.Screen)).Append(" | ")
                .Append(Status(item.Print)).Append(" | ")
                .Append(Status(item.ScreenToPage)).AppendLine(" |");
        }
        builder.AppendLine();
        builder.AppendLine("## Capability results");
        builder.AppendLine();
        builder.AppendLine("| Profile | Capability | Intents | Selected cases | Status |");
        builder.AppendLine("| --- | --- | --- | ---: | --- |");
        foreach (HtmlCorpusCapabilityAcceptance item in acceptance.Capabilities) {
            builder.Append("| `").Append(item.ProfileId).Append("` | `").Append(item.CapabilityId).Append("` | ")
                .Append(string.Join(", ", item.Intents.Select(Escape))).Append(" | ")
                .Append(item.RequiredCaseIds.Count.ToString(CultureInfo.InvariantCulture)).Append(" | ")
                .Append(item.Passed ? "Passed" : "Failed: " + string.Join(", ", item.FailedCaseIds.Select(Escape)))
                .AppendLine(" |");
        }
        builder.AppendLine();
        builder.AppendLine("## Visual review artifacts");
        builder.AppendLine();
        foreach (HtmlCorpusCaseEvidence item in evidence) {
            builder.Append("### ").AppendLine(item.Id);
            builder.AppendLine();
            builder.Append("- Screen: [OfficeIMO](").Append(item.Id).Append('/').Append(item.OfficeImo?.ScreenPng.RelativePath)
                .Append("), [Chromium](").Append(item.Id).Append('/').Append(item.Browser?.ScreenPng.RelativePath)
                .Append("), [difference](").Append(item.Id).Append('/').Append(item.Comparisons.ScreenPixels?.DifferenceRelativePath).AppendLine(")");
            builder.Append("- Screen-to-page: [difference](").Append(item.Id).Append('/')
                .Append(item.Comparisons.ScreenToPage?.DifferenceRelativePath).AppendLine(")");
            HtmlCorpusPageComparison[] printPages = item.Comparisons.PrintPagesToChromium.ToArray();
            for (int index = 0; index < printPages.Length; index++) {
                HtmlCorpusPageComparison page = printPages[index];
                HtmlCorpusPageArtifact? officePage = item.OfficeImo?.PrintPdf.Pages.FirstOrDefault(value => value.PageNumber == page.PageNumber);
                HtmlCorpusPageArtifact? browserPage = item.Browser?.PrintPdf.Pages.FirstOrDefault(value => value.PageNumber == page.PageNumber);
                HtmlCorpusPageArtifact? peachPage = item.PeachPdf?.Pdf.Pages.FirstOrDefault(value => value.PageNumber == page.PageNumber);
                HtmlCorpusPageComparison? peachToBrowser = item.Comparisons.PeachPdfPrintPagesToChromium
                    .FirstOrDefault(value => value.PageNumber == page.PageNumber);
                builder.Append("- Print page ").Append(page.PageNumber.ToString(CultureInfo.InvariantCulture))
                    .Append(": [OfficeIMO](").Append(item.Id).Append('/').Append(officePage?.RelativePath)
                    .Append("), [Chromium](").Append(item.Id).Append('/').Append(browserPage?.RelativePath)
                    .Append("), [OfficeIMO difference](").Append(item.Id).Append('/').Append(page.Pixels?.DifferenceRelativePath).Append(')');
                if (peachPage != null && peachToBrowser?.Pixels?.DifferenceRelativePath != null) {
                    builder.Append(", [PeachPDF](").Append(item.Id).Append('/').Append(peachPage.RelativePath)
                        .Append("), [PeachPDF difference](").Append(item.Id).Append('/')
                        .Append(peachToBrowser.Pixels.DifferenceRelativePath).Append(')');
                }
                builder.AppendLine();
            }
            builder.AppendLine();
        }
        if (acceptance.Failures.Count > 0) {
            builder.AppendLine("## Failures");
            builder.AppendLine();
            foreach (string failure in acceptance.Failures) builder.Append("- ").AppendLine(failure);
        }
        await File.WriteAllTextAsync(path, builder.ToString(), new UTF8Encoding(false)).ConfigureAwait(false);
    }

    private static string Status(HtmlCorpusIntentAcceptance item) =>
        (item.Passed ? "Passed" : "Failed") + " (`" + item.Classification + "`)";

    private static string Escape(string value) => value.Replace("|", "\\|", StringComparison.Ordinal);
}
