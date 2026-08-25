using System.Text;

namespace OfficeIMO.Confluence.Benchmarks;

internal sealed record ConfluenceManagedSectionCorpus(string ExistingBody, string Replacement);

internal static class ConfluenceManagedSectionCorpusFactory {
    internal const string SectionId = "benchmark-report";

    internal static ConfluenceManagedSectionCorpus Create(int pageCharacters) {
        if (pageCharacters <= 0) throw new ArgumentOutOfRangeException(nameof(pageCharacters));
        string start = OfficeIMO.Confluence.ConfluenceManagedSection.StartMarker(SectionId);
        string end = OfficeIMO.Confluence.ConfluenceManagedSection.EndMarker(SectionId);
        string paragraph = "<p>Representative Confluence storage content with stable text and links.</p>\n";
        var prefix = new StringBuilder(pageCharacters / 2 + paragraph.Length);
        while (prefix.Length < pageCharacters / 2) prefix.Append(paragraph);
        var suffix = new StringBuilder(pageCharacters / 2 + paragraph.Length);
        while (suffix.Length < pageCharacters / 2) suffix.Append(paragraph);
        string existingBody = prefix + start + "\n<p>old managed content</p>\n" + end + "\n" + suffix;

        var replacement = new StringBuilder(Math.Max(1024, pageCharacters / 16));
        while (replacement.Length < Math.Max(1024, pageCharacters / 16)) {
            replacement.Append("<p>Updated managed report row with deterministic content.</p>\n");
        }
        return new ConfluenceManagedSectionCorpus(existingBody, replacement.ToString());
    }
}
