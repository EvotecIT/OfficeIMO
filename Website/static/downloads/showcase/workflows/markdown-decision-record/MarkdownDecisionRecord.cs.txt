using System.IO;
using OfficeIMO.Markdown;

namespace OfficeIMO.Examples.Showcase.Workflows;

/// <summary>Builds a repository-friendly architecture decision record with context and consequences.</summary>
internal static class MarkdownDecisionRecord {
    internal static void Create(string folder) {
        MarkdownDoc document = MarkdownDoc.Create()
            .H1("ADR-007: Keep report inputs with the output")
            .P("Status: Accepted | Owner: Reporting team | Date: 14 September 2026")
            .H2("Context")
            .P("A generated PDF is easy to share, but reviewers also need to understand which data and configuration produced it.")
            .H2("Decision")
            .P("Deliver the report, its structured input snapshot, and a small manifest as one versioned bundle.")
            .Code("text", "report-bundle/\n  report.pdf\n  input.json\n  manifest.json")
            .H2("Alternatives considered")
            .Table(table => table.Headers("Alternative", "Reason not selected")
                .Row("PDF only", "Insufficient context for reproducing the report")
                .Row("Live dashboard link only", "The underlying data can change after review")
                .Row("Full database export", "Unnecessary volume and unrelated information"))
            .H2("Consequences")
            .Ul(list => list.Item("The input snapshot must contain only the fields used by the report.")
                .Item("The manifest records the generator version and file hashes.")
                .Item("Retention rules apply to the whole bundle."))
            .H2("Revisit when")
            .P("The report contains data that cannot be retained, or the generation volume makes the bundle impractical.");
        File.WriteAllText(Path.Combine(folder, "example.md"), document.ToMarkdown());
    }
}
