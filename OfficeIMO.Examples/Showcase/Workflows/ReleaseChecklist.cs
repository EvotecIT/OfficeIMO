using OfficeIMO.Markdown;

namespace OfficeIMO.Examples.Showcase.Workflows {
    /// <summary>Creates a portable release-readiness checklist with explicit gates and a rollback record.</summary>
    internal static class ReleaseChecklist {
        internal static void Create(string folder) {
            var document = MarkdownDocument.Create()
                .H1("Release readiness checklist")
                .P("Northwind request portal / release candidate 1.4 / illustrative process")
                .H2("Verify the candidate")
                .Ul(list => list.ItemTask("Record the exact source and artifact identities")
                    .ItemTask("Run the supported-path smoke checks")
                    .ItemTask("Review migration and recovery steps")
                    .ItemTask("Confirm the support owner and release window"))
                .H2("Release gates")
                .Table(table => table.Headers("Gate", "Owner", "Evidence")
                    .Row("Qualification", "Engineering", "Current candidate checks")
                    .Row("Operations", "Service owner", "Recovery rehearsal")
                    .Row("Approval", "Release owner", "Recorded decision and conditions"))
                .H2("Rollback record")
                .Table(table => table.Headers("Record", "Value")
                    .Row("Previous artifact", "Record the verified identity")
                    .Row("Recovery trigger", "Agreed service-health threshold")
                    .Row("Recovery owner", "Name before the release window"))
                .H2("After the window")
                .Ul(list => list.ItemTask("Verify the deployed identity")
                    .ItemTask("Observe the agreed service checks")
                    .ItemTask("Record the outcome and remaining actions"));
            File.WriteAllText(Path.Combine(folder, "example.md"), document.ToMarkdown());
        }
    }
}
