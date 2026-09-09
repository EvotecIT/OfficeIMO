using System.IO;
using OfficeIMO.Markdown;

namespace OfficeIMO.Examples.Showcase.Workflows;

/// <summary>Authors a portable incident runbook with structured steps, a decision table, and a warning.</summary>
internal static class IncidentRunbook {
    internal static void Create(string folder) {
        MarkdownDoc document = MarkdownDoc.Create()
            .H1("Service incident runbook")
            .P("Owner: Operations | Applies to: the Northwind request portal")
            .Callout("warning", "Before changing the service",
                "Confirm impact and name the incident owner. Record evidence before restarting components.")
            .H2("1. Establish the situation")
            .Ul(list => list.Item("Record the first observed failure and the affected user groups.")
                .Item("Check whether a deployment or configuration change preceded the failure.")
                .Item("Open an incident record and agree the next update time."))
            .H2("2. Choose the next check")
            .Table(table => table.Headers("Observation", "Next check", "Owner")
                .Row("All users affected", "Service health and dependencies", "Operations")
                .Row("One group affected", "Access and routing", "Support")
                .Row("Failure after a change", "Deployment evidence and rollback readiness", "Engineering"))
            .H2("3. Communicate")
            .Code("text", "Impact: <what users cannot do>\nOwner: <named coordinator>\nNext action: <specific check>\nNext update: <time>")
            .H2("4. Close with evidence")
            .P("Confirm recovery with an affected user, record the timeline, and assign follow-up actions.");
        File.WriteAllText(Path.Combine(folder, "example.md"), document.ToMarkdown());
    }
}
