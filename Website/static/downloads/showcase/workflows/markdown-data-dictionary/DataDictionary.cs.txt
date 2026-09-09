using OfficeIMO.Markdown;

namespace OfficeIMO.Examples.Showcase.Workflows;

/// <summary>Turns a structured field catalog into a data dictionary with constraints and a sample record.</summary>
internal static class DataDictionary {
    internal static void Create(string folder) {
        var fields = new[] {
            new Field("request_id", "string", "Required; unique", "Stable identifier, for example SR-2048"),
            new Field("status", "string", "new | active | waiting | closed", "Current workflow state"),
            new Field("owner", "string", "Required while active", "Team responsible for the next action"),
            new Field("opened_at", "datetime", "ISO 8601 UTC", "When the request entered the service"),
            new Field("age_days", "integer", "Zero or greater", "Whole elapsed days at export time")
        };
        var document = MarkdownDoc.Create()
            .H1("Service request export dictionary")
            .P("Schema 1.0 / one record per request / UTF-8 JSON Lines")
            .H2("Fields")
            .Table(table => {
                table.Headers("Field", "Type", "Constraint", "Meaning");
                foreach (var field in fields) table.Row(field.Name, field.Type, field.Constraint, field.Meaning);
            })
            .H2("Sample record")
            .Code("json", "{\n  \"request_id\": \"SR-2048\",\n  \"status\": \"active\",\n  \"owner\": \"Support\",\n  \"opened_at\": \"2026-09-01T09:00:00Z\",\n  \"age_days\": 2\n}")
            .H2("Consumer rules")
            .Ul(list => list.Item("Treat request_id as an identifier, not a number.")
                .Item("Preserve unknown fields when forwarding records.")
                .Item("Reject negative age_days and report the source record.")
                .Item("Use opened_at for time-based analysis; age_days is a snapshot."))
            .Callout("note", "Example contract", "These fields describe a fictional service export. Replace the catalog with your application's actual schema.");
        File.WriteAllText(Path.Combine(folder, "example.md"), document.ToMarkdown());
    }

    private sealed record Field(string Name, string Type, string Constraint, string Meaning);
}
