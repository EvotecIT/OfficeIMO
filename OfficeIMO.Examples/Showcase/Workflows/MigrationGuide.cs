using OfficeIMO.Markdown;

namespace OfficeIMO.Examples.Showcase.Workflows;

/// <summary>Produces a portable migration guide with explicit old-to-new mappings and a verification sequence.</summary>
internal static class MigrationGuide {
    internal static void Create(string folder) {
        var document = MarkdownDoc.Create()
            .H1("Configuration migration guide")
            .P("Move an illustrative report job from separate output settings to a named delivery block.")
            .P("Before you start, back up the configuration and use a temporary output folder for the first run.")
            .H2("Map the settings")
            .Table(table => table.Headers("Version 1", "Version 2", "Meaning")
                .Row("outputPath", "delivery.folder", "Destination for generated files")
                .Row("filePrefix", "delivery.prefix", "Prefix applied to each report")
                .Row("overwrite", "delivery.conflict", "Use replace or fail explicitly"))
            .H2("Before")
            .Code("json", "{\n  \"outputPath\": \"./reports\",\n  \"filePrefix\": \"weekly-\",\n  \"overwrite\": false\n}")
            .H2("After")
            .Code("json", "{\n  \"delivery\": {\n    \"folder\": \"./reports\",\n    \"prefix\": \"weekly-\",\n    \"conflict\": \"fail\"\n  }\n}")
            .H2("Verify the migration")
            .Ul(list => list.Item("Run once with a small known input.")
                .Item("Compare filenames, row counts and totals with the previous job.")
                .Item("Repeat the run to confirm the chosen conflict behavior.")
                .Item("Keep the original configuration until the scheduled run succeeds."));
        File.WriteAllText(Path.Combine(folder, "example.md"), document.ToMarkdown());
    }
}
