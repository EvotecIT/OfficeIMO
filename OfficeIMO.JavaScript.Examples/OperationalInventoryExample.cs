using System.Globalization;
using System.Text;
using System.Text.Json;
using HtmlForgeX;
using OfficeIMO.JavaScript.Assets;

/// <summary>Produces an actual offline HtmlForgeX inventory and independent C# selection expectations.</summary>
internal static class OperationalInventoryExample {
    private static readonly CultureInfo Invariant = CultureInfo.InvariantCulture;
    private static readonly JsonSerializerOptions Json = new() { WriteIndented = true };
    private static readonly Field[] Fields = [
        new("service", "Service", "string", 20), new("site", "Site", "string", 14),
        new("status", "Status", "string", 14), new("latency", "Latency (ms)", "number", 16, "0.0"),
        new("availability", "Availability", "number", 16, "0.00%"), new("checked", "Checked (UTC)", "date", 23, "yyyy-mm-dd hh:mm"),
        new("requests", "Requests / min", "number", 19, "#,##0"), new("errorRate", "Error rate", "number", 16, "0.00%"),
        new("cpu", "CPU usage", "number", 16, "0.0%"), new("memory", "Memory (MiB)", "number", 18, "#,##0"),
        new("storage", "Storage (GiB)", "number", 18, "#,##0.0"), new("certificateExpiry", "Certificate expiry (UTC)", "date", 25, "yyyy-mm-dd"),
        new("version", "Version", "string", 20), new("lastIncident", "Last incident (UTC)", "date", 24, "yyyy-mm-dd hh:mm"),
        new("ticket", "Ticket", "string", 18), new("supportWindow", "Support window", "string", 24)
    ];

    /// <summary>Writes a generated page, source model and expected captures without executing an export.</summary>
    internal static void Write(string output) {
        Directory.CreateDirectory(output);
        InventoryRow[] rows = Enumerable.Range(0, 100).Select(CreateRow).ToArray();
        using var document = new Document { LibraryMode = LibraryMode.Offline };
        document.Head.AddTitle("Operational inventory exports");
        document.Head.AddDefaultStyles();
        document.Body.Add(new HtmlTag("h1").Value("Operational inventory exports"));
        document.Body.Add(new HtmlTag("p").Value("Select services, filter the inventory or change its columns, then export the current view. Generated with HtmlForgeX; documents run locally in a worker."));
        document.Body.Add(new HtmlTag("label").Attribute("for", "inventory-filter").Value("Find service or site "));
        document.Body.Add(new HtmlTag("input").Attribute("type", "search").Id("inventory-filter"));
        var controls = new TablerButtonBar();
        foreach (var button in new[] { ("inventory-sort", "Reverse service order"), ("inventory-hide-site", "Hide/show Site"),
            ("inventory-move-latency", "Move Latency first/last"), ("inventory-select", "Select visible services"),
            ("inventory-export-xlsx", "Export Excel"), ("inventory-export-pdf", "Export PDF") })
            controls.Add(new TablerButton(button.Item2).Type("button").Attribute("id", button.Item1));
        document.Body.Add(controls);
        var table = new Table().AddHeaders(["Select", .. Fields.Select(field => field.Header)]);
        foreach (InventoryRow row in rows) {
            var cells = new List<TableCell> { TableCell.FromElement(new HtmlTag("input").Attribute("type", "checkbox")
                .Attribute("checked", "checked").Attribute("aria-label", "Select " + row.Id).Attribute("data-select-service", row.Id)) };
            foreach (Field field in Fields) {
                object? value = row.Values[field.Key];
                var span = new HtmlTag("span").Attribute("data-field", field.Key).Attribute("data-value-type", value is null ? "null" : field.Type)
                    .Attribute("data-raw-value", Raw(value)).Value(Display(value, field));
                if (field.Key == "supportWindow" && value is string note && note.Contains('\n'))
                    span.Class("d-block").Style("width", "20rem").Style("max-height", "4.5rem")
                        .Style("overflow-y", "auto").Style("white-space", "pre-wrap")
                        .Attribute("tabindex", "0").Attribute("role", "region")
                        .Attribute("aria-label", "Full support note for " + row.Id);
                if (field.Key == "service") span.Attribute("data-row-id", row.Id).Attribute("data-portfolio", row.Path[0])
                    .Attribute("data-environment", row.Path[1]).Attribute("data-owner", row.Path[2]);
                cells.Add(TableCell.FromElement(span));
            }
            table.AddRow(cells);
        }
        // Table owns the headers and cells; the wrappers attach identity and existing Tabler table styles.
        document.Body.Add(new HtmlTag("div").Class("table-responsive").Value(new HtmlTag("table")
            .Class("table table-nowrap align-top").Style("min-width", "100rem").Id("operational-inventory")
            .Attribute("data-column-fields", string.Join(' ', Fields.Select(field => field.Key)))
            .Attribute("data-column-schema", JsonSerializer.Serialize(Fields.Select(field => new {
                key = field.Key, type = field.Type, width = field.Width, format = field.Format, rowHeader = field.Key == "service"
            }))).Value(table)));
        document.Body.Add(new HtmlTag("p").Id("inventory-status").Attribute("role", "status").Attribute("aria-live", "polite").Value("Ready"));
        string Read(string name) => File.ReadAllText(Path.Combine(AppContext.BaseDirectory, name));
        var assets = new { classicScript = JavaScriptAssets.Script.Content, executionScript = JavaScriptAssets.ExecutionScript.Content,
            workerScript = Read("operational-inventory-worker.js"),
            fonts = new { regular = Convert.ToBase64String(File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Carlito-Regular.ttf"))),
                bold = Convert.ToBase64String(File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Carlito-Bold.ttf"))) } };
        document.Head.AddJsInline(TrustedJavaScript.FromTrustedSource(JavaScriptAssets.Script.Content));
        document.Head.AddJsInline(TrustedJavaScript.FromTrustedSource(JavaScriptAssets.ExecutionScript.Content));
        document.Head.AddJsInline(TrustedJavaScript.FromTrustedSource("globalThis.operationalInventoryAssets = " + JsonSerializer.Serialize(assets) + ";"));
        document.Body.Add(new InlineScript(TrustedJavaScript.FromTrustedSource(Read("operational-inventory.js"))));
        document.Save(Path.Combine(output, "operational-inventory.html"), openInBrowser: false);
        string[] selectedKeys = ["latency", .. Fields.Where(field => field.Key is not ("latency" or "site")).Select(field => field.Key)];
        object Expected(string name, IEnumerable<InventoryRow> selected, string[] keys) => new {
            name, keys, headers = keys.Select(key => Fields.Single(field => field.Key == key).Header).ToArray(),
            sourceIdColumn = "service", rowHeaderColumns = new[] { new { key = "service", columnIndex = Array.IndexOf(keys, "service") } },
            rows = selected.Select(row => new { id = row.Id, values = keys.Select(key => row.Values[key]).ToArray(), path = row.Path }).ToArray()
        };
        var metadata = new { producer = "HtmlForgeX 2.0.0 typed Document/Table", schemaVersion = 1,
            groupKeys = new[] { "portfolio", "environment", "owner" },
            columns = Fields.Select(field => new { key = field.Key, header = field.Header, type = field.Type, width = field.Width, format = field.Format, rowHeader = field.Key == "service" }).ToArray(),
            source = rows, cases = new[] { Expected("all", rows, Fields.Select(field => field.Key).ToArray()),
                Expected("selected", rows.Where(row => Equals(row.Values["site"], "Warsaw") && row.Id != "OPS-0001").Reverse(), selectedKeys),
                Expected("empty", [], selectedKeys) } };
        File.WriteAllText(Path.Combine(output, "operational-inventory.expected.json"), JsonSerializer.Serialize(metadata, Json), new UTF8Encoding(false));
        Console.WriteLine("Wrote actual HtmlForgeX operational inventory and independent selection metadata to " + output);
    }

    private static InventoryRow CreateRow(int index) {
        var time = new DateTimeOffset(2026, 10, 22, 8, 0, 0, TimeSpan.Zero).AddMinutes(index);
        string id = "OPS-" + (index + 1).ToString("D4", Invariant);
        return new InventoryRow(id, new Dictionary<string, object?> {
            ["service"] = id, ["site"] = new[] { "Warsaw", "Łódź", "Praha", "Berlin" }[index / 5 % 4],
            ["status"] = index % 17 == 0 ? "Offline" : index % 7 == 0 ? "Attention" : "Healthy",
            ["latency"] = index % 19 == 0 ? null : 4.75 + index / 10d, ["availability"] = index % 23 == 0 ? null : .995 - index % 17 / 10000d,
            ["checked"] = time, ["requests"] = 1500 + index * 11, ["errorRate"] = index % 11 / 1000d,
            ["cpu"] = .15 + index % 35 / 100d, ["memory"] = 512 + index * 16, ["storage"] = 20.5 + index * 1.25,
            ["certificateExpiry"] = time.AddDays(90), ["version"] = "3.1." + index % 13,
            ["lastIncident"] = index % 9 == 0 ? null : time.AddDays(-index % 20), ["ticket"] = index % 7 == 0 ? "INC-" + (2000 + index) : "",
            ["supportWindow"] = index == 84 ? string.Join('\n', Enumerable.Range(0, 36).Select(line => "OPSNOTE084-" + line.ToString("D3", Invariant)))
                : index % 4 == 0 ? "24 × 7" : "Business hours"
        }, ["Portfolio " + index / 50, index / 25 % 2 == 0 ? "Production" : "Staging", "Owner " + index / 5 % 5]);
    }

    private static string Raw(object? value) => value switch {
        null => "", DateTimeOffset time => time.ToString("yyyy-MM-dd'T'HH:mm:ss.fff'Z'", Invariant),
        IFormattable number => number.ToString(null, Invariant), _ => value.ToString()!
    };
    private static string Display(object? value, Field field) => value is null ? "—" : value is DateTimeOffset time ? time.ToString("yyyy-MM-dd HH:mm", Invariant)
        : field.Key is "availability" or "errorRate" or "cpu" ? ((double)value).ToString("P2", Invariant) : Raw(value);
    private sealed record Field(string Key, string Header, string Type, int Width, string? Format = null);
    private sealed record InventoryRow(string Id, Dictionary<string, object?> Values, string[] Path);
}
