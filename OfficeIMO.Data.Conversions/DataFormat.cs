namespace OfficeIMO.Data.Conversions;

/// <summary>The formats the hub reads and writes. Excel is added by OfficeIMO.Data.Conversions.Excel.</summary>
internal enum DataFormat {
    Json,
    JsonLines,
    Yaml,
    Xml,
    Csv,
    Tsv,
    Toml
}

internal static class DataFormats {
    public static string Name(DataFormat format) => format switch {
        DataFormat.Json => "JSON",
        DataFormat.JsonLines => "JSON Lines",
        DataFormat.Yaml => "YAML",
        DataFormat.Xml => "XML",
        DataFormat.Csv => "CSV",
        DataFormat.Tsv => "TSV",
        DataFormat.Toml => "TOML",
        _ => format.ToString()
    };

    public static DataFormat? FromExtension(string? path) => Path.GetExtension(path)?.ToLowerInvariant() switch {
        ".json" => DataFormat.Json,
        ".jsonl" or ".ndjson" => DataFormat.JsonLines,
        ".yaml" or ".yml" => DataFormat.Yaml,
        ".xml" => DataFormat.Xml,
        ".csv" => DataFormat.Csv,
        ".tsv" or ".tab" => DataFormat.Tsv,
        ".toml" => DataFormat.Toml,
        _ => null
    };

    public static bool IsTable(DataFormat format) => format is DataFormat.Csv or DataFormat.Tsv;
}
