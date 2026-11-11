using OfficeIMO.Toml;

namespace OfficeIMO.Toml.Benchmarks;

/// <summary>
/// The same three operations in each library. Every operation returns a count of what it saw, which setup compares
/// across libraries before anything is timed.
/// </summary>
public static class TomlLibraries {
    // ---- Parse: one pass over the text, materializing keys and string values -------------------------------------

    /// <summary>OfficeIMO's pull reader; keys and strings are read as strings, as the reference parsers produce them.</summary>
    public static long ParseOfficeIMO(string toml) {
        var reader = new TomlReader(toml);
        long values = 0;
        while (reader.Read()) {
            switch (reader.TokenType) {
                case TomlTokenType.Key:
                    _ = reader.GetString();
                    break;
                case TomlTokenType.String:
                    _ = reader.GetString();
                    values++;
                    break;
                case TomlTokenType.Integer:
                case TomlTokenType.Float:
                case TomlTokenType.Boolean:
                case TomlTokenType.OffsetDateTime:
                case TomlTokenType.LocalDateTime:
                case TomlTokenType.LocalDate:
                case TomlTokenType.LocalTime:
                    values++;
                    break;
            }
        }
        return values;
    }

    /// <summary>Tomlyn's incremental event parser.</summary>
    public static long ParseTomlyn(string toml) {
        var parser = Tomlyn.Parsing.TomlParser.Create(toml);
        long values = 0;
        while (parser.MoveNext()) {
            switch (parser.Current.Kind) {
                case Tomlyn.Parsing.TomlParseEventKind.PropertyName:
                    _ = parser.GetPropertyName();
                    break;
                case Tomlyn.Parsing.TomlParseEventKind.String:
                    _ = parser.GetString();
                    values++;
                    break;
                case Tomlyn.Parsing.TomlParseEventKind.Integer:
                case Tomlyn.Parsing.TomlParseEventKind.Float:
                case Tomlyn.Parsing.TomlParseEventKind.Boolean:
                case Tomlyn.Parsing.TomlParseEventKind.DateTime:
                    values++;
                    break;
            }
        }
        return values;
    }

    // ---- Load: the library's document model ----------------------------------------------------------------------

    public static TomlDocument LoadOfficeIMO(string toml) => TomlDocument.Parse(toml);

    public static Tomlyn.Model.TomlTable LoadTomlyn(string toml) => Tomlyn.TomlSerializer.Deserialize<Tomlyn.Model.TomlTable>(toml)!;

    public static Tomlet.Models.TomlDocument LoadTomlet(string toml) => new Tomlet.TomlParser().Parse(toml);

    /// <summary>Counts leaf values; OfficeIMO decodes each one, which the reference loaders do while loading.</summary>
    public static long CountOfficeIMO(TomlNode node) {
        switch (node) {
            case TomlTable table: {
                long count = 0;
                foreach (KeyValuePair<string, TomlNode> entry in table) count += CountOfficeIMO(entry.Value);
                return count;
            }
            case TomlTableArray tables: return tables.Sum(t => CountOfficeIMO(t));
            case TomlArray array: return array.Sum(CountOfficeIMO);
            default:
                _ = ((TomlValue)node).Value;
                return 1;
        }
    }

    public static long CountTomlyn(object? value) => value switch {
        Tomlyn.Model.TomlTable table => table.Sum(entry => CountTomlyn(entry.Value)),
        Tomlyn.Model.TomlTableArray tables => tables.Sum(t => CountTomlyn(t)),
        Tomlyn.Model.TomlArray array => array.Sum(CountTomlyn),
        _ => 1
    };

    public static long CountTomlet(Tomlet.Models.TomlValue value) => value switch {
        Tomlet.Models.TomlTable table => table.Entries.Values.Sum(CountTomlet),
        Tomlet.Models.TomlArray array => array.ArrayValues.Sum(CountTomlet),
        _ => 1
    };

    // ---- Unchanged load and save ---------------------------------------------------------------------------------

    public static string SaveOfficeIMO(string toml) => TomlDocument.Parse(toml).ToString();

    /// <summary>Tomlyn's full-fidelity syntax tree, which also reproduces its source.</summary>
    public static string SaveTomlyn(string toml) => Tomlyn.Parsing.SyntaxParser.Parse(toml, null, true).ToString();

    /// <summary>Tomlet re-serializes its model; the output is equivalent TOML, not the source text.</summary>
    public static string SaveTomlet(string toml) => new Tomlet.TomlParser().Parse(toml).SerializedValue;
}
