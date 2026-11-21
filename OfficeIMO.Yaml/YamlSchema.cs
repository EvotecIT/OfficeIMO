namespace OfficeIMO.Yaml;

/// <summary>
/// The rules that resolve untagged plain scalars to null, boolean, integer or floating-point values.
/// </summary>
public enum YamlSchema : byte {
    /// <summary>YAML 1.2 failsafe schema: every scalar is a string.</summary>
    Failsafe = 0,
    /// <summary>YAML 1.2 JSON schema: only JSON spellings (<c>null</c>, <c>true</c>, <c>false</c>, JSON numbers) resolve.</summary>
    Json,
    /// <summary>YAML 1.2 core schema, the default: JSON plus <c>~</c>, empty nulls, case variants, <c>0o</c>/<c>0x</c> integers and <c>.inf</c>/<c>.nan</c>.</summary>
    Core,
    /// <summary>
    /// YAML 1.1 types: <c>yes</c>/<c>no</c>/<c>on</c>/<c>off</c>/<c>y</c>/<c>n</c> booleans, <c>0b</c> and <c>0</c>-prefixed octal integers,
    /// underscores in numbers, sexagesimal numbers (<c>1:20</c>), timestamps and <c>&lt;&lt;</c> merge keys.
    /// </summary>
    Yaml11
}
