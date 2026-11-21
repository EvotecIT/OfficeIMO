namespace OfficeIMO.Yaml;

/// <summary>
/// The ids of the rules <see cref="YamlLinter"/> checks. They match the yamllint rule names where yamllint has the same
/// rule, so findings and configuration read the same in both tools.
/// </summary>
public static class YamlLintRules {
    /// <summary>The text is not valid YAML. Always reported; the other rules then check only lines.</summary>
    public const string Syntax = "syntax";

    /// <summary>Nested block collections indented by a different width than the rest of the file.</summary>
    public const string Indentation = "indentation";

    /// <summary>Spaces or tabs at the end of a line.</summary>
    public const string TrailingSpaces = "trailing-spaces";

    /// <summary>Lines longer than the configured maximum.</summary>
    public const string LineLength = "line-length";

    /// <summary>Plain values such as <c>yes</c>, <c>on</c> or <c>no</c> that YAML 1.1 and YAML 1.2 readers read differently.</summary>
    public const string Truthy = "truthy";

    /// <summary>A mapping key that repeats an earlier key of the same mapping.</summary>
    public const string KeyDuplicates = "key-duplicates";

    /// <summary>The document start marker <c>---</c> is missing, or present where it is forbidden.</summary>
    public const string DocumentStart = "document-start";

    /// <summary>A mapping value or sequence item left empty, which reads as null.</summary>
    public const string EmptyValues = "empty-values";

    /// <summary>A comment without a space after <c>#</c>, or too close to the content before it.</summary>
    public const string Comments = "comments";

    /// <summary>Integers such as <c>010</c> that YAML 1.1 readers take as octal and YAML 1.2 readers as decimal.</summary>
    public const string OctalValues = "octal-values";

    /// <summary>Tab characters in a line's indentation.</summary>
    public const string Tabs = "tabs";

    /// <summary>The file does not end with a line break.</summary>
    public const string NewLineAtEndOfFile = "new-line-at-end-of-file";

    /// <summary>Blank lines beyond the configured maximum, at the start, inside or at the end of the file.</summary>
    public const string EmptyLines = "empty-lines";

    /// <summary>String values quoted in a different style than the configured one, or not quoted as required.</summary>
    public const string QuotedStrings = "quoted-strings";

    /// <summary>Mapping keys that are not in ascending order.</summary>
    public const string KeyOrdering = "key-ordering";
}
