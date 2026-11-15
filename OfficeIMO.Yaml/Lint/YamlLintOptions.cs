namespace OfficeIMO.Yaml;

/// <summary>
/// The rules <see cref="YamlLinter"/> checks and how. The defaults report what is likely a mistake or reads
/// differently in other tools, and leave matters of taste (document markers, empty values, quoting, key order) off.
/// </summary>
public sealed class YamlLintOptions {
    private readonly YamlLintRuleOptions[] _rules;

    /// <summary>Creates options with every rule at its default.</summary>
    public YamlLintOptions() {
        _rules = [
            Indentation, TrailingSpaces, LineLength, Truthy, KeyDuplicates, DocumentStart, EmptyValues, Comments,
            OctalValues, Tabs, NewLineAtEndOfFile, EmptyLines, QuotedStrings, KeyOrdering
        ];
    }

    /// <summary>Gets the <see cref="YamlLintRules.Indentation"/> rule. On by default.</summary>
    public YamlLintIndentationOptions Indentation { get; } = new();

    /// <summary>Gets the <see cref="YamlLintRules.TrailingSpaces"/> rule. On by default.</summary>
    public YamlLintRuleOptions TrailingSpaces { get; } = new(YamlLintRules.TrailingSpaces, enabled: true, OfficeDiagnosticSeverity.Warning);

    /// <summary>Gets the <see cref="YamlLintRules.LineLength"/> rule. On by default, at 120 characters.</summary>
    public YamlLintLineLengthOptions LineLength { get; } = new();

    /// <summary>Gets the <see cref="YamlLintRules.Truthy"/> rule. On by default.</summary>
    public YamlLintTruthyOptions Truthy { get; } = new();

    /// <summary>Gets the <see cref="YamlLintRules.KeyDuplicates"/> rule. On by default, as an error.</summary>
    public YamlLintRuleOptions KeyDuplicates { get; } = new(YamlLintRules.KeyDuplicates, enabled: true, OfficeDiagnosticSeverity.Error);

    /// <summary>Gets the <see cref="YamlLintRules.DocumentStart"/> rule. Off by default.</summary>
    public YamlLintDocumentStartOptions DocumentStart { get; } = new();

    /// <summary>Gets the <see cref="YamlLintRules.EmptyValues"/> rule. Off by default.</summary>
    public YamlLintEmptyValuesOptions EmptyValues { get; } = new();

    /// <summary>Gets the <see cref="YamlLintRules.Comments"/> rule. On by default.</summary>
    public YamlLintCommentsOptions Comments { get; } = new();

    /// <summary>Gets the <see cref="YamlLintRules.OctalValues"/> rule. On by default, for implicit octals only.</summary>
    public YamlLintOctalValuesOptions OctalValues { get; } = new();

    /// <summary>Gets the <see cref="YamlLintRules.Tabs"/> rule. On by default.</summary>
    public YamlLintRuleOptions Tabs { get; } = new(YamlLintRules.Tabs, enabled: true, OfficeDiagnosticSeverity.Warning);

    /// <summary>Gets the <see cref="YamlLintRules.NewLineAtEndOfFile"/> rule. On by default.</summary>
    public YamlLintRuleOptions NewLineAtEndOfFile { get; } = new(YamlLintRules.NewLineAtEndOfFile, enabled: true, OfficeDiagnosticSeverity.Warning);

    /// <summary>Gets the <see cref="YamlLintRules.EmptyLines"/> rule. On by default.</summary>
    public YamlLintEmptyLinesOptions EmptyLines { get; } = new();

    /// <summary>Gets the <see cref="YamlLintRules.QuotedStrings"/> rule. Off by default.</summary>
    public YamlLintQuotedStringsOptions QuotedStrings { get; } = new();

    /// <summary>Gets the <see cref="YamlLintRules.KeyOrdering"/> rule. Off by default.</summary>
    public YamlLintRuleOptions KeyOrdering { get; } = new(YamlLintRules.KeyOrdering, enabled: false, OfficeDiagnosticSeverity.Warning);

    /// <summary>Gets every configurable rule, in the order findings of one line are reported.</summary>
    public IReadOnlyList<YamlLintRuleOptions> Rules => _rules;

    /// <summary>
    /// Gets or sets how the text is loaded, or <see langword="null"/> (the default) for a lenient load that admits
    /// duplicate keys so <see cref="YamlLintRules.KeyDuplicates"/> can report them.
    /// </summary>
    public YamlLoadOptions? LoadOptions { get; set; }

    /// <summary>Returns the rule with an id from <see cref="YamlLintRules"/>, or <see langword="null"/> for an unknown id.</summary>
    public YamlLintRuleOptions? GetRule(string ruleId) {
        foreach (YamlLintRuleOptions rule in _rules) {
            if (string.Equals(rule.Id, ruleId, StringComparison.OrdinalIgnoreCase)) return rule;
        }
        return null;
    }
}

/// <summary>Whether a rule runs and how severe its findings are.</summary>
public class YamlLintRuleOptions {
    internal YamlLintRuleOptions(string id, bool enabled, OfficeDiagnosticSeverity severity) {
        Id = id;
        Enabled = enabled;
        Severity = severity;
    }

    /// <summary>Gets the rule id, one of <see cref="YamlLintRules"/>.</summary>
    public string Id { get; }

    /// <summary>Gets or sets whether the rule runs.</summary>
    public bool Enabled { get; set; }

    /// <summary>Gets or sets the severity of the rule's findings.</summary>
    public OfficeDiagnosticSeverity Severity { get; set; }
}

/// <summary>How block sequences under a mapping key are indented.</summary>
public enum YamlLintSequenceIndentation {
    /// <summary>Every sequence in the file uses the style of the first one.</summary>
    Consistent,
    /// <summary>The dash is indented under the key: <c>key:\n  - item</c>.</summary>
    Indented,
    /// <summary>The dash is level with the key: <c>key:\n- item</c>.</summary>
    NotIndented,
    /// <summary>Either style, mixed.</summary>
    Any
}

/// <summary>Settings of the <see cref="YamlLintRules.Indentation"/> rule.</summary>
public sealed class YamlLintIndentationOptions : YamlLintRuleOptions {
    private int? _spaces;

    internal YamlLintIndentationOptions() : base(YamlLintRules.Indentation, enabled: true, OfficeDiagnosticSeverity.Warning) { }

    /// <summary>
    /// Gets or sets the indentation width of nested block collections, or <see langword="null"/> (the default) for the
    /// width the file uses first. 1 to 9.
    /// </summary>
    public int? Spaces {
        get => _spaces;
        set {
            if (value is < 1 or > 9) throw new ArgumentOutOfRangeException(nameof(value), value, "Spaces must be between 1 and 9.");
            _spaces = value;
        }
    }

    /// <summary>Gets or sets how sequences under a key are indented. Defaults to <see cref="YamlLintSequenceIndentation.Consistent"/>.</summary>
    public YamlLintSequenceIndentation Sequences { get; set; }
}

/// <summary>Settings of the <see cref="YamlLintRules.LineLength"/> rule.</summary>
public sealed class YamlLintLineLengthOptions : YamlLintRuleOptions {
    private int _max = 120;

    internal YamlLintLineLengthOptions() : base(YamlLintRules.LineLength, enabled: true, OfficeDiagnosticSeverity.Warning) { }

    /// <summary>Gets or sets the longest allowed line, in characters. Defaults to 120.</summary>
    public int Max {
        get => _max;
        set => _max = value > 0 ? value : throw new ArgumentOutOfRangeException(nameof(value), value, "Max must be positive.");
    }

    /// <summary>Gets or sets whether a line whose content is one word without spaces, such as a URL, may be longer. Defaults to <see langword="true"/>.</summary>
    public bool AllowNonBreakableWords { get; set; } = true;

    /// <summary>Gets or sets whether <c>key: word</c> may be longer when the value is one word without spaces. Defaults to <see langword="true"/>.</summary>
    public bool AllowNonBreakableInlineMappings { get; set; } = true;
}

/// <summary>Settings of the <see cref="YamlLintRules.Truthy"/> rule.</summary>
public sealed class YamlLintTruthyOptions : YamlLintRuleOptions {
    internal YamlLintTruthyOptions() : base(YamlLintRules.Truthy, enabled: true, OfficeDiagnosticSeverity.Warning) { }

    /// <summary>
    /// Gets the boolean spellings that are not reported. Defaults to the YAML 1.2 core spellings (<c>true</c>,
    /// <c>True</c>, <c>TRUE</c> and the same for false), so only <c>yes</c>, <c>no</c>, <c>on</c> and <c>off</c> in
    /// any case are reported. Remove <c>True</c> and the other capitalized forms to require lowercase booleans.
    /// </summary>
    public ISet<string> AllowedValues { get; } = new HashSet<string>(StringComparer.Ordinal) { "true", "True", "TRUE", "false", "False", "FALSE" };

    /// <summary>
    /// Gets or sets whether mapping keys are checked too. Defaults to <see langword="false"/>, so a GitHub Actions
    /// <c>on:</c> key is not reported.
    /// </summary>
    public bool CheckKeys { get; set; }
}

/// <summary>Settings of the <see cref="YamlLintRules.DocumentStart"/> rule.</summary>
public sealed class YamlLintDocumentStartOptions : YamlLintRuleOptions {
    internal YamlLintDocumentStartOptions() : base(YamlLintRules.DocumentStart, enabled: false, OfficeDiagnosticSeverity.Warning) { }

    /// <summary>
    /// Gets or sets whether the first document must start with <c>---</c> (<see langword="true"/>, the default) or
    /// must not (<see langword="false"/>). Later documents of a stream always need the marker.
    /// </summary>
    public bool Present { get; set; } = true;
}

/// <summary>Settings of the <see cref="YamlLintRules.EmptyValues"/> rule.</summary>
public sealed class YamlLintEmptyValuesOptions : YamlLintRuleOptions {
    internal YamlLintEmptyValuesOptions() : base(YamlLintRules.EmptyValues, enabled: false, OfficeDiagnosticSeverity.Warning) { }

    /// <summary>Gets or sets whether <c>key:</c> without a value is reported in block mappings. Defaults to <see langword="true"/>.</summary>
    public bool ForbidInBlockMappings { get; set; } = true;

    /// <summary>Gets or sets whether <c>{key: }</c> without a value is reported in flow mappings. Defaults to <see langword="true"/>.</summary>
    public bool ForbidInFlowMappings { get; set; } = true;

    /// <summary>Gets or sets whether <c>-</c> without an item is reported in block sequences. Defaults to <see langword="true"/>.</summary>
    public bool ForbidInBlockSequences { get; set; } = true;
}

/// <summary>Settings of the <see cref="YamlLintRules.Comments"/> rule.</summary>
public sealed class YamlLintCommentsOptions : YamlLintRuleOptions {
    private int _minSpacesFromContent = 2;

    internal YamlLintCommentsOptions() : base(YamlLintRules.Comments, enabled: true, OfficeDiagnosticSeverity.Warning) { }

    /// <summary>Gets or sets whether the comment text must start with a space after <c>#</c> (<c>#####</c> rulers excepted). Defaults to <see langword="true"/>.</summary>
    public bool RequireStartingSpace { get; set; } = true;

    /// <summary>Gets or sets whether a <c>#!</c> line at the start of the file is exempt. Defaults to <see langword="true"/>.</summary>
    public bool IgnoreShebangs { get; set; } = true;

    /// <summary>Gets or sets the fewest spaces between content and a comment on the same line; 0 turns the check off. Defaults to 2.</summary>
    public int MinSpacesFromContent {
        get => _minSpacesFromContent;
        set => _minSpacesFromContent = value >= 0 ? value : throw new ArgumentOutOfRangeException(nameof(value), value, "MinSpacesFromContent cannot be negative.");
    }
}

/// <summary>Settings of the <see cref="YamlLintRules.OctalValues"/> rule.</summary>
public sealed class YamlLintOctalValuesOptions : YamlLintRuleOptions {
    internal YamlLintOctalValuesOptions() : base(YamlLintRules.OctalValues, enabled: true, OfficeDiagnosticSeverity.Warning) { }

    /// <summary>Gets or sets whether <c>010</c>, an octal 8 to YAML 1.1 readers and a decimal 10 to YAML 1.2 readers, is reported. Defaults to <see langword="true"/>.</summary>
    public bool ForbidImplicitOctal { get; set; } = true;

    /// <summary>Gets or sets whether <c>0o10</c>, an octal 8 to YAML 1.2 readers and a string to YAML 1.1 readers, is reported. Defaults to <see langword="false"/>.</summary>
    public bool ForbidExplicitOctal { get; set; }
}

/// <summary>Settings of the <see cref="YamlLintRules.EmptyLines"/> rule.</summary>
public sealed class YamlLintEmptyLinesOptions : YamlLintRuleOptions {
    private int _max = 2;
    private int _maxStart;
    private int _maxEnd;

    internal YamlLintEmptyLinesOptions() : base(YamlLintRules.EmptyLines, enabled: true, OfficeDiagnosticSeverity.Warning) { }

    /// <summary>Gets or sets the most consecutive blank lines inside the file. Defaults to 2.</summary>
    public int Max {
        get => _max;
        set => _max = CheckCount(value, nameof(Max));
    }

    /// <summary>Gets or sets the most blank lines at the start of the file. Defaults to 0.</summary>
    public int MaxStart {
        get => _maxStart;
        set => _maxStart = CheckCount(value, nameof(MaxStart));
    }

    /// <summary>Gets or sets the most blank lines at the end of the file. Defaults to 0.</summary>
    public int MaxEnd {
        get => _maxEnd;
        set => _maxEnd = CheckCount(value, nameof(MaxEnd));
    }

    private static int CheckCount(int value, string name) =>
        value >= 0 ? value : throw new ArgumentOutOfRangeException(name, value, name + " cannot be negative.");
}

/// <summary>Which quotes the <see cref="YamlLintRules.QuotedStrings"/> rule expects.</summary>
public enum YamlLintQuoteType {
    /// <summary>Single or double quotes.</summary>
    Any,
    /// <summary>Single quotes, except where the value needs double-quoted escapes.</summary>
    Single,
    /// <summary>Double quotes.</summary>
    Double
}

/// <summary>When the <see cref="YamlLintRules.QuotedStrings"/> rule expects string values to be quoted.</summary>
public enum YamlLintQuoteRequirement {
    /// <summary>Every string value is quoted.</summary>
    Always,
    /// <summary>Quoting is optional; only the quote type is checked.</summary>
    Optional,
    /// <summary>Only values that would read as something else, or cannot be written plain, are quoted.</summary>
    OnlyWhenNeeded
}

/// <summary>Settings of the <see cref="YamlLintRules.QuotedStrings"/> rule.</summary>
public sealed class YamlLintQuotedStringsOptions : YamlLintRuleOptions {
    internal YamlLintQuotedStringsOptions() : base(YamlLintRules.QuotedStrings, enabled: false, OfficeDiagnosticSeverity.Warning) { }

    /// <summary>Gets or sets the expected quotes. Defaults to <see cref="YamlLintQuoteType.Any"/>.</summary>
    public YamlLintQuoteType QuoteType { get; set; }

    /// <summary>Gets or sets when string values must be quoted. Defaults to <see cref="YamlLintQuoteRequirement.Always"/>.</summary>
    public YamlLintQuoteRequirement Required { get; set; }

    /// <summary>Gets or sets whether mapping keys are checked too. Defaults to <see langword="false"/>.</summary>
    public bool CheckKeys { get; set; }
}
