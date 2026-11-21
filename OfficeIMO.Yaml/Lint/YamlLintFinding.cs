namespace OfficeIMO.Yaml;

/// <summary>One problem <see cref="YamlLinter"/> found, with where it is and, when it can be repaired safely, how.</summary>
public sealed class YamlLintFinding {
    internal YamlLintFinding(string ruleId, OfficeDiagnosticSeverity severity, YamlPosition position, string message, YamlLintFix? fix) {
        RuleId = ruleId;
        Severity = severity;
        Position = position;
        Message = message;
        Fix = fix;
    }

    /// <summary>Gets the rule that reported the finding, one of the <see cref="YamlLintRules"/> ids.</summary>
    public string RuleId { get; }

    /// <summary>Gets the severity configured for the rule.</summary>
    public OfficeDiagnosticSeverity Severity { get; }

    /// <summary>Gets where the problem is in the linted text.</summary>
    public YamlPosition Position { get; }

    /// <summary>Gets the one-based line of <see cref="Position"/>.</summary>
    public int Line => Position.Line;

    /// <summary>Gets the one-based column of <see cref="Position"/>, in UTF-16 code units.</summary>
    public int Column => Position.Column;

    /// <summary>Gets a sentence describing the problem.</summary>
    public string Message { get; }

    /// <summary>
    /// Gets the text edit that repairs the problem without changing any value, or <see langword="null"/> when the rule
    /// has no safe repair for it. <see cref="YamlLinter.Fix(string, YamlLintOptions?)"/> applies these edits.
    /// </summary>
    public YamlLintFix? Fix { get; }

    /// <summary>Formats the finding as <c>line:column severity message (rule)</c>, the layout yamllint prints.</summary>
    public override string ToString() =>
        Position + " " + (Severity == OfficeDiagnosticSeverity.Error ? "error" : Severity == OfficeDiagnosticSeverity.Warning ? "warning" : "info") +
        " " + Message + " (" + RuleId + ")";
}

/// <summary>
/// A replacement of a range of the linted text. Edits touch only the lines they repair; everything else, comments and
/// spellings included, stays as it was.
/// </summary>
public sealed class YamlLintFix {
    internal YamlLintFix(int offset, int length, string replacement) {
        Offset = offset;
        Length = length;
        Replacement = replacement;
    }

    /// <summary>Gets the zero-based offset of the replaced range, in UTF-16 code units of the linted text.</summary>
    public int Offset { get; }

    /// <summary>Gets the length of the replaced range; 0 for an insertion.</summary>
    public int Length { get; }

    /// <summary>Gets the text that replaces the range; empty for a deletion.</summary>
    public string Replacement { get; }
}

/// <summary>The outcome of <see cref="YamlLinter.Fix(string, YamlLintOptions?)"/>.</summary>
public sealed class YamlLintFixResult {
    internal YamlLintFixResult(string text, IReadOnlyList<YamlLintFinding> applied, IReadOnlyList<YamlLintFinding> remaining) {
        Text = text;
        Applied = applied;
        Remaining = remaining;
    }

    /// <summary>Gets the repaired text; the input itself when nothing was fixed.</summary>
    public string Text { get; }

    /// <summary>Gets whether any fix was applied.</summary>
    public bool Changed => Applied.Count > 0;

    /// <summary>
    /// Gets the findings whose fixes were applied. Most are positioned in the input; a fix that only an earlier fix made
    /// possible, such as a blank line left behind, is positioned in the text that earlier fix produced.
    /// </summary>
    public IReadOnlyList<YamlLintFinding> Applied { get; }

    /// <summary>Gets the findings that linting the repaired text still reports.</summary>
    public IReadOnlyList<YamlLintFinding> Remaining { get; }
}
