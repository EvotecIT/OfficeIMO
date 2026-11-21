namespace OfficeIMO.Yaml;

/// <summary>
/// Checks YAML text against a configurable rule set modelled on yamllint (<see cref="YamlLintRules"/>) and repairs
/// what can be repaired without changing a value.
/// </summary>
/// <remarks>
/// Linting never changes the text. Rules that read the structure (indentation, booleans, keys, quoting) need text that
/// loads; when it does not, the result holds a <see cref="YamlLintRules.Syntax"/> error and the line rules, without
/// fixes.
/// </remarks>
public static class YamlLinter {
    /// <summary>Passes of <see cref="Fix(string, YamlLintOptions?)"/>; a fix can expose another, such as blank lines left by removed spaces.</summary>
    private const int FixPasses = 4;

    /// <summary>Lints YAML text.</summary>
    /// <param name="yaml">The text to check.</param>
    /// <param name="options">The rules to apply, or <see langword="null"/> for the defaults.</param>
    /// <returns>The findings, ordered by position.</returns>
    public static IReadOnlyList<YamlLintFinding> Lint(string yaml, YamlLintOptions? options = null) {
        if (yaml is null) throw new ArgumentNullException(nameof(yaml));
        return new YamlLintRun(yaml, options ?? new YamlLintOptions()).Run();
    }

    /// <summary>
    /// Applies the safe fixes of every enabled rule and returns the repaired text. Each fix replaces only the text of
    /// the lines it repairs, so comments, spellings and layout elsewhere stay byte for byte. The repaired text is
    /// loaded again and compared with the input; if any value, tag or anchor differs, nothing is applied.
    /// </summary>
    /// <param name="yaml">The text to repair.</param>
    /// <param name="options">The rules to apply, or <see langword="null"/> for the defaults.</param>
    public static YamlLintFixResult Fix(string yaml, YamlLintOptions? options = null) {
        if (yaml is null) throw new ArgumentNullException(nameof(yaml));
        options ??= new YamlLintOptions();
        var applied = new List<YamlLintFinding>();
        string text = yaml;
        IReadOnlyList<YamlLintFinding> findings = Lint(text, options);
        for (int pass = 0; pass < FixPasses; pass++) {
            string next = Apply(text, findings, applied, out int count);
            if (count == 0) break;
            text = next;
            findings = Lint(text, options);
        }
        if (ReferenceEquals(text, yaml)) return new YamlLintFixResult(yaml, Array.Empty<YamlLintFinding>(), findings);
        if (!YamlValueEquivalence.SameValues(yaml, text, options.LoadOptions)) {
            return new YamlLintFixResult(yaml, Array.Empty<YamlLintFinding>(), Lint(yaml, options));
        }
        return new YamlLintFixResult(text, applied, findings);
    }

    /// <summary>Applies the non-overlapping fixes of <paramref name="findings"/>, earliest first.</summary>
    private static string Apply(string text, IReadOnlyList<YamlLintFinding> findings, List<YamlLintFinding> applied, out int count) {
        var fixes = new List<YamlLintFinding>();
        foreach (YamlLintFinding finding in findings) {
            if (finding.Fix is not null) fixes.Add(finding);
        }
        count = 0;
        if (fixes.Count == 0) return text;
        // A stable sort keeps rule order among fixes at one offset.
        YamlLintFinding[] ordered = fixes.OrderBy(static f => f.Fix!.Offset).ThenBy(static f => f.Fix!.Length).ToArray();
        var output = new StringBuilder(text.Length + 16);
        int copied = 0, lastOffset = -1;
        foreach (YamlLintFinding finding in ordered) {
            YamlLintFix fix = finding.Fix!;
            // Overlapping edits, or two at one offset, wait for the next pass, which lints the text the first leaves.
            if (fix.Offset < copied || fix.Offset == lastOffset) continue;
            output.Append(text, copied, fix.Offset - copied).Append(fix.Replacement);
            copied = fix.Offset + fix.Length;
            lastOffset = fix.Offset;
            count++;
            applied.Add(finding);
        }
        output.Append(text, copied, text.Length - copied);
        return output.ToString();
    }
}
