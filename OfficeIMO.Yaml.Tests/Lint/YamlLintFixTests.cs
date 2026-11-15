using OfficeIMO.Yaml.Tests.Conformance;
using Xunit;

namespace OfficeIMO.Yaml.Tests;

public class YamlLintFixTests {
    [Fact]
    public void FixChangesOnlyTheLinesItRepairs() {
        string yaml =
            "#header\n" +
            "name:   web    # keep   spacing inside\n" +
            "enabled: yes\n" +
            "mode: 0755\n" +
            "flag: TRUE\n" +
            "script: |\n" +
            "  echo  \n" +
            "list: [ 1,2 ]\n" +
            "\n\n\n\n" +
            "tail: \"x\"   \n" +
            "\n";
        YamlLintFixResult result = YamlLinter.Fix(yaml);
        // The file quotes with double quotes, so the quoted 'yes' follows it.
        Assert.Equal(
            "# header\n" +
            "name:   web    # keep   spacing inside\n" +
            "enabled: \"yes\"\n" +
            "mode: 0755\n" +
            "flag: TRUE\n" +
            "script: |\n" +
            "  echo  \n" +
            "list: [ 1,2 ]\n" +
            "\n\n" +
            "tail: \"x\"\n", result.Text);
        Assert.True(result.Changed);
        Assert.Equal(
            new[] { YamlLintRules.Comments, YamlLintRules.Truthy, YamlLintRules.EmptyLines, YamlLintRules.TrailingSpaces, YamlLintRules.EmptyLines },
            result.Applied.Select(f => f.RuleId));
        // What has no safe fix stays: the octal and the spaces inside the block scalar.
        Assert.Equal(new[] { YamlLintRules.OctalValues, YamlLintRules.TrailingSpaces }, result.Remaining.Select(f => f.RuleId));
    }

    [Fact]
    public void FixLeavesCleanTextUntouched() {
        string yaml = "a: 1\n";
        YamlLintFixResult result = YamlLinter.Fix(yaml);
        Assert.False(result.Changed);
        Assert.Same(yaml, result.Text);
    }

    [Fact]
    public void FixDoesNothingWhenTheTextDoesNotLoad() {
        YamlLintFixResult result = YamlLinter.Fix("a: 1   \nb: [1\n");
        Assert.False(result.Changed);
        Assert.Contains(result.Remaining, f => f.RuleId == YamlLintRules.Syntax);
    }

    [Fact]
    public void FixWithEveryRuleKeepsCorpusValuesAndConverges() {
        var options = new YamlLintOptions();
        foreach (YamlLintRuleOptions rule in options.Rules) rule.Enabled = true;
        options.QuotedStrings.QuoteType = YamlLintQuoteType.Single;
        options.QuotedStrings.Required = YamlLintQuoteRequirement.OnlyWhenNeeded;
        int changed = 0;
        foreach (YamlCorpusFile file in YamlCorpus.Load()) {
            string source = YamlTextTestHelpers.DecodeSuiteInput(file.Bytes);
            YamlLintFixResult result = YamlLinter.Fix(source, options);
            if (!result.Changed) continue;
            changed++;
            Assert.True(YamlValueEquivalence.SameValues(source, result.Text, null), file.Path + ": a fix changed a value.");
            Assert.DoesNotContain(result.Remaining, f => f.Fix is not null);
            Assert.False(YamlLinter.Fix(result.Text, options).Changed, file.Path + ": fixing again changed the text.");
        }
        Assert.True(changed > 0);
    }

    [Fact]
    public void SpacesEndingABlockScalarAtTheEndOfTheFileAreKept() {
        const string yaml = "b: 1   \na: |\n  text  ";
        YamlLintFixResult result = YamlLinter.Fix(yaml);
        Assert.Equal("b: 1\na: |\n  text  ", result.Text);
        Assert.Equal(new[] { YamlLintRules.TrailingSpaces }, result.Applied.Select(f => f.RuleId));
        Assert.Equal(new[] { YamlLintRules.TrailingSpaces, YamlLintRules.NewLineAtEndOfFile }, result.Remaining.Select(f => f.RuleId));
    }
}
