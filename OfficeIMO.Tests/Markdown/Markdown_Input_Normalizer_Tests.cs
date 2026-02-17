using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public class Markdown_Input_Normalizer_Tests {
    [Fact]
    public void Normalize_DefaultOptions_LeavesInputUnchanged() {
        var markdown = "**Status\nHEALTHY** and `a\nb`";

        var normalized = MarkdownInputNormalizer.Normalize(markdown);
        Assert.Equal(markdown, normalized);
    }

    [Fact]
    public void Normalize_SoftWrappedStrong_WhenEnabled() {
        var options = new MarkdownInputNormalizationOptions {
            NormalizeSoftWrappedStrongSpans = true
        };

        var normalized = MarkdownInputNormalizer.Normalize("**Status\nHEALTHY**", options);
        Assert.Equal("**Status HEALTHY**", normalized);
    }

    [Fact]
    public void Normalize_InlineCodeLineBreaks_WhenEnabled() {
        var options = new MarkdownInputNormalizationOptions {
            NormalizeInlineCodeSpanLineBreaks = true
        };

        var normalized = MarkdownInputNormalizer.Normalize("`a\nb`", options);
        Assert.Equal("`a b`", normalized);
    }

    [Fact]
    public void Normalize_EscapedInlineCodeSpans_WhenEnabled() {
        var options = new MarkdownInputNormalizationOptions {
            NormalizeEscapedInlineCodeSpans = true
        };

        var normalized = MarkdownInputNormalizer.Normalize(@"Use \`/act act_001\` now.", options);
        Assert.Equal("Use `/act act_001` now.", normalized);
    }

    [Fact]
    public void Normalize_TightStrongBoundaries_WhenEnabled() {
        var options = new MarkdownInputNormalizationOptions {
            NormalizeTightStrongBoundaries = true
        };

        var normalized = MarkdownInputNormalizer.Normalize("Status **Healthy**next", options);
        Assert.Equal("Status **Healthy** next", normalized);
    }

    [Fact]
    public void Normalize_LooseStrongDelimiters_WhenEnabled() {
        var options = new MarkdownInputNormalizationOptions {
            NormalizeLooseStrongDelimiters = true
        };

        var normalized = MarkdownInputNormalizer.Normalize("check ** LDAP/Kerberos health on all DCs** and **unresolved privileged SID targets ** now", options);
        Assert.Equal("check **LDAP/Kerberos health on all DCs** and **unresolved privileged SID targets** now", normalized);
    }

    [Fact]
    public void Normalize_TightStrongBoundaries_DoesNotCorruptAdjacentStrongSpans() {
        var options = new MarkdownInputNormalizationOptions {
            NormalizeTightStrongBoundaries = true
        };

        var markdown = "If you want, I can run a **“Top 8 high-signal security pack”** now, or list only **GPO-related** reports.";

        var normalized = MarkdownInputNormalizer.Normalize(markdown, options);
        Assert.Equal(markdown, normalized);
    }

    [Fact]
    public void Normalize_DoesNotChangeFencedCodeBlocks_ForEscapedCodeAndStrongSpacing() {
        var options = new MarkdownInputNormalizationOptions {
            NormalizeEscapedInlineCodeSpans = true,
            NormalizeTightStrongBoundaries = true
        };

        var markdown = """
```text
Use \`/act act_001\`
Status **Healthy**next
```
""";

        var normalized = MarkdownInputNormalizer.Normalize(markdown, options);
        Assert.Equal(markdown, normalized);
    }

    [Fact]
    public void Normalize_DoesNotChangeTildeFencedCodeBlocks() {
        var options = new MarkdownInputNormalizationOptions {
            NormalizeEscapedInlineCodeSpans = true,
            NormalizeTightStrongBoundaries = true
        };

        var markdown = """
~~~text
Use \`/act act_001\`
Status **Healthy**next
~~~
""";

        var normalized = MarkdownInputNormalizer.Normalize(markdown, options);
        Assert.Equal(markdown, normalized);
    }

    [Fact]
    public void Normalize_LooseStrongDelimiters_DoesNotChangeFencedCodeBlocks() {
        var options = new MarkdownInputNormalizationOptions {
            NormalizeLooseStrongDelimiters = true
        };

        var markdown = """
```text
check ** LDAP/Kerberos health on all DCs** next
```
""";

        var normalized = MarkdownInputNormalizer.Normalize(markdown, options);
        Assert.Equal(markdown, normalized);
    }

    [Fact]
    public void Normalize_OrderedListMarkerSpacing_WhenEnabled() {
        var options = new MarkdownInputNormalizationOptions {
            NormalizeOrderedListMarkerSpacing = true
        };

        var markdown = """
1. **Privilege hygiene sweep**(Domain Admins + other privileged groups)
2.** Delegation risk audit**(unconstrained / constrained / protocol transition)
3.** Replication + DC health snapshot** (stale links, failing partners)
""";

        var normalized = MarkdownInputNormalizer.Normalize(markdown, options);
        Assert.Contains("\n2. ** Delegation risk audit**", normalized, StringComparison.Ordinal);
        Assert.Contains("\n3. ** Replication + DC health snapshot**", normalized, StringComparison.Ordinal);
    }

    [Fact]
    public void Normalize_OrderedListMarkerSpacing_DoesNotChangeFencedCodeBlocks() {
        var options = new MarkdownInputNormalizationOptions {
            NormalizeOrderedListMarkerSpacing = true
        };

        var markdown = """
```text
2.** Delegation risk audit**
3.** Replication + DC health snapshot**
```
""";

        var normalized = MarkdownInputNormalizer.Normalize(markdown, options);
        Assert.Equal(markdown, normalized);
    }

    [Fact]
    public void Normalize_SoftWrappedStrong_DoesNotCollapse_OrderedListBoundaries() {
        var options = new MarkdownInputNormalizationOptions {
            NormalizeSoftWrappedStrongSpans = true,
            NormalizeOrderedListMarkerSpacing = true
        };

        var markdown = """
1. **Privilege hygiene sweep**(Domain Admins + other privileged groups)
2.** Delegation risk audit**(unconstrained / constrained / protocol transition)
3.** Replication + DC health snapshot** (stale links, failing partners)
""";

        var normalized = MarkdownInputNormalizer.Normalize(markdown, options);
        Assert.Contains("\n2. ** Delegation risk audit**", normalized, StringComparison.Ordinal);
        Assert.Contains("\n3. ** Replication + DC health snapshot**", normalized, StringComparison.Ordinal);
        Assert.DoesNotContain("groups) 2.", normalized, StringComparison.Ordinal);
    }
}
