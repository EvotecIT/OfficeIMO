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
    public void Normalize_OrderedListParenMarkers_WhenEnabled() {
        var options = new MarkdownInputNormalizationOptions {
            NormalizeOrderedListParenMarkers = true
        };

        var markdown = "1)First check\n2)   Second check";
        var normalized = MarkdownInputNormalizer.Normalize(markdown, options);

        Assert.Equal("1. First check\n2. Second check", normalized);
    }

    [Fact]
    public void Normalize_OrderedListCaretArtifacts_WhenEnabled() {
        var options = new MarkdownInputNormalizationOptions {
            NormalizeOrderedListCaretArtifacts = true
        };

        var markdown = "1. ^ **Privilege hygiene sweep**\n2.^ **Delegation risk audit**";
        var normalized = MarkdownInputNormalizer.Normalize(markdown, options);

        Assert.Equal("1. **Privilege hygiene sweep**\n2. **Delegation risk audit**", normalized);
    }

    [Fact]
    public void Normalize_TightParentheticalSpacing_WhenEnabled() {
        var options = new MarkdownInputNormalizationOptions {
            NormalizeTightParentheticalSpacing = true
        };

        var markdown = """
1. **Deleted object remnants**(SID left in ACL path)
Top-IDs(AD1/AD2)
""";

        var normalized = MarkdownInputNormalizer.Normalize(markdown, options);

        Assert.Contains("**Deleted object remnants** (SID left in ACL path)", normalized, StringComparison.Ordinal);
        Assert.Contains("Top-IDs (AD1/AD2)", normalized, StringComparison.Ordinal);
    }

    [Fact]
    public void Normalize_TightParentheticalSpacing_DoesNotChangeInlineCodeSpans() {
        var options = new MarkdownInputNormalizationOptions {
            NormalizeTightParentheticalSpacing = true
        };

        var markdown = "Use `Get-ADUser(SIDHistory)` and **Deleted object remnants**(SID left in ACL path)";
        var normalized = MarkdownInputNormalizer.Normalize(markdown, options);

        Assert.Contains("`Get-ADUser(SIDHistory)`", normalized, StringComparison.Ordinal);
        Assert.DoesNotContain("`Get-ADUser (SIDHistory)`", normalized, StringComparison.Ordinal);
        Assert.Contains("**Deleted object remnants** (SID left in ACL path)", normalized, StringComparison.Ordinal);
    }

    [Fact]
    public void Normalize_OrderedListParenAndParentheticalSpacing_DoesNotChangeFencedCodeBlocks() {
        var options = new MarkdownInputNormalizationOptions {
            NormalizeOrderedListParenMarkers = true,
            NormalizeOrderedListCaretArtifacts = true,
            NormalizeTightParentheticalSpacing = true
        };

        var markdown = """
```text
1)First check
2.^ **Delegation risk audit**
3. **Deleted object remnants**(SID left in ACL path)
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

    [Fact]
    public void Normalize_SoftWrappedStrong_DoesNotCollapse_UnorderedListBoundaries() {
        var options = new MarkdownInputNormalizationOptions {
            NormalizeSoftWrappedStrongSpans = true
        };

        var markdown = """
- **AD1:** 875 Events  
- **AD2:** 353 Events
""";

        var normalized = MarkdownInputNormalizer.Normalize(markdown, options);
        Assert.Contains("- **AD1:** 875 Events", normalized, StringComparison.Ordinal);
        Assert.Contains("\n- **AD2:** 353 Events", normalized, StringComparison.Ordinal);
        Assert.DoesNotContain("Events -** AD2:**", normalized, StringComparison.Ordinal);
    }

    [Fact]
    public void Normalize_NestedStrongDelimiters_WhenEnabled() {
        var options = new MarkdownInputNormalizationOptions {
            NormalizeNestedStrongDelimiters = true
        };

        var markdown = "- Signal **AD1 has very high `7034/7023` volume, mostly from **Service Control Manager**.**";
        var normalized = MarkdownInputNormalizer.Normalize(markdown, options);

        Assert.Equal("- Signal **AD1 has very high `7034/7023` volume, mostly from Service Control Manager.**", normalized);
    }

    [Fact]
    public void Normalize_NestedStrongDelimiters_DoesNotChangeFencedCodeBlocks() {
        var options = new MarkdownInputNormalizationOptions {
            NormalizeNestedStrongDelimiters = true
        };

        var markdown = """
```text
- Signal **AD1 has very high `7034/7023` volume, mostly from **Service Control Manager**.**
```
""";

        var normalized = MarkdownInputNormalizer.Normalize(markdown, options);
        Assert.Equal(markdown, normalized);
    }
}
