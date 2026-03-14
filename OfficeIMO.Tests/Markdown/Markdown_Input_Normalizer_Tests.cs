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
    public void Normalize_TightArrowStrongBoundaries_WhenEnabled() {
        var options = new MarkdownInputNormalizationOptions {
            NormalizeTightArrowStrongBoundaries = true
        };

        var normalized = MarkdownInputNormalizer.Normalize("Signal ->**Why it matters:**coverage", options);
        Assert.Equal("Signal -> **Why it matters:**coverage", normalized);
    }

    [Fact]
    public void Normalize_BrokenStrongArrowLabels_WhenEnabled() {
        var options = new MarkdownInputNormalizationOptions {
            NormalizeBrokenStrongArrowLabels = true
        };

        var normalized = MarkdownInputNormalizer.Normalize("- Signal **No current failures -> **Why it matters:** transport/auth issues", options);
        Assert.Equal("- Signal **No current failures** -> **Why it matters:** transport/auth issues", normalized);
    }

    [Fact]
    public void Normalize_WrappedSignalFlowStrongRuns_WhenEnabled() {
        var options = new MarkdownInputNormalizationOptions {
            NormalizeWrappedSignalFlowStrongRuns = true,
            NormalizeTightArrowStrongBoundaries = true,
            NormalizeTightStrongBoundaries = true
        };

        var markdown = "- Signal **Catalog count includes hidden/disabled/deprecated rules -> **Why it matters:**external/custom rules can drift or disappear between hosts ->**Next action:**break down `rule_origin` (`builtin` vs `external`) and confirm expected external rules are present.**";
        var normalized = MarkdownInputNormalizer.Normalize(markdown, options);

        Assert.Equal("- Signal **Catalog count includes hidden/disabled/deprecated rules** -> **Why it matters:** external/custom rules can drift or disappear between hosts -> **Next action:** break down `rule_origin` (`builtin` vs `external`) and confirm expected external rules are present.", normalized);
    }

    [Fact]
    public void Normalize_CollapsedMetricChains_WhenEnabled() {
        var options = new MarkdownInputNormalizationOptions {
            NormalizeCollapsedMetricChains = true,
            NormalizeMetricValueStrongRuns = true
        };

        var markdown = "**Status: HEALTHY** - **Servers checked:**5 -**Replication edges:**62 -*Failed edges:**0 -*Stale edges (>24h):**0 - **Servers with failures:**0";
        var normalized = MarkdownInputNormalizer.Normalize(markdown, options);

        Assert.Contains("Status **HEALTHY**", normalized, StringComparison.Ordinal);
        Assert.Contains("- Servers checked **5**", normalized, StringComparison.Ordinal);
        Assert.Contains("- Replication edges **62**", normalized, StringComparison.Ordinal);
        Assert.Contains("- Failed edges **0**", normalized, StringComparison.Ordinal);
        Assert.Contains("- Stale edges (>24h) **0**", normalized, StringComparison.Ordinal);
        Assert.Contains("- Servers with failures **0**", normalized, StringComparison.Ordinal);
        Assert.DoesNotContain("-**", normalized, StringComparison.Ordinal);
        Assert.DoesNotContain("**Servers checked:**", normalized, StringComparison.Ordinal);
    }

    [Fact]
    public void Normalize_CollapsedMetricChains_WhenEnabled_WithCrLfInput() {
        var options = new MarkdownInputNormalizationOptions {
            NormalizeCollapsedMetricChains = true,
            NormalizeMetricValueStrongRuns = true
        };

        var markdown = "**Status: HEALTHY**\r\n- **Servers checked:**5\r\n- **Replication edges:**62\r\n- **Failed edges:**0\r\n- **Stale edges (>24h):**0\r\n- **Servers with failures:**0\r\n";
        var normalized = MarkdownInputNormalizer.Normalize(markdown, options);

        Assert.Contains("- Servers checked **5**", normalized, StringComparison.Ordinal);
        Assert.Contains("- Replication edges **62**", normalized, StringComparison.Ordinal);
        Assert.Contains("- Failed edges **0**", normalized, StringComparison.Ordinal);
        Assert.Contains("- Stale edges (>24h) **0**", normalized, StringComparison.Ordinal);
        Assert.Contains("- Servers with failures **0**", normalized, StringComparison.Ordinal);
        Assert.DoesNotContain("**Servers checked:**", normalized, StringComparison.Ordinal);
    }

    [Fact]
    public void Normalize_CollapsedMetricChains_DoesNotSplitInlineHyphenProse() {
        var options = new MarkdownInputNormalizationOptions {
            NormalizeCollapsedMetricChains = true
        };

        var markdown = "Health note: foo - **bar** should stay inline.";
        var normalized = MarkdownInputNormalizer.Normalize(markdown, options);

        Assert.Equal(markdown, normalized);
    }

    [Fact]
    public void Normalize_HostLabelBulletArtifacts_WhenEnabled() {
        var options = new MarkdownInputNormalizationOptions {
            NormalizeHostLabelBulletArtifacts = true
        };

        var markdown = "-AD1\nhealthy for directory access\n-AD2 ready";
        var normalized = MarkdownInputNormalizer.Normalize(markdown, options);

        Assert.Equal("- AD1 healthy for directory access\n- AD2 ready", normalized);
    }

    [Fact]
    public void Normalize_HostLabelBulletArtifacts_DoesNotRewriteFencedCode() {
        var options = new MarkdownInputNormalizationOptions {
            NormalizeHostLabelBulletArtifacts = true
        };

        var markdown = "```text\n-AD1\nhealthy for directory access\n```";
        var normalized = MarkdownInputNormalizer.Normalize(markdown, options);

        Assert.Equal(markdown, normalized);
    }

    [Fact]
    public void Normalize_StandaloneHashHeadingSeparators_WhenEnabled() {
        var options = new MarkdownInputNormalizationOptions {
            NormalizeStandaloneHashHeadingSeparators = true
        };

        var markdown = "#\n\n### Forest Replication Status\n- Overall health ✅ Healthy****";
        var normalized = MarkdownInputNormalizer.Normalize(markdown, options);

        Assert.Equal("### Forest Replication Status\n- Overall health ✅ Healthy****", normalized);
    }

    [Fact]
    public void Normalize_StandaloneHashHeadingSeparators_DoesNotRewriteOrdinaryHashLine() {
        var options = new MarkdownInputNormalizationOptions {
            NormalizeStandaloneHashHeadingSeparators = true
        };

        var markdown = "Inventory legend:\n#\nkeep this line as-is";
        var normalized = MarkdownInputNormalizer.Normalize(markdown, options);

        Assert.Equal(markdown, normalized);
    }

    [Fact]
    public void Normalize_BrokenTwoLineStrongLeadIns_WhenEnabled() {
        var options = new MarkdownInputNormalizationOptions {
            NormalizeBrokenTwoLineStrongLeadIns = true
        };

        var markdown = """
**Result
all 5 are healthy for directory access** with recommended LDAPS endpoints.
""";
        var normalized = MarkdownInputNormalizer.Normalize(markdown, options);

        Assert.Equal("**Result:** all 5 are healthy for directory access with recommended LDAPS endpoints.", normalized);
    }

    [Fact]
    public void Normalize_BrokenTwoLineStrongLeadIns_PreservesLegitimateMultilineBoldContent() {
        var options = new MarkdownInputNormalizationOptions {
            NormalizeBrokenTwoLineStrongLeadIns = true
        };

        var markdown = """
**Keep
this together**
""";
        var normalized = MarkdownInputNormalizer.Normalize(markdown, options);

        Assert.Equal(markdown, normalized);
    }

    [Fact]
    public void Normalize_CompactHeadingAndStrongLabelListBoundaries_WhenEnabled() {
        var options = new MarkdownInputNormalizationOptions {
            NormalizeHeadingListBoundaries = true,
            NormalizeCompactStrongLabelListBoundaries = true
        };

        var markdown = "## Wynik ogólny- **Replication:** wcześniej zdrowa ✅- **FSMO:** technicznie OK";
        var normalized = MarkdownInputNormalizer.Normalize(markdown, options);

        Assert.Equal("## Wynik ogólny\n- **Replication:** wcześniej zdrowa ✅\n- **FSMO:** technicznie OK", normalized);
    }

    [Fact]
    public void Normalize_CompactHeadingBoundaries_WhenEnabled() {
        var options = new MarkdownInputNormalizationOptions {
            NormalizeCompactHeadingBoundaries = true
        };

        var normalized = MarkdownInputNormalizer.Normalize("previous shutdown was unexpected### Reason", options);
        Assert.Equal("previous shutdown was unexpected\n### Reason", normalized);
    }

    [Fact]
    public void Normalize_ColonListBoundaries_WhenEnabled() {
        var options = new MarkdownInputNormalizationOptions {
            NormalizeColonListBoundaries = true
        };

        var normalized = MarkdownInputNormalizer.Normalize("Następny najlepszy krok:- **`ad_domain_controller_facts`**", options);
        Assert.Equal("Następny najlepszy krok:\n- **`ad_domain_controller_facts`**", normalized);
    }

    [Fact]
    public void Normalize_CompactJsonFenceBodyBoundary_WhenEnabled() {
        var options = new MarkdownInputNormalizationOptions {
            NormalizeCompactFenceBodyBoundaries = true
        };

        var markdown = "```json{\"log_name\":\"System\"}\n```";
        var normalized = MarkdownInputNormalizer.Normalize(markdown, options);

        Assert.Equal("```json\n{\"log_name\":\"System\"}\n```", normalized);
    }

    [Fact]
    public void Normalize_CompactMermaidFenceBodyBoundary_WhenEnabled() {
        var options = new MarkdownInputNormalizationOptions {
            NormalizeCompactFenceBodyBoundaries = true
        };

        var markdown = "```mermaidflowchart LR A-->B\n```";
        var normalized = MarkdownInputNormalizer.Normalize(markdown, options);

        Assert.Equal("```mermaid\nflowchart LR A-->B\n```", normalized);
    }

    [Fact]
    public void Normalize_CompactQuotedMermaidFenceBodyBoundary_WhenEnabled() {
        var options = new MarkdownInputNormalizationOptions {
            NormalizeCompactFenceBodyBoundaries = true
        };

        var markdown = "> ```mermaidflowchart LR A-->B\n> ```";
        var normalized = MarkdownInputNormalizer.Normalize(markdown, options);

        Assert.Equal("> ```mermaid\n> flowchart LR A-->B\n> ```", normalized);
    }

    [Fact]
    public void Normalize_CompactNestedQuotedMermaidFenceBodyBoundary_WhenEnabled() {
        var options = new MarkdownInputNormalizationOptions {
            NormalizeCompactFenceBodyBoundaries = true
        };

        var markdown = "> > ```mermaidflowchart LR A-->B\n> > ```";
        var normalized = MarkdownInputNormalizer.Normalize(markdown, options);

        Assert.Equal("> > ```mermaid\n> > flowchart LR A-->B\n> > ```", normalized);
    }

    [Fact]
    public void Normalize_CompactListQuotedMermaidFenceBodyBoundary_WhenEnabled() {
        var options = new MarkdownInputNormalizationOptions {
            NormalizeCompactFenceBodyBoundaries = true
        };

        var markdown = "- item\n\n  > ```mermaidflowchart LR A-->B\n  > ```";
        var normalized = MarkdownInputNormalizer.Normalize(markdown, options);

        Assert.Equal("- item\n\n  > ```mermaid\n  > flowchart LR A-->B\n  > ```", normalized);
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
    public void Normalize_RepeatedStrongDelimiterRuns_WhenEnabled() {
        var options = new MarkdownInputNormalizationOptions {
            NormalizeLooseStrongDelimiters = true
        };

        var normalized = MarkdownInputNormalizer.Normalize("- Overall health ****healthy****", options);
        Assert.Equal("- Overall health **healthy**", normalized);
    }

    [Fact]
    public void Normalize_DanglingTrailingStrongListClosers_WhenEnabled() {
        var options = new MarkdownInputNormalizationOptions {
            NormalizeDanglingTrailingStrongListClosers = true
        };

        var normalized = MarkdownInputNormalizer.Normalize("- Overall health ✅ Healthy****", options);
        Assert.Equal("- Overall health ✅ **Healthy**", normalized);
    }

    [Fact]
    public void Normalize_DanglingTrailingStrongListClosers_DoesNotRewriteOrdinaryProse() {
        var options = new MarkdownInputNormalizationOptions {
            NormalizeDanglingTrailingStrongListClosers = true
        };

        var markdown = "Literal marker code****";
        var normalized = MarkdownInputNormalizer.Normalize(markdown, options);
        Assert.Equal(markdown, normalized);
    }

    [Fact]
    public void Normalize_MetricValueStrongRuns_WhenEnabled() {
        var options = new MarkdownInputNormalizationOptions {
            NormalizeMetricValueStrongRuns = true
        };

        var markdown = """
- Overall health ******healthy**
- Overall health **✅****Healthy**
- LDAP/LDAPS across all DCs **healthy on FQDN endpoints for all 5 servers*
""";
        var normalized = MarkdownInputNormalizer.Normalize(markdown, options);

        Assert.Contains("- Overall health **healthy**", normalized, StringComparison.Ordinal);
        Assert.Contains("- Overall health ✅ **Healthy**", normalized, StringComparison.Ordinal);
        Assert.Contains("- LDAP/LDAPS across all DCs **healthy on FQDN endpoints for all 5 servers**", normalized, StringComparison.Ordinal);
    }

    [Fact]
    public void Presets_ChatTranscript_AlignsWithLegacyBridgeContract() {
        var options = MarkdownInputNormalizationPresets.CreateChatTranscript();

        Assert.True(options.NormalizeLooseStrongDelimiters);
        Assert.True(options.NormalizeTightStrongBoundaries);
        Assert.True(options.NormalizeOrderedListMarkerSpacing);
        Assert.True(options.NormalizeOrderedListParenMarkers);
        Assert.True(options.NormalizeOrderedListCaretArtifacts);
        Assert.True(options.NormalizeTightParentheticalSpacing);
        Assert.True(options.NormalizeNestedStrongDelimiters);
        Assert.True(options.NormalizeTightArrowStrongBoundaries);
        Assert.True(options.NormalizeTightColonSpacing);
        Assert.True(options.NormalizeWrappedSignalFlowStrongRuns);
        Assert.True(options.NormalizeCollapsedMetricChains);
        Assert.True(options.NormalizeHostLabelBulletArtifacts);
        Assert.True(options.NormalizeStandaloneHashHeadingSeparators);
        Assert.True(options.NormalizeBrokenTwoLineStrongLeadIns);
        Assert.True(options.NormalizeDanglingTrailingStrongListClosers);
        Assert.True(options.NormalizeMetricValueStrongRuns);
        Assert.False(options.NormalizeSoftWrappedStrongSpans);
        Assert.False(options.NormalizeBrokenStrongArrowLabels);
        Assert.False(options.NormalizeCompactFenceBodyBoundaries);
    }

    [Fact]
    public void Presets_ChatStrict_RepairsRepresentativeTranscriptArtifacts() {
        var options = MarkdownInputNormalizationPresets.CreateChatStrict();
        var markdown = """
## Wynik ogólny- **Replication:** wcześniej zdrowa ✅- **FSMO:** technicznie OK
Signal ->**Why it matters:**coverage
- Signal **No current failures -> **Why it matters:** transport/auth issues
Następny najlepszy krok:- **`ad_domain_controller_facts`**
```json{"log_name":"System"}
```
Use \`/act act_001\` now.
""";

        var normalized = MarkdownInputNormalizer.Normalize(markdown, options);

        Assert.Contains("## Wynik ogólny\n- **Replication:**", normalized);
        Assert.Contains("Why it matters:", normalized);
        Assert.Contains("```json\n{\"log_name\":\"System\"}", normalized);
        Assert.Contains("Use `/act act_001` now.", normalized);
    }

    [Fact]
    public void Presets_DocsLoose_RemainsConservative() {
        var options = MarkdownInputNormalizationPresets.CreateDocsLoose();

        Assert.True(options.NormalizeLooseStrongDelimiters);
        Assert.True(options.NormalizeTightStrongBoundaries);
        Assert.True(options.NormalizeOrderedListMarkerSpacing);
        Assert.True(options.NormalizeOrderedListParenMarkers);
        Assert.True(options.NormalizeOrderedListCaretArtifacts);
        Assert.True(options.NormalizeTightParentheticalSpacing);
        Assert.True(options.NormalizeNestedStrongDelimiters);
        Assert.False(options.NormalizeWrappedSignalFlowStrongRuns);
        Assert.False(options.NormalizeCollapsedMetricChains);
        Assert.False(options.NormalizeHostLabelBulletArtifacts);
        Assert.False(options.NormalizeStandaloneHashHeadingSeparators);
        Assert.False(options.NormalizeBrokenTwoLineStrongLeadIns);
        Assert.False(options.NormalizeDanglingTrailingStrongListClosers);
        Assert.False(options.NormalizeMetricValueStrongRuns);
        Assert.False(options.NormalizeTightColonSpacing);
        Assert.False(options.NormalizeBrokenStrongArrowLabels);
        Assert.False(options.NormalizeHeadingListBoundaries);
        Assert.False(options.NormalizeCompactFenceBodyBoundaries);
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
            NormalizeTightStrongBoundaries = true,
            NormalizeTightArrowStrongBoundaries = true,
            NormalizeBrokenStrongArrowLabels = true,
            NormalizeHeadingListBoundaries = true,
            NormalizeCompactStrongLabelListBoundaries = true,
            NormalizeCompactHeadingBoundaries = true,
            NormalizeColonListBoundaries = true,
            NormalizeCompactFenceBodyBoundaries = true
        };

        var markdown = """
```text
Use \`/act act_001\`
Status **Healthy**next
Signal ->**Why it matters:**coverage
- Signal **No current failures -> **Why it matters:** transport/auth issues
## Wynik ogólny- **Replication:** wcześniej zdrowa ✅- **FSMO:** technicznie OK
unexpected### Reason
Następny najlepszy krok:- **`ad_domain_controller_facts`**
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
    public void Normalize_InlineCodeLineBreaks_DoesNotChangeQuotedFencedCodeBlocks() {
        var options = new MarkdownInputNormalizationOptions {
            NormalizeInlineCodeSpanLineBreaks = true
        };

        var markdown = """
> ```ix-chart
> {"type":"bar","data":{"labels":["A"],"datasets":[{"label":"Count","data":[1]}]}}
> ```
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
