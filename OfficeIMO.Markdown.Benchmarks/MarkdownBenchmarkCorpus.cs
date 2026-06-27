using System.Text;

namespace OfficeIMO.Markdown.Benchmarks;

internal static class MarkdownBenchmarkCorpus {
    private static readonly IReadOnlyDictionary<string, string> Corpora = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) {
        ["PortableReadme"] = BuildPortableReadme(),
        ["Transcript"] = BuildTranscript(),
        ["TechnicalDoc"] = BuildTechnicalDoc(),
        ["RichAst"] = BuildRichAst(),
        ["LongNestedList"] = BuildLongNestedList(),
        ["LargeTable"] = BuildLargeTable(),
        ["NormalizationStress"] = BuildNormalizationStress()
    };

    public static IEnumerable<string> Names => Corpora.Keys;

    public static string Get(string name) => Corpora[name];

    private static string BuildPortableReadme() {
        var section = """
# OfficeIMO Markdown Overview

OfficeIMO.Markdown can build, parse, inspect, and render Markdown for chat, docs, and reports.

## Features

- Fluent builders for headings, paragraphs, tables, lists, and fenced code
- Typed document queries for headings, list items, top-level blocks, and descendants
- HTML rendering with anchor links and table-of-contents support

## Sample Table

| Area | Status | Notes |
| --- | --- | --- |
| Reader | Active | Typed document traversal |
| Renderer | Active | Portable HTML output |
| Docs | Improving | More examples and contract docs |

## Example

```csharp
var doc = MarkdownDoc.Create();
doc.H1("Status");
doc.P("Everything is working.");
```

Additional notes:

1. Keep README output readable.
2. Keep parser behavior predictable.
3. Keep portability in mind.

> This quote keeps the corpus close to a normal README.

""";

        return string.Concat(Enumerable.Repeat(section + Environment.NewLine, 12));
    }

    private static string BuildTranscript() {
        var section = """
## User

Please summarize the deployment status and list any blockers.

## Assistant

Deployment summary:

1. API rollout completed in staging.
2. Windows smoke tests are still running.
3. Package-mode validation passed locally.

### Notes

- Environment: `staging`
- Region: `westeurope`
- Follow-up: confirm package publication plan

```json
{
  "rollout": "staging",
  "result": "pending-final-check"
}
```

""";

        return string.Concat(Enumerable.Repeat(section + Environment.NewLine, 20));
    }

    private static string BuildTechnicalDoc() {
        var section = """
# Rendering Contract

## Normalization

The pipeline first normalizes transcript boundaries, then applies runtime-specific rendering options.

### Constraints

- Preserve literal fenced code blocks
- Preserve angle-bracket links
- Keep nested list structure stable

## Reference

| Component | Responsibility |
| --- | --- |
| Preparation | App-side transcript cleanup |
| Contract | Shared normalization across render/export |
| Runtime | Renderer and DOCX capability probing |

> [!NOTE]
> OfficeIMO-specific callouts are part of the default profile.

### Example

```markdown
> [!TIP]
> Keep package-mode validation in CI.
```

Trailing paragraph with **bold**, _emphasis_, `code`, and [a link](https://example.com).

""";

        return string.Concat(Enumerable.Repeat(section + Environment.NewLine, 16));
    }

    private static string BuildRichAst() {
        var section = """
# Incident Summary

> [!NOTE] Timeline
> First signal detected in the edge logs.
> Second signal came from the API health monitor.

Term A: First paragraph
  continued

  Second paragraph

Term B: Value

| Area | Status | Notes |
| --- | --- | --- |
| Edge | Active | Uses [docs](https://example.com/docs) |
| API | Pending | First line<br>Second line |

Trailing paragraph with **bold**, _emphasis_, `code`, and ![image](https://example.com/a.png).

[^ref]: Nested footnote body
  - item one
  - item two
""";

        return string.Concat(Enumerable.Repeat(section + Environment.NewLine, 18));
    }

    private static string BuildLongNestedList() {
        var builder = new StringBuilder();
        builder.AppendLine("# Long Operations Checklist");
        builder.AppendLine();

        for (int section = 1; section <= 18; section++) {
            builder.AppendLine($"## Section {section}");
            builder.AppendLine();
            for (int item = 1; item <= 18; item++) {
                builder.AppendLine($"- Check {section}.{item} starts with **status** and a [runbook](https://example.com/runbooks/{section}/{item}).");
                builder.AppendLine($"  - Evidence path `{section:D2}-{item:D2}` is present.");
                builder.AppendLine($"  - Nested note {item} includes a task marker:");
                builder.AppendLine($"    - [ ] capture result for check {section}.{item}");
                builder.AppendLine($"    - [x] preserve existing remediation text");
            }

            builder.AppendLine();
        }

        return builder.ToString();
    }

    private static string BuildLargeTable() {
        var builder = new StringBuilder();
        builder.AppendLine("# Large Compatibility Table");
        builder.AppendLine();
        builder.AppendLine("| Area | Capability | Status | Evidence | Notes |");
        builder.AppendLine("| --- | --- | --- | --- | --- |");

        for (int row = 1; row <= 420; row++) {
            builder
                .Append("| Markdown | Capability ")
                .Append(row)
                .Append(" | Partial | [case ")
                .Append(row)
                .Append("](https://example.com/cases/")
                .Append(row)
                .Append(") | Uses `code`, **strong**, _emphasis_, and escaped pipe \\| text |")
                .AppendLine();
        }

        builder.AppendLine();
        builder.AppendLine("Final paragraph after the table keeps paragraph-boundary costs visible.");
        return builder.ToString();
    }

    private static string BuildNormalizationStress() {
        var section = """
previous shutdown was unexpected### Reason

- Signal **Healthy baseline exists now** ->**Why it matters:**missing coverage ->**Next action:**review defaults.
- Signal **Catalog count includes hidden/disabled/deprecated rules -> **Why it matters:**external/custom rules can drift ->**Next action:**compare inventories.**

Następny najlepszy krok:- **`ad_domain_controller_facts`**

1) First check
2.^ **Delegation risk audit**
3. **Deleted object remnants**(SID left in ACL path)

Command: `Get-ADUser(SIDHistory)`

""";

        return string.Concat(Enumerable.Repeat(section + Environment.NewLine, 80));
    }
}
