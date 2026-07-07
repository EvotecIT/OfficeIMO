using System.Text;

namespace OfficeIMO.Markdown.Benchmarks;

internal static class HtmlToMarkdownBenchmarkCorpus {
    private static readonly IReadOnlyDictionary<string, string> Corpora = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) {
        ["Article"] = BuildArticle(),
        ["LargeArticle"] = BuildLargeArticle(),
        ["Table"] = BuildTable(),
        ["NestedLists"] = BuildNestedLists(),
        ["MixedDocument"] = BuildMixedDocument()
    };

    public static IEnumerable<string> Names => Corpora.Keys;

    public static string Get(string name) => Corpora[name];

    public static string GetExpectedFragment(string name) => name switch {
        "Article" => "OfficeIMO Markdown Notes",
        "LargeArticle" => "Section 120",
        "Table" => "Capability 240",
        "NestedLists" => "Check 40.8",
        "MixedDocument" => "Release Report",
        _ => throw new ArgumentOutOfRangeException(nameof(name), name, "Unknown benchmark corpus.")
    };

    private static string BuildArticle() {
        const string section = """
<article>
  <h1>OfficeIMO Markdown Notes</h1>
  <p>OfficeIMO can convert <strong>HTML</strong> into <a href="https://example.com/docs">Markdown</a>.</p>
  <p>The converter keeps code such as <code>MarkdownDoc.Create()</code> readable.</p>
  <blockquote><p>Representative fixtures should look like real docs.</p></blockquote>
  <ul>
    <li>Parse the document</li>
    <li>Build a typed model</li>
    <li>Render Markdown</li>
  </ul>
</article>
""";

        return section;
    }

    private static string BuildLargeArticle() {
        var builder = new StringBuilder();
        builder.AppendLine("<main>");
        for (int i = 1; i <= 120; i++) {
            builder.AppendLine("<section>");
            builder.Append("<h2>Section ").Append(i).AppendLine("</h2>");
            builder.Append("<p>Paragraph ").Append(i).Append(" links to <a href=\"https://example.com/docs/")
                .Append(i)
                .Append("\">https://example.com/docs/")
                .Append(i)
                .AppendLine("</a> and contains <strong>important</strong> details.</p>");
            builder.AppendLine("<p>Second paragraph includes <em>emphasis</em>, <code>inline code</code>, and normal text.</p>");
            builder.AppendLine("</section>");
        }

        builder.AppendLine("</main>");
        return builder.ToString();
    }

    private static string BuildTable() {
        var builder = new StringBuilder();
        builder.AppendLine("<table>");
        builder.AppendLine("<thead><tr><th>Area</th><th>Status</th><th>Notes</th></tr></thead>");
        builder.AppendLine("<tbody>");
        for (int row = 1; row <= 240; row++) {
            builder.Append("<tr><td>Capability ")
                .Append(row)
                .Append("</td><td>Active</td><td>Uses <strong>typed</strong> conversion and <a href=\"https://example.com/cases/")
                .Append(row)
                .Append("\">case ")
                .Append(row)
                .AppendLine("</a>.</td></tr>");
        }

        builder.AppendLine("</tbody>");
        builder.AppendLine("</table>");
        return builder.ToString();
    }

    private static string BuildNestedLists() {
        var builder = new StringBuilder();
        builder.AppendLine("<article><h1>Checklist</h1><ol>");
        for (int section = 1; section <= 40; section++) {
            builder.Append("<li>Section ")
                .Append(section)
                .AppendLine("<ul>");
            for (int item = 1; item <= 8; item++) {
                builder.Append("<li><strong>Check ")
                    .Append(section)
                    .Append('.')
                    .Append(item)
                    .Append("</strong> with <a href=\"https://example.com/runbooks/")
                    .Append(section)
                    .Append('/')
                    .Append(item)
                    .AppendLine("\">runbook</a></li>");
            }

            builder.AppendLine("</ul></li>");
        }

        builder.AppendLine("</ol></article>");
        return builder.ToString();
    }

    private static string BuildMixedDocument() {
        return """
<article>
  <h1>Release Report</h1>
  <p><a href="https://example.com">https://example.com</a></p>
  <figure>
    <img src="https://example.com/logo.png" alt="Logo" width="256" height="128" />
    <figcaption>Project logo</figcaption>
  </figure>
  <details>
    <summary>Deployment details</summary>
    <p>Staging is complete and production is pending.</p>
  </details>
  <pre><code class="language-csharp">var doc = MarkdownDoc.Create();
doc.H1("Status");</code></pre>
  <table>
    <tr><th>Name</th><th>Value</th></tr>
    <tr><td>Runtime</td><td>.NET</td></tr>
    <tr><td>Mode</td><td>Benchmark</td></tr>
  </table>
  <custom-widget data-id="42"><strong>Custom</strong> payload</custom-widget>
</article>
""";
    }
}
