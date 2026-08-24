using Markdig;
using Markdig.Extensions.Tables;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using OfficeMarkdownReaderOptions = OfficeIMO.Markdown.MarkdownReaderOptions;
using System.Security.Cryptography;
using System.Text;

namespace OfficeIMO.Markup.Benchmarks;

internal static class OfficeMarkupBenchmarkValidation {
    internal static readonly MarkdownPipeline MarkdigPipeline = new MarkdownPipelineBuilder().UsePipeTables().Build();
    internal static readonly OfficeMarkupParserOptions OfficeOptions = CreateOfficeOptions();

    private static OfficeMarkupParserOptions CreateOfficeOptions() {
        OfficeMarkdownReaderOptions markdownOptions = OfficeMarkdownReaderOptions.CreateCommonMarkProfile();
        markdownOptions.Tables = true;
        return new OfficeMarkupParserOptions {
            Profile = OfficeMarkupProfile.Common,
            Validate = false,
            MarkdownOptions = markdownOptions
        };
    }

    internal static void ValidateAll() {
        foreach (string scale in OfficeMarkupBenchmarkCorpus.Scales) {
            OfficeMarkupBenchmarkFixture fixture = OfficeMarkupBenchmarkCorpus.Get(scale);
            (OfficeMarkupParseResult office, MarkdownDocument markdig) = Validate(fixture);
            Console.WriteLine(
                $"{scale,-6} input {Encoding.UTF8.GetByteCount(fixture.Source),10:N0} bytes | " +
                $"events {CreateOfficeSnapshot(office).EventCount,8:N0} | sections {fixture.SectionCount,6:N0} | " +
                $"records {fixture.RecordCount,7:N0} | digest {CreateMarkdigSnapshot(markdig).Digest[..12]}");
        }
    }

    internal static (OfficeMarkupParseResult Office, MarkdownDocument Markdig) Validate(OfficeMarkupBenchmarkFixture fixture) {
        OfficeMarkupParseResult office = OfficeMarkupParser.Parse(fixture.Source, OfficeOptions);
        MarkdownDocument markdig = Markdig.Markdown.Parse(fixture.Source, MarkdigPipeline);
        if (office.HasErrors) throw new InvalidOperationException(fixture.Scale + " produced Office Markup errors.");
        SemanticSnapshot officeSnapshot = CreateOfficeSnapshot(office);
        SemanticSnapshot markdigSnapshot = CreateMarkdigSnapshot(markdig);
        if (officeSnapshot != markdigSnapshot) {
            throw new InvalidOperationException(
                $"{fixture.Scale} semantic snapshots differ: OfficeIMO {officeSnapshot}, Markdig {markdigSnapshot}.");
        }
        if (officeSnapshot.HeadingCount != fixture.SectionCount + 1
            || officeSnapshot.ListCount != fixture.SectionCount
            || officeSnapshot.TableCount != fixture.SectionCount) {
            throw new InvalidOperationException(fixture.Scale + " did not produce the expected semantic structure.");
        }
        return (office, markdig);
    }

    internal static SemanticSnapshot CreateOfficeSnapshot(OfficeMarkupParseResult result) {
        var events = new StringBuilder();
        int headings = 0;
        int paragraphs = 0;
        int lists = 0;
        int tables = 0;
        int listItems = 0;
        int tableRows = 0;
        foreach (OfficeMarkupBlock block in result.Document.Blocks) {
            switch (block) {
                case OfficeMarkupHeadingBlock heading:
                    headings++;
                    Append(events, "H", heading.Level.ToString(), heading.Text);
                    break;
                case OfficeMarkupParagraphBlock paragraph:
                    paragraphs++;
                    Append(events, "P", paragraph.Text);
                    break;
                case OfficeMarkupListBlock list:
                    lists++;
                    Append(events, "L", list.Ordered ? "1" : "0", list.Start.ToString());
                    foreach (OfficeMarkupListItem item in list.Items) {
                        listItems++;
                        Append(events, "I", item.Text);
                    }
                    break;
                case OfficeMarkupTableBlock table:
                    tables++;
                    Append(events, "T", string.Join("\u001f", table.Headers));
                    foreach (IReadOnlyList<string> row in table.Rows) {
                        tableRows++;
                        Append(events, "R", string.Join("\u001f", row));
                    }
                    break;
                default:
                    throw new InvalidOperationException("Unsupported Office Markup benchmark block: " + block.Kind);
            }
        }
        return Snapshot(events, headings, paragraphs, lists, listItems, tables, tableRows);
    }

    internal static SemanticSnapshot CreateMarkdigSnapshot(MarkdownDocument document) {
        var events = new StringBuilder();
        int headings = 0;
        int paragraphs = 0;
        int lists = 0;
        int tables = 0;
        int listItems = 0;
        int tableRows = 0;
        foreach (Block block in document) {
            switch (block) {
                case HeadingBlock heading:
                    headings++;
                    Append(events, "H", heading.Level.ToString(), PlainText(heading.Inline));
                    break;
                case ParagraphBlock paragraph:
                    paragraphs++;
                    Append(events, "P", PlainText(paragraph.Inline));
                    break;
                case ListBlock list:
                    lists++;
                    Append(events, "L", list.IsOrdered ? "1" : "0", list.OrderedStart ?? "1");
                    foreach (ListItemBlock item in list.OfType<ListItemBlock>()) {
                        listItems++;
                        ParagraphBlock paragraph = item.OfType<ParagraphBlock>().First();
                        Append(events, "I", PlainText(paragraph.Inline));
                    }
                    break;
                case Table table:
                    tables++;
                    TableRow[] rows = table.OfType<TableRow>().ToArray();
                    Append(events, "T", string.Join("\u001f", Cells(rows[0])));
                    for (int index = 1; index < rows.Length; index++) {
                        tableRows++;
                        Append(events, "R", string.Join("\u001f", Cells(rows[index])));
                    }
                    break;
                default:
                    throw new InvalidOperationException("Unsupported Markdig benchmark block: " + block.GetType().Name);
            }
        }
        return Snapshot(events, headings, paragraphs, lists, listItems, tables, tableRows);
    }

    private static IEnumerable<string> Cells(TableRow row) =>
        row.OfType<TableCell>().Select(cell => PlainText(cell.Descendants<ParagraphBlock>().Single().Inline));

    private static string PlainText(ContainerInline? container) {
        if (container == null) return string.Empty;
        var text = new StringBuilder();
        AppendInline(container.FirstChild, text);
        return text.ToString();
    }

    private static void AppendInline(Inline? inline, StringBuilder text) {
        while (inline != null) {
            switch (inline) {
                case LiteralInline literal:
                    text.Append(literal.Content.ToString());
                    break;
                case CodeInline code:
                    text.Append(code.Content);
                    break;
                case LineBreakInline:
                    text.Append('\n');
                    break;
                case ContainerInline nested:
                    AppendInline(nested.FirstChild, text);
                    break;
            }
            inline = inline.NextSibling;
        }
    }

    private static void Append(StringBuilder events, params string[] values) {
        for (int index = 0; index < values.Length; index++) {
            if (index > 0) events.Append('\u001e');
            events.Append(NormalizeWhitespace(values[index]));
        }
        events.Append('\n');
    }

    private static string NormalizeWhitespace(string value) {
        var normalized = new StringBuilder(value.Length);
        bool pendingSpace = false;
        for (int index = 0; index < value.Length; index++) {
            if (char.IsWhiteSpace(value[index])) {
                pendingSpace = normalized.Length > 0;
                continue;
            }
            if (pendingSpace) normalized.Append(' ');
            normalized.Append(value[index]);
            pendingSpace = false;
        }
        return normalized.ToString();
    }

    private static SemanticSnapshot Snapshot(
        StringBuilder events,
        int headings,
        int paragraphs,
        int lists,
        int listItems,
        int tables,
        int tableRows) {
        string canonicalEvents = events.ToString();
        byte[] bytes = Encoding.UTF8.GetBytes(canonicalEvents);
        return new SemanticSnapshot(
            headings,
            paragraphs,
            lists,
            listItems,
            tables,
            tableRows,
            headings + paragraphs + lists + listItems + tables + tableRows,
            Convert.ToHexString(SHA256.HashData(bytes)));
    }
}

internal sealed record SemanticSnapshot(
    int HeadingCount,
    int ParagraphCount,
    int ListCount,
    int ListItemCount,
    int TableCount,
    int TableRowCount,
    int EventCount,
    string Digest);
