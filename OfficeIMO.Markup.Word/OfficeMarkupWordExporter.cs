using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using ImageSharpColor = SixLabors.ImageSharp.Color;

namespace OfficeIMO.Markup.Word;

public sealed class OfficeMarkupWordExporter {
    private static readonly ImageSharpColor[] ChartColors = new[] {
        ImageSharpColor.CornflowerBlue,
        ImageSharpColor.SeaGreen,
        ImageSharpColor.IndianRed,
        ImageSharpColor.Goldenrod,
        ImageSharpColor.MediumPurple,
        ImageSharpColor.DarkCyan
    };

    public void Export(OfficeMarkupDocument document, OfficeMarkupWordExportOptions options) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        if (document.Profile != OfficeMarkupProfile.Document) {
            throw new InvalidOperationException("Word export requires the Document OfficeIMO markup profile.");
        }

        if (string.IsNullOrWhiteSpace(options.OutputPath)) {
            throw new InvalidOperationException("Word export requires an output path.");
        }

        var directory = Path.GetDirectoryName(Path.GetFullPath(options.OutputPath));
        if (!string.IsNullOrEmpty(directory)) {
            Directory.CreateDirectory(directory);
        }

        using var word = WordDocument.Create(options.OutputPath);
        var context = new WordExportContext(word, options);
        foreach (var block in document.Blocks) {
            ExportBlock(context, block);
        }

        word.Save();
    }

    private static void ExportBlock(WordExportContext context, OfficeMarkupBlock block) {
        switch (block) {
            case OfficeMarkupHeadingBlock heading:
                AddHeading(context, heading);
                break;
            case OfficeMarkupParagraphBlock paragraph:
                context.AddParagraph(paragraph.Text);
                break;
            case OfficeMarkupListBlock list:
                AddList(context, list);
                break;
            case OfficeMarkupCodeBlock code:
                AddCode(context, code);
                break;
            case OfficeMarkupImageBlock image:
                AddImage(context, image);
                break;
            case OfficeMarkupTableBlock table:
                AddTable(context, table);
                break;
            case OfficeMarkupDiagramBlock diagram:
                AddDiagram(context, diagram);
                break;
            case OfficeMarkupPageBreakBlock:
                context.Document.AddPageBreak();
                break;
            case OfficeMarkupSectionBlock section:
                AddSection(context, section);
                break;
            case OfficeMarkupHeaderFooterBlock headerFooter:
                AddHeaderFooter(context, headerFooter);
                break;
            case OfficeMarkupTableOfContentsBlock toc:
                AddTableOfContents(context, toc);
                break;
            case OfficeMarkupChartBlock chart:
                AddChart(context, chart);
                break;
            case OfficeMarkupExtensionBlock extension when context.Options.IncludeUnsupportedBlocksAsText:
                context.AddParagraph(extension.Body.Trim());
                break;
            case OfficeMarkupRawMarkdownBlock raw when context.Options.IncludeUnsupportedBlocksAsText:
                context.AddParagraph(raw.Markdown);
                break;
        }
    }

    private static void AddHeading(WordExportContext context, OfficeMarkupHeadingBlock heading) {
        var paragraph = context.AddParagraph(heading.Text);
        paragraph.Style = HeadingStyle(heading.Level);
    }

    private static void AddList(WordExportContext context, OfficeMarkupListBlock list) {
        if (context.CurrentSection != null) {
            for (var index = 0; index < list.Items.Count; index++) {
                var item = list.Items[index];
                var prefix = list.Ordered ? $"{list.Start + index}. " : "- ";
                context.CurrentSection.AddParagraph(prefix + item.Text);
            }

            return;
        }

        var wordList = list.Ordered ? context.Document.AddListNumbered() : context.Document.AddListBulleted();
        foreach (var item in list.Items) {
            wordList.AddItem(item.Text);
        }
    }

    private static void AddCode(WordExportContext context, OfficeMarkupCodeBlock code) {
        var label = string.IsNullOrWhiteSpace(code.Language) ? "code" : code.Language;
        context.AddParagraph(label + ":");
        context.AddParagraph(code.Content);
    }

    private static void AddImage(WordExportContext context, OfficeMarkupImageBlock image) {
        var path = ResolvePath(context.Options, image.Source);
        if (!File.Exists(path)) {
            if (context.Options.IncludeUnsupportedBlocksAsText) {
                context.AddParagraph($"Image: {image.Source}");
            }

            return;
        }

        context.AddParagraph().AddImage(path, image.Width, image.Height, description: image.Alt ?? image.Title ?? string.Empty);
    }

    private static void AddTable(WordExportContext context, OfficeMarkupTableBlock table) {
        var rowCount = table.Rows.Count + (table.Headers.Count > 0 ? 1 : 0);
        var columnCount = Math.Max(table.Headers.Count, table.Rows.Select(row => row.Count).DefaultIfEmpty(0).Max());
        if (rowCount <= 0 || columnCount <= 0) {
            return;
        }

        var wordTable = context.AddTable(rowCount, columnCount);
        var rowIndex = 0;
        if (table.Headers.Count > 0) {
            for (var column = 0; column < table.Headers.Count; column++) {
                wordTable.Rows[rowIndex].Cells[column].Paragraphs[0].Text = table.Headers[column];
            }

            wordTable.RepeatHeaderRowAtTheTopOfEachPage = true;
            rowIndex++;
        }

        foreach (var row in table.Rows) {
            for (var column = 0; column < row.Count; column++) {
                wordTable.Rows[rowIndex].Cells[column].Paragraphs[0].Text = row[column];
            }

            rowIndex++;
        }
    }

    private static void AddDiagram(WordExportContext context, OfficeMarkupDiagramBlock diagram) {
        context.AddParagraph($"{diagram.Language} diagram:");
        context.AddParagraph(diagram.Content);
    }

    private static void AddSection(WordExportContext context, OfficeMarkupSectionBlock section) {
        var wordSection = context.Document.AddSection(ParseSectionMark(section.Attributes));
        context.CurrentSection = wordSection;
        if (!string.IsNullOrWhiteSpace(section.Name)) {
            wordSection.AddParagraph(section.Name!).Style = WordParagraphStyles.Heading1;
        }

        foreach (var block in section.Blocks) {
            ExportBlock(context, block);
        }
    }

    private static void AddHeaderFooter(WordExportContext context, OfficeMarkupHeaderFooterBlock headerFooter) {
        var section = context.CurrentSection ?? context.Document.Sections.FirstOrDefault();
        if (section == null) {
            context.Document.AddHeadersAndFooters();
            section = context.Document.Sections.First();
        }

        var type = ParseHeaderFooterType(headerFooter.Attributes);
        if (string.Equals(headerFooter.HeaderFooterKind, "footer", StringComparison.OrdinalIgnoreCase)) {
            section.GetOrCreateFooter(type).AddParagraph(headerFooter.Text);
        } else {
            section.GetOrCreateHeader(type).AddParagraph(headerFooter.Text);
        }
    }

    private static void AddTableOfContents(WordExportContext context, OfficeMarkupTableOfContentsBlock toc) {
        var tableOfContent = context.Document.AddTableOfContent(
            TableOfContentStyle.Template1,
            Math.Max(1, toc.MinLevel ?? 1),
            Math.Max(toc.MinLevel ?? 1, toc.MaxLevel ?? 3));
        if (!string.IsNullOrWhiteSpace(toc.Title)) {
            tableOfContent.Text = toc.Title!;
        }

        context.Document.Settings.UpdateFieldsOnOpen = true;
    }

    private static void AddChart(WordExportContext context, OfficeMarkupChartBlock chart) {
        if (chart.Data.Count <= 1) {
            if (context.Options.IncludeUnsupportedBlocksAsText) {
                context.AddParagraph($"Chart: {chart.Title ?? chart.ChartType}");
            }

            return;
        }

        var data = ParseChartData(chart);
        var wordChart = context.Document.AddChart(
            chart.Title ?? string.Empty,
            roundedCorners: false,
            width: GetInt(chart.Attributes, "width") ?? GetInt(chart.Attributes, "w") ?? context.Options.DefaultChartWidthPixels,
            height: GetInt(chart.Attributes, "height") ?? GetInt(chart.Attributes, "h") ?? context.Options.DefaultChartHeightPixels);

        switch (Normalize(chart.ChartType)) {
            case "line":
                wordChart.AddChartAxisX(data.Categories);
                for (var index = 0; index < data.Series.Count; index++) {
                    wordChart.AddLine(data.Series[index].Name, data.Series[index].Values, ChartColors[index % ChartColors.Length]);
                }

                break;
            case "area":
                wordChart.AddCategories(data.Categories);
                for (var index = 0; index < data.Series.Count; index++) {
                    wordChart.AddArea(data.Series[index].Name, data.Series[index].Values, ChartColors[index % ChartColors.Length]);
                }

                break;
            case "pie":
            case "doughnut":
            case "donut":
                var firstSeries = data.Series.FirstOrDefault();
                if (firstSeries != null) {
                    for (var index = 0; index < data.Categories.Count && index < firstSeries.Values.Count; index++) {
                        wordChart.AddPie(data.Categories[index], firstSeries.Values[index]);
                    }
                }

                break;
            default:
                wordChart.AddCategories(data.Categories);
                for (var index = 0; index < data.Series.Count; index++) {
                    wordChart.AddBar(data.Series[index].Name, data.Series[index].Values, ChartColors[index % ChartColors.Length]);
                }

                break;
        }
    }

    private static WordParagraphStyles HeadingStyle(int level) =>
        level switch {
            <= 1 => WordParagraphStyles.Heading1,
            2 => WordParagraphStyles.Heading2,
            3 => WordParagraphStyles.Heading3,
            4 => WordParagraphStyles.Heading4,
            5 => WordParagraphStyles.Heading5,
            6 => WordParagraphStyles.Heading6,
            7 => WordParagraphStyles.Heading7,
            8 => WordParagraphStyles.Heading8,
            _ => WordParagraphStyles.Heading9
        };

    private static SectionMarkValues? ParseSectionMark(IDictionary<string, string> attributes) {
        var value = GetAttribute(attributes, "break") ?? GetAttribute(attributes, "sectionBreak");
        return Normalize(value) switch {
            "continuous" => SectionMarkValues.Continuous,
            "evenpage" => SectionMarkValues.EvenPage,
            "oddpage" => SectionMarkValues.OddPage,
            _ => SectionMarkValues.NextPage
        };
    }

    private static HeaderFooterValues ParseHeaderFooterType(IDictionary<string, string> attributes) {
        var value = Normalize(GetAttribute(attributes, "type") ?? GetAttribute(attributes, "kind"));
        return value switch {
            "first" or "firstpage" => HeaderFooterValues.First,
            "even" or "evenpage" => HeaderFooterValues.Even,
            _ => HeaderFooterValues.Default
        };
    }

    private static ChartData ParseChartData(OfficeMarkupChartBlock chart) {
        var header = chart.Data[0];
        var categories = chart.Data.Skip(1).Select(row => row.Count > 0 ? row[0] : string.Empty).ToList();
        var series = new List<ChartSeries>();
        for (var column = 1; column < header.Count; column++) {
            var values = new List<double>();
            foreach (var row in chart.Data.Skip(1)) {
                values.Add(row.Count > column && double.TryParse(row[column], NumberStyles.Any, CultureInfo.InvariantCulture, out var value) ? value : 0d);
            }

            series.Add(new ChartSeries(header[column], values));
        }

        return new ChartData(categories, series);
    }

    private static string ResolvePath(OfficeMarkupWordExportOptions options, string source) {
        if (Path.IsPathRooted(source) || string.IsNullOrWhiteSpace(options.BaseDirectory)) {
            return source;
        }

        return Path.Join(options.BaseDirectory!, source);
    }

    private static int? GetInt(IDictionary<string, string> attributes, string name) {
        var value = GetAttribute(attributes, name);
        if (string.IsNullOrWhiteSpace(value) || value!.IndexOf('%') >= 0) {
            return null;
        }

        return int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var result) ? result : null;
    }

    private static string? GetAttribute(IDictionary<string, string> attributes, string name) =>
        attributes.TryGetValue(name, out var value) ? value : null;

    private static string Normalize(string? value) =>
        (value ?? string.Empty).Trim().Replace("-", string.Empty).Replace("_", string.Empty).Replace(" ", string.Empty).ToLowerInvariant();

    private sealed class WordExportContext {
        public WordExportContext(WordDocument document, OfficeMarkupWordExportOptions options) {
            Document = document;
            Options = options;
        }

        public WordDocument Document { get; }
        public OfficeMarkupWordExportOptions Options { get; }
        public WordSection? CurrentSection { get; set; }

        public WordParagraph AddParagraph(string text = "") =>
            CurrentSection != null ? CurrentSection.AddParagraph(text) : Document.AddParagraph(text);

        public WordTable AddTable(int rows, int columns) =>
            CurrentSection != null ? CurrentSection.AddTable(rows, columns) : Document.AddTable(rows, columns);
    }

    private sealed class ChartData {
        public ChartData(List<string> categories, List<ChartSeries> series) {
            Categories = categories;
            Series = series;
        }

        public List<string> Categories { get; }
        public List<ChartSeries> Series { get; }
    }

    private sealed class ChartSeries {
        public ChartSeries(string name, List<double> values) {
            Name = name;
            Values = values;
        }

        public string Name { get; }
        public List<double> Values { get; }
    }
}
