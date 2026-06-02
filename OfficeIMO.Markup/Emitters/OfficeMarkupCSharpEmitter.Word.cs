namespace OfficeIMO.Markup;

public sealed partial class OfficeMarkupCSharpEmitter {
    private static void EmitWordDocument(OfficeMarkupDocument document, OfficeMarkupEmitterOptions options, StringBuilder sb) {
        sb.AppendLine("using OfficeIMO.Word;");
        sb.AppendLine();
        sb.AppendLine($"using WordDocument document = WordDocument.Create({options.FilePathVariable});");
        foreach (var block in document.Blocks) {
            EmitWordBlock(block, "document", sb);
        }

        sb.AppendLine("document.Save();");
    }

    private static void EmitWordBlock(OfficeMarkupBlock block, string documentVariable, StringBuilder sb) {
        switch (block) {
            case OfficeMarkupHeadingBlock heading:
                sb.AppendLine($"{documentVariable}.AddParagraph({CsString(heading.Text)}).SetStyle(WordParagraphStyles.Heading{Math.Max(1, Math.Min(6, heading.Level))});");
                break;
            case OfficeMarkupParagraphBlock paragraph:
                sb.AppendLine($"{documentVariable}.AddParagraph({CsString(paragraph.Text)});");
                break;
            case OfficeMarkupListBlock list:
                foreach (var item in list.Items) {
                    sb.AppendLine($"{documentVariable}.AddParagraph({CsString(item.Text)});");
                }

                break;
            case OfficeMarkupPageBreakBlock:
                sb.AppendLine($"{documentVariable}.AddPageBreak();");
                break;
            case OfficeMarkupHeaderFooterBlock headerFooter when string.Equals(headerFooter.HeaderFooterKind, "header", StringComparison.OrdinalIgnoreCase):
                sb.AppendLine($"{documentVariable}.HeaderDefaultOrCreate.AddParagraph({CsString(headerFooter.Text)});");
                break;
            case OfficeMarkupHeaderFooterBlock headerFooter:
                sb.AppendLine($"{documentVariable}.FooterDefaultOrCreate.AddParagraph({CsString(headerFooter.Text)});");
                break;
            case OfficeMarkupTableOfContentsBlock toc:
                if (!string.IsNullOrWhiteSpace(toc.Title)) {
                    sb.AppendLine($"{documentVariable}.AddParagraph({CsString(toc.Title!)});");
                }

                sb.AppendLine($"{documentVariable}.AddTableOfContent(TableOfContentStyle.Template1, {toc.MinLevel ?? 1}, {toc.MaxLevel ?? 3});");
                break;
            case OfficeMarkupSectionBlock section:
                sb.AppendLine($"// Section: {CsString(section.Name ?? "section")}");
                foreach (var child in section.Blocks) {
                    EmitWordBlock(child, documentVariable, sb);
                }
                break;
            case OfficeMarkupImageBlock image:
                sb.AppendLine($"{documentVariable}.AddParagraph().AddImage({CsString(image.Source)});");
                break;
            case OfficeMarkupDiagramBlock diagram:
                sb.AppendLine($"// Render {diagram.Language} diagram to an image, then add it with document.AddParagraph().AddImage(...).");
                break;
            case OfficeMarkupChartBlock chart:
                EmitWordChart(chart, documentVariable, sb);
                break;
            case OfficeMarkupTableBlock table:
                sb.AppendLine($"var wordTable = {documentVariable}.AddTable({Math.Max(1, table.Rows.Count + (table.Headers.Count > 0 ? 1 : 0))}, {Math.Max(1, table.Headers.Count > 0 ? table.Headers.Count : table.Rows.Select(row => row.Count).DefaultIfEmpty(1).Max())});");
                sb.AppendLine("// Fill wordTable cells from the semantic table AST.");
                break;
            default:
                sb.AppendLine($"// {block.Kind}: {CsString(Describe(block))}");
                break;
        }
    }
}
