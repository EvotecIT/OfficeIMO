using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using OfficeIMO.Html;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.OneNote.Html;

/// <summary>Imports prepared ordinary HTML into typed offline OneNote models.</summary>
public static class HtmlOneNoteConverterExtensions {
    private const string ComponentName = "OfficeIMO.OneNote.Html";

    /// <summary>Imports HTML as a OneNote section or throws when an error diagnostic is produced.</summary>
    public static OneNoteSection ToOneNoteSection(this HtmlConversionDocument document, HtmlToOneNoteOptions? options = null) =>
        Require(document.ToOneNoteSectionResult(options));

    /// <summary>Imports HTML as a OneNote section with structured diagnostics and counters.</summary>
    public static HtmlToOneNoteSectionResult ToOneNoteSectionResult(this HtmlConversionDocument document, HtmlToOneNoteOptions? options = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        HtmlToOneNoteOptions resolved = options?.Clone() ?? new HtmlToOneNoteOptions();
        resolved.Limits.Validate();
        var section = new OneNoteSection { Name = CleanName(resolved.SectionName, "Imported") };
        var result = new HtmlToOneNoteSectionResult(section);
        IHtmlDocument adapterDocument = document.CreateDocumentForConversion(HtmlCssMediaContext.Screen);
        foreach (HtmlDiagnostic diagnostic in document.Diagnostics) result.AddImportDiagnostic(diagnostic);
        ImportPages(adapterDocument, section, resolved, result);
        return result;
    }

    /// <summary>Imports HTML as a single-section OneNote notebook or throws on conversion errors.</summary>
    public static OneNoteNotebook ToOneNoteNotebook(this HtmlConversionDocument document, HtmlToOneNoteOptions? options = null) =>
        Require(document.ToOneNoteNotebookResult(options));

    /// <summary>Imports HTML as a single-section OneNote notebook with structured evidence.</summary>
    public static HtmlToOneNoteNotebookResult ToOneNoteNotebookResult(this HtmlConversionDocument document, HtmlToOneNoteOptions? options = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        HtmlToOneNoteOptions resolved = options?.Clone() ?? new HtmlToOneNoteOptions();
        HtmlToOneNoteSectionResult sectionResult = document.ToOneNoteSectionResult(resolved);
        var notebook = new OneNoteNotebook { Name = CleanName(resolved.NotebookName, "Imported") };
        notebook.Sections.Add(sectionResult.Value);
        var result = new HtmlToOneNoteNotebookResult(notebook) {
            Sections = 1,
            Pages = sectionResult.Pages,
            Elements = sectionResult.Elements,
            Tables = sectionResult.Tables,
            Images = sectionResult.Images
        };
        foreach (HtmlDiagnostic diagnostic in sectionResult.Report.Diagnostics) result.AddImportDiagnostic(diagnostic);
        return result;
    }

    private static void ImportPages(
        IHtmlDocument document,
        OneNoteSection target,
        HtmlToOneNoteOptions options,
        HtmlToOneNoteSectionResult result) {
        var budget = new HtmlImportBudget(options.Limits);
        foreach (HtmlGenericSectionProjection projection in HtmlGenericDocumentProjector.CreateSections(document)) {
            if (!budget.TryReserveSemanticContainer(out string containerLimit)) {
                Add(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
                    "Additional HTML sections were omitted because the shared page limit was reached.",
                    HtmlDiagnosticSeverity.Error, HtmlConversionLossKind.Omission, containerLimit);
                break;
            }

            var page = new OneNotePage { Title = CleanName(projection.Title, "Imported") };
            var outline = new OneNoteOutline();
            foreach (IElement block in HtmlGenericDocumentProjector.EnumerateBlocks(projection)) {
                ImportBlock(block, outline.Children, options, result, budget);
            }
            if (outline.Children.Count > 0) page.Outlines.Add(outline);
            target.Pages.Add(page);
            result.Pages++;
        }
    }

    private static void ImportBlock(
        IElement block,
        IList<OneNoteElement> target,
        HtmlToOneNoteOptions options,
        HtmlToOneNoteSectionResult result,
        HtmlImportBudget budget) {
        if (HtmlGenericDocumentProjector.IsTable(block)) {
            ImportTable(block, target, result, budget);
            return;
        }
        if (HtmlGenericDocumentProjector.IsImage(block)) {
            if (options.ImportImages) ImportImage(block, target, result, budget);
            return;
        }
        if (string.Equals(block.LocalName, "ul", StringComparison.OrdinalIgnoreCase)
            || string.Equals(block.LocalName, "ol", StringComparison.OrdinalIgnoreCase)) {
            bool ordered = string.Equals(block.LocalName, "ol", StringComparison.OrdinalIgnoreCase);
            foreach (IElement item in block.Children.Where(child => string.Equals(child.LocalName, "li", StringComparison.OrdinalIgnoreCase))) {
                OneNoteParagraph? paragraph = CreateParagraph(item, result, budget);
                if (paragraph == null) break;
                paragraph.List = new OneNoteListInfo { Ordered = ordered, Level = 0 };
                target.Add(paragraph);
            }
            return;
        }
        OneNoteParagraph? text = CreateParagraph(block, result, budget);
        if (text != null) target.Add(text);
    }

    private static OneNoteParagraph? CreateParagraph(
        IElement source,
        HtmlToOneNoteSectionResult result,
        HtmlImportBudget budget) {
        string plainText = HtmlGenericDocumentProjector.GetBlockText(source);
        if (plainText.Length == 0) return null;
        if (!budget.TryReserveShape(out string shapeLimit)) {
            Add(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
                "Additional HTML blocks were omitted because the shared element limit was reached.",
                HtmlDiagnosticSeverity.Warning, HtmlConversionLossKind.Omission, shapeLimit);
            return null;
        }
        if (!budget.IsMetadataWithinLimit(plainText, out string metadataLimit)) {
            Add(result, HtmlConversionDiagnosticCodes.SemanticMetadataLimitExceeded,
                "An HTML text block was omitted because it exceeded the shared field limit.",
                HtmlDiagnosticSeverity.Warning, HtmlConversionLossKind.Omission, metadataLimit);
            return null;
        }

        var paragraph = new OneNoteParagraph();
        AppendRuns(source, paragraph, default);
        if (paragraph.Runs.Count == 0) paragraph.Runs.Add(new OneNoteTextRun { Text = plainText });
        TrimRuns(paragraph);
        if (HtmlGenericDocumentProjector.IsHeading(source)) {
            paragraph.Style.StyleId = "Heading" + source.LocalName.Substring(1);
        }
        result.Elements++;
        return paragraph;
    }

    private static void ImportTable(
        IElement source,
        IList<OneNoteElement> target,
        HtmlToOneNoteSectionResult result,
        HtmlImportBudget budget) {
        string tableLimit = string.Empty;
        string shapeLimit = string.Empty;
        if (!budget.TryReserveTable(out tableLimit) || !budget.TryReserveShape(out shapeLimit)) {
            Add(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
                "An HTML table was omitted because the shared import limit was reached.",
                HtmlDiagnosticSeverity.Warning, HtmlConversionLossKind.Omission,
                tableLimit.Length > 0 ? tableLimit : shapeLimit);
            return;
        }
        var table = new OneNoteTable { BordersVisible = true };
        int cells = 0;
        int maxTableCells = budget.Limits.MaxTableCells;
        foreach (IElement rowElement in DirectRows(source)) {
            var row = new OneNoteTableRow();
            foreach (IElement cellElement in rowElement.Children.Where(IsTableCell)) {
                if (++cells > maxTableCells) {
                    Add(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
                        "Remaining HTML table cells were omitted because the configured table limit was reached.",
                        HtmlDiagnosticSeverity.Warning, HtmlConversionLossKind.Omission,
                        "limit=" + maxTableCells);
                    break;
                }
                var cell = new OneNoteTableCell();
                OneNoteParagraph? paragraph = CreateParagraph(cellElement, result, budget);
                if (paragraph != null) cell.Content.Add(paragraph);
                row.Cells.Add(cell);
            }
            if (row.Cells.Count > 0) table.Rows.Add(row);
            if (cells >= maxTableCells) break;
        }
        if (table.Rows.Count == 0) return;
        target.Add(table);
        result.Elements++;
        result.Tables++;
    }

    private static void ImportImage(
        IElement imageElement,
        IList<OneNoteElement> target,
        HtmlToOneNoteSectionResult result,
        HtmlImportBudget budget) {
        if (!HtmlImageDataUri.TryParse(imageElement.GetAttribute("src"), out HtmlImageDataUri dataUri)) return;
        string shapeLimit = string.Empty;
        string imageLimit = string.Empty;
        if (!budget.TryReserveShape(out shapeLimit) || !budget.TryReserveImage(dataUri, out imageLimit)) {
            Add(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
                "An embedded HTML image was omitted because the shared import limit was reached.",
                HtmlDiagnosticSeverity.Warning, HtmlConversionLossKind.Omission,
                shapeLimit.Length > 0 ? shapeLimit : imageLimit);
            return;
        }
        if (!dataUri.TryDecodeBytes(out byte[] bytes)) {
            Add(result, HtmlConversionDiagnosticCodes.ResourceDecodeFailed,
                "An embedded HTML image could not be decoded.", HtmlDiagnosticSeverity.Warning, HtmlConversionLossKind.Omission);
            return;
        }
        target.Add(new OneNoteImage {
            AltText = imageElement.GetAttribute("alt"),
            MediaType = dataUri.MediaType,
            FileName = "image" + dataUri.FileExtension,
            Payload = OneNoteBinaryPayload.FromBytes(bytes)
        });
        result.Elements++;
        result.Images++;
    }

    private static IEnumerable<IElement> DirectRows(IElement table) {
        foreach (IElement child in table.Children) {
            if (string.Equals(child.LocalName, "tr", StringComparison.OrdinalIgnoreCase)) yield return child;
            else if (string.Equals(child.LocalName, "thead", StringComparison.OrdinalIgnoreCase)
                || string.Equals(child.LocalName, "tbody", StringComparison.OrdinalIgnoreCase)
                || string.Equals(child.LocalName, "tfoot", StringComparison.OrdinalIgnoreCase)) {
                foreach (IElement row in child.Children.Where(candidate => string.Equals(candidate.LocalName, "tr", StringComparison.OrdinalIgnoreCase))) yield return row;
            }
        }
    }

    private static bool IsTableCell(IElement element) =>
        string.Equals(element.LocalName, "th", StringComparison.OrdinalIgnoreCase)
        || string.Equals(element.LocalName, "td", StringComparison.OrdinalIgnoreCase);

    private static void AppendRuns(INode node, OneNoteParagraph paragraph, InlineState state) {
        if (node is IText text) {
            string value = text.Data;
            if (value.Length == 0) return;
            var run = new OneNoteTextRun { Text = value, Hyperlink = state.Hyperlink };
            run.Style.Bold = state.Bold ? true : null;
            run.Style.Italic = state.Italic ? true : null;
            run.Style.Underline = state.Underline ? true : null;
            run.Style.Strikethrough = state.Strikethrough ? true : null;
            run.Style.Superscript = state.Superscript ? true : null;
            run.Style.Subscript = state.Subscript ? true : null;
            paragraph.Runs.Add(run);
            return;
        }
        if (!(node is IElement element)) return;
        string name = element.LocalName;
        if (string.Equals(name, "br", StringComparison.OrdinalIgnoreCase)) {
            paragraph.Runs.Add(new OneNoteTextRun { Text = "\n" });
            return;
        }
        InlineState nested = state.With(element);
        foreach (INode child in element.ChildNodes) AppendRuns(child, paragraph, nested);
    }

    private static void TrimRuns(OneNoteParagraph paragraph) {
        while (paragraph.Runs.Count > 0 && string.IsNullOrWhiteSpace(paragraph.Runs[0].Text)) paragraph.Runs.RemoveAt(0);
        while (paragraph.Runs.Count > 0 && string.IsNullOrWhiteSpace(paragraph.Runs[paragraph.Runs.Count - 1].Text)) paragraph.Runs.RemoveAt(paragraph.Runs.Count - 1);
        if (paragraph.Runs.Count > 0) paragraph.Runs[0].Text = paragraph.Runs[0].Text.TrimStart();
        if (paragraph.Runs.Count > 0) paragraph.Runs[paragraph.Runs.Count - 1].Text = paragraph.Runs[paragraph.Runs.Count - 1].Text.TrimEnd();
    }

    private static void Add(
        HtmlToOneNoteSectionResult result,
        string code,
        string message,
        HtmlDiagnosticSeverity severity,
        HtmlConversionLossKind lossKind,
        string? detail = null) =>
        result.AddImportDiagnostic(new HtmlDiagnostic(ComponentName, code, message, severity, detail: detail, lossKind: lossKind));

    private static T Require<T>(HtmlConversionResult<T> result) {
        if (result.Succeeded) return result.Value;
        throw new HtmlConversionException(result.Report.Diagnostics);
    }

    private static string CleanName(string? value, string fallback) {
        string normalized = string.Join(" ", (value ?? string.Empty).Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
        return normalized.Length == 0 ? fallback : normalized;
    }

    private readonly struct InlineState {
        internal bool Bold { get; }
        internal bool Italic { get; }
        internal bool Underline { get; }
        internal bool Strikethrough { get; }
        internal bool Superscript { get; }
        internal bool Subscript { get; }
        internal string? Hyperlink { get; }

        private InlineState(bool bold, bool italic, bool underline, bool strikethrough, bool superscript, bool subscript, string? hyperlink) {
            Bold = bold; Italic = italic; Underline = underline; Strikethrough = strikethrough;
            Superscript = superscript; Subscript = subscript; Hyperlink = hyperlink;
        }

        internal InlineState With(IElement element) {
            string name = element.LocalName;
            return new InlineState(
                Bold || name == "strong" || name == "b",
                Italic || name == "em" || name == "i",
                Underline || name == "u",
                Strikethrough || name == "s" || name == "strike" || name == "del",
                Superscript || name == "sup",
                Subscript || name == "sub",
                name == "a" ? element.GetAttribute("href") : Hyperlink);
        }
    }
}
