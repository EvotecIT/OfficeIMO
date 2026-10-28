using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

/// <summary>Configurable typography and quiet table styling for printed reports, independent of a product theme.</summary>
public sealed partial class WordPaperStyle {
    /// <summary>Gets or sets the body typeface.</summary>
    public string FontFamily { get; set; } = "Aptos";
    /// <summary>Gets or sets the display typeface.</summary>
    public string DisplayFontFamily { get; set; } = "Aptos Display";
    /// <summary>Gets or sets the primary text color, as RGB hex.</summary>
    public string Ink { get; set; } = "202428";
    /// <summary>Gets or sets secondary text color, as RGB hex.</summary>
    public string Muted { get; set; } = "687078";
    /// <summary>Gets or sets the eyebrow and label accent, as RGB hex.</summary>
    public string Accent { get; set; } = "405060";
    /// <summary>Gets or sets the quiet table header and metric surface, as RGB hex.</summary>
    public string Surface { get; set; } = "F5F6F7";
    /// <summary>Gets or sets horizontal table rule color, as RGB hex.</summary>
    public string Border { get; set; } = "D9DDE1";
    /// <summary>Gets or sets the hairline thickness in eighths of a point (4 is 0.5 pt).</summary>
    public uint RuleSize { get; set; } = 4;

    /// <summary>Gets or sets table body type size in points. Header cells use 9 points.</summary>
    public double TableFontSizePoints { get; set; } = 10;
    /// <summary>Gets or sets top and bottom table-cell padding in twentieths of a point.</summary>
    public short TableCellVerticalPadding { get; set; } = 90;
    /// <summary>Gets or sets left and right table-cell padding in twentieths of a point.</summary>
    public short TableCellHorizontalPadding { get; set; } = 100;
    /// <summary>Gets or sets spacing before and after table paragraphs, in points.</summary>
    public double TableParagraphSpacingPoints { get; set; } = 2;
    /// <summary>Gets or sets exact table line spacing in points; null preserves inherited line spacing.</summary>
    public double? TableLineSpacingPoints { get; set; }

    /// <summary>Applies body typography and dark headings to existing document paragraphs without restyling tables.</summary>
    public void ApplyTypography(WordDocument document) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        foreach (WordParagraph paragraph in document.Paragraphs) {
            paragraph.FontFamily = FontFamily;
            int level = paragraph.Style == WordParagraphStyles.Heading1 ? 1
                : paragraph.Style == WordParagraphStyles.Heading2 ? 2
                : paragraph.Style == WordParagraphStyles.Heading3 ? 3
                : paragraph.Style == WordParagraphStyles.Heading4 ? 4 : 0;
            if (level == 0) continue;
            paragraph.FontSize = level == 1 ? 18 : level == 2 ? 14 : level == 3 ? 11 : 10;
            paragraph.ColorHex = Ink;
            paragraph.Bold = true;
            paragraph.LineSpacingBeforePoints = level == 1 ? 16 : level == 2 ? 12 : 8;
            paragraph.LineSpacingAfterPoints = level == 1 ? 8 : level == 2 ? 6 : 4;
            (paragraph._paragraph.ParagraphProperties ??= new ParagraphProperties()).KeepNext = new KeepNext();
        }
    }

    /// <summary>Writes multiline text into an empty table cell with paper typography.</summary>
    public void SetCellText(WordTableCell cell, string text) {
        if (cell == null) throw new ArgumentNullException(nameof(cell));
        if (text == null) throw new ArgumentNullException(nameof(text));
        string[] lines = text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
        for (int index = 0; index < lines.Length; index++) {
            WordParagraph paragraph = index == 0 ? cell.Paragraphs[0] : cell.AddParagraph();
            paragraph.Text = lines[index];
            Format(paragraph, 10, Ink);
        }
    }

    /// <summary>Applies a semantic text tone and optional tint after the table's base styling.</summary>
    /// <param name="cell">The cell to format.</param>
    /// <param name="tone">Text color as RGB hex.</param>
    /// <param name="tint">Optional background color as RGB hex; null retains the current fill.</param>
    public void ApplyCellTone(WordTableCell cell, string tone, string? tint = null) {
        if (cell == null) throw new ArgumentNullException(nameof(cell));
        if (tint != null) cell.ShadingFillColorHex = tint;
        foreach (WordParagraph paragraph in cell.Paragraphs) paragraph.ColorHex = tone;
    }

    /// <summary>Formats a value card with a small label, emphasized value and a semantic top rule.</summary>
    /// <param name="cell">Destination cell.</param>
    /// <param name="label">Short label above the value.</param>
    /// <param name="value">Evidence or state to emphasize.</param>
    /// <param name="tone">Rule color as RGB hex.</param>
    /// <param name="tint">Background color as RGB hex.</param>
    public void ApplyValueCard(WordTableCell cell, string label, string value, string tone, string tint) {
        if (cell == null) throw new ArgumentNullException(nameof(cell));
        cell.ShadingFillColorHex = tint;
        cell.VerticalAlignment = WordTableVerticalAlignment.Center;
        cell.MarginTopWidth = 150;
        cell.MarginBottomWidth = 150;
        cell.MarginLeftWidth = 170;
        cell.MarginRightWidth = 170;
        cell.Borders.TopStyle = WordBorderStyle.Single;
        cell.Borders.TopColorHex = tone;
        cell.Borders.TopSize = 20;
        WordParagraph caption = cell.AddParagraph(label.ToUpperInvariant(), true);
        Format(caption, 8, Muted, true);
        caption.LineSpacingAfterPoints = 2;
        WordParagraph content = cell.AddParagraph(value);
        Format(content, 11, Ink, true);
        content.LineSpacingBeforePoints = 0;
        content.LineSpacingAfterPoints = 0;
    }

    /// <summary>Formats an empty, unpadded cell used as a gutter between cards.</summary>
    /// <param name="cell">Destination cell.</param>
    /// <param name="background">Page background as RGB hex.</param>
    public void ApplySpacerCell(WordTableCell cell, string background = "FFFFFF") {
        if (cell == null) throw new ArgumentNullException(nameof(cell));
        cell.ShadingFillColorHex = background;
        cell.MarginTopWidth = 0;
        cell.MarginBottomWidth = 0;
        cell.MarginLeftWidth = 0;
        cell.MarginRightWidth = 0;
        cell.AddParagraph(string.Empty, true);
    }

    /// <summary>Adds a cover's eyebrow, title and subtitle using the paper type scale.</summary>
    public void AddCover(WordDocument document, string eyebrow, string title, string subtitle) {
        WordParagraph label = document.AddParagraph(eyebrow.ToUpperInvariant());
        Format(label, 10, Accent, true);
        label.LineSpacingBeforePoints = 16;
        label.LineSpacingAfterPoints = 18;
        WordParagraph heading = document.AddParagraph(title);
        Format(heading, 28, Ink, true);
        heading.FontFamily = DisplayFontFamily;
        heading.LineSpacingAfterPoints = 12;
        WordParagraph description = document.AddParagraph(subtitle);
        Format(description, 13, Muted);
        description.LineSpacingAfterPoints = 24;
    }

    /// <summary>Adds a small muted source or coverage note separated from preceding content.</summary>
    public WordParagraph AddNote(WordDocument document, string text) {
        WordParagraph paragraph = document.AddParagraph(text);
        Format(paragraph, 9, Muted);
        paragraph.LineSpacingBeforePoints = 12;
        paragraph.LineSpacingAfterPoints = 12;
        return paragraph;
    }

    /// <summary>Adds a shaded callout with a 3 pt left rule. The caller supplies semantic theme colors.</summary>
    public WordParagraph AddCallout(WordDocument document, string title, string body, string tone, string tint) {
        WordParagraph paragraph = document.AddParagraph();
        paragraph.AddFormattedText(title + ": ", bold: true);
        paragraph.AddText(body);
        Format(paragraph, 10, Ink);
        paragraph.ShadingFillColorHex = tint;
        paragraph.Borders.LeftStyle = WordBorderStyle.Single;
        paragraph.Borders.LeftSize = 24;
        paragraph.Borders.LeftColorHex = tone;
        paragraph.LineSpacingBeforePoints = 10;
        paragraph.LineSpacingAfterPoints = 12;
        return paragraph;
    }

    /// <summary>Adds a row of metric cards, each with a large value, label and optional qualifier.</summary>
    public WordTable AddMetrics(WordDocument document, params WordPaperMetric[] metrics) {
        if (metrics == null || metrics.Length == 0) throw new ArgumentException("At least one metric is required.", nameof(metrics));
        WordTable table = document.AddTable(1, metrics.Length, WordTableStyle.TableNormal);
        ApplyTable(table, hasHeader: false);
        for (int index = 0; index < metrics.Length; index++) {
            WordPaperMetric metric = metrics[index];
            WordTableCell cell = table.Rows[0].Cells[index];
            cell.ShadingFillColorHex = Surface;
            WordParagraph value = cell.Paragraphs[0];
            value.Text = metric.Value;
            Format(value, 24, metric.Color ?? Ink, true);
            Format(cell.AddParagraph(metric.Label), 9, Muted);
            if (!string.IsNullOrWhiteSpace(metric.Qualifier)) Format(cell.AddParagraph(metric.Qualifier!), 9, Muted);
        }
        return table;
    }

    /// <summary>Adds title and scope running headers and product, scope and page-total footers.</summary>
    public void AddPageChrome(WordDocument document, string title, string product, string scope) {
        document.AddHeadersAndFooters();
        document.DifferentOddAndEvenPages = true;
        string header = string.IsNullOrWhiteSpace(scope) ? title : title + " · " + scope;
        string footer = product + (string.IsNullOrWhiteSpace(scope) ? "" : " · " + scope) + " · ";
        foreach (WordHeader part in new[] { document.HeaderDefaultOrCreate, document.HeaderEvenOrCreate }) {
            WordParagraph paragraph = part.AddParagraph(header);
            Format(paragraph, 8, Muted);
            paragraph.ParagraphAlignment = WordParagraphAlignment.Center;
        }
        foreach (WordFooter part in new[] { document.FooterDefaultOrCreate, document.FooterEvenOrCreate }) {
            WordParagraph paragraph = part.AddParagraph(footer);
            paragraph.AddPageNumber(includeTotalPages: true, separator: " / ");
            Format(paragraph, 8, Muted);
            paragraph.ParagraphAlignment = WordParagraphAlignment.Center;
        }
    }

    private void Format(WordParagraph paragraph, int size, string color, bool bold = false) {
        paragraph.FontFamily = FontFamily;
        paragraph.FontSize = size;
        paragraph.ColorHex = color;
        if (bold) paragraph.Bold = true;
    }

    private static bool IsNumeric(string? text) => WordNumericText.IsNumber(text);
}
