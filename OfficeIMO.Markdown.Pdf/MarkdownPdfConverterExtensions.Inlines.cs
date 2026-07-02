using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;
using PdfTextRun = OfficeIMO.Pdf.TextRun;

namespace OfficeIMO.Markdown.Pdf;

/// <summary>
/// First-party Markdown to PDF conversion helpers.
/// </summary>
public static partial class MarkdownPdfConverterExtensions {
    private static InlineStyle CreateInlineStyle(MarkdownPdfVisualTheme visualTheme) =>
        InlineStyle.Default.With(linkColor: visualTheme.LinkColorSnapshot, underlineLinks: visualTheme.UnderlineLinksSnapshot);

    private static void AppendInlines(PdfCore.PdfParagraphBuilder builder, InlineSequence sequence, InlineStyle style) {
        foreach (IMarkdownInline inline in sequence.Nodes) {
            AppendInline(builder, inline, style);
        }
    }

    private static void AppendInline(PdfCore.PdfParagraphBuilder builder, IMarkdownInline inline, InlineStyle style) {
        switch (inline) {
            case OfficeIMO.Markdown.TextRun text:
                ApplyStyle(builder, style).Text(text.Text);
                break;
            case BoldInline bold:
                ApplyStyle(builder, style.With(bold: true)).Text(bold.Text);
                break;
            case ItalicInline italic:
                ApplyStyle(builder, style.With(italic: true)).Text(italic.Text);
                break;
            case BoldItalicInline boldItalic:
                ApplyStyle(builder, style.With(bold: true, italic: true)).Text(boldItalic.Text);
                break;
            case UnderlineInline underline:
                ApplyStyle(builder, style.With(underline: true)).Text(underline.Text);
                break;
            case InsertedInline inserted:
                ApplyStyle(builder, style.With(underline: true)).Text(inserted.Text);
                break;
            case SuperscriptInline superscript:
                ApplyStyle(builder, style.With(baseline: PdfCore.PdfTextBaseline.Superscript)).Text(superscript.Text);
                break;
            case SubscriptInline subscript:
                ApplyStyle(builder, style.With(baseline: PdfCore.PdfTextBaseline.Subscript)).Text(subscript.Text);
                break;
            case StrikethroughInline strike:
                ApplyStyle(builder, style.With(strike: true)).Text(strike.Text);
                break;
            case HighlightInline highlight:
                ApplyStyle(builder, style.With(background: PdfCore.PdfColor.FromRgb(254, 243, 199))).Text(highlight.Text);
                break;
            case CodeSpanInline code:
                ApplyStyle(builder, style.With(background: PdfCore.PdfColor.FromRgb(241, 245, 249), color: PdfCore.PdfColor.FromRgb(30, 41, 59))).Text(code.Text);
                break;
            case LinkInline link:
                AppendLinkInline(builder, link, style);
                break;
            case BoldSequenceInline boldSequence:
                AppendInlines(builder, boldSequence.Inlines, style.With(bold: true));
                break;
            case ItalicSequenceInline italicSequence:
                AppendInlines(builder, italicSequence.Inlines, style.With(italic: true));
                break;
            case BoldItalicSequenceInline boldItalicSequence:
                AppendInlines(builder, boldItalicSequence.Inlines, style.With(bold: true, italic: true));
                break;
            case StrikethroughSequenceInline strikethroughSequence:
                AppendInlines(builder, strikethroughSequence.Inlines, style.With(strike: true));
                break;
            case HighlightSequenceInline highlightSequence:
                AppendInlines(builder, highlightSequence.Inlines, style.With(background: PdfCore.PdfColor.FromRgb(254, 243, 199)));
                break;
            case InsertedSequenceInline insertedSequence:
                AppendInlines(builder, insertedSequence.Inlines, style.With(underline: true));
                break;
            case SuperscriptSequenceInline superscriptSequence:
                AppendInlines(builder, superscriptSequence.Inlines, style.With(baseline: PdfCore.PdfTextBaseline.Superscript));
                break;
            case SubscriptSequenceInline subscriptSequence:
                AppendInlines(builder, subscriptSequence.Inlines, style.With(baseline: PdfCore.PdfTextBaseline.Subscript));
                break;
            case ImageInline image:
                ApplyStyle(builder, style.With(italic: true)).Text("[Image: " + (image.PlainAlt.Length == 0 ? image.Src : image.PlainAlt) + "]");
                break;
            case ImageLinkInline imageLink:
                ApplyStyle(builder, style.With(italic: true)).Text("[Image: " + (imageLink.PlainAlt.Length == 0 ? imageLink.ImageUrl : imageLink.PlainAlt) + "]");
                break;
            case HardBreakInline:
                builder.LineBreak();
                break;
            case HtmlTagSequenceInline htmlTag:
                AppendHtmlTagInline(builder, htmlTag, style);
                break;
            case IInlineContainerMarkdownInline container when container.NestedInlines != null:
                AppendInlines(builder, container.NestedInlines!, style);
                break;
            case IPlainTextMarkdownInline plain:
                var textBuilder = new StringBuilder();
                plain.AppendPlainText(textBuilder);
                ApplyStyle(builder, style).Text(textBuilder.ToString());
                break;
        }
    }

    private static void AppendLinkInline(PdfCore.PdfParagraphBuilder builder, LinkInline link, InlineStyle style) {
        string label = string.IsNullOrEmpty(link.Text) ? link.Url : link.Text;
        bool underline = style.UnderlineLinks ?? true;
        InlineStyle linkStyle = style.With(underline: underline, color: style.LinkColor ?? PdfCore.PdfColor.FromRgb(37, 99, 235));
        if (TryGetBookmarkTarget(link.Url, out string? bookmark)) {
            ApplyStyle(builder, linkStyle).LinkToBookmark(label, bookmark!, color: linkStyle.Color, underline: underline, contents: link.Title ?? label);
            return;
        }

        string? absolute = NormalizeAbsoluteLink(link.Url);
        if (absolute != null) {
            ApplyStyle(builder, linkStyle).Link(label, absolute, color: linkStyle.Color, underline: underline, contents: link.Title ?? label);
            return;
        }

        ApplyStyle(builder, style).Text(label);
    }

    private static void AppendHtmlTagInline(PdfCore.PdfParagraphBuilder builder, HtmlTagSequenceInline htmlTag, InlineStyle style) {
        InlineStyle tagStyle = htmlTag.TagName switch {
            "strong" or "b" => style.With(bold: true),
            "em" or "i" => style.With(italic: true),
            "u" or "ins" => style.With(underline: true),
            "del" or "s" => style.With(strike: true),
            "sup" => style.With(baseline: PdfCore.PdfTextBaseline.Superscript),
            "sub" => style.With(baseline: PdfCore.PdfTextBaseline.Subscript),
            "mark" => style.With(background: PdfCore.PdfColor.FromRgb(254, 243, 199)),
            _ => style
        };
        AppendInlines(builder, htmlTag.Inlines, tagStyle);
    }

    private static IReadOnlyList<PdfTextRun> ToTextRuns(InlineSequence sequence, InlineStyle style) {
        var runs = new List<PdfTextRun>();
        foreach (IMarkdownInline inline in sequence.Nodes) {
            AddTextRuns(runs, inline, style);
        }

        return runs;
    }

    private static void AddTextRuns(List<PdfTextRun> runs, IMarkdownInline inline, InlineStyle style) {
        switch (inline) {
            case OfficeIMO.Markdown.TextRun text:
                runs.Add(CreateRun(text.Text, style));
                break;
            case BoldInline bold:
                runs.Add(CreateRun(bold.Text, style.With(bold: true)));
                break;
            case ItalicInline italic:
                runs.Add(CreateRun(italic.Text, style.With(italic: true)));
                break;
            case BoldItalicInline boldItalic:
                runs.Add(CreateRun(boldItalic.Text, style.With(bold: true, italic: true)));
                break;
            case UnderlineInline underline:
                runs.Add(CreateRun(underline.Text, style.With(underline: true)));
                break;
            case InsertedInline inserted:
                runs.Add(CreateRun(inserted.Text, style.With(underline: true)));
                break;
            case SuperscriptInline superscript:
                runs.Add(CreateRun(superscript.Text, style.With(baseline: PdfCore.PdfTextBaseline.Superscript)));
                break;
            case SubscriptInline subscript:
                runs.Add(CreateRun(subscript.Text, style.With(baseline: PdfCore.PdfTextBaseline.Subscript)));
                break;
            case StrikethroughInline strike:
                runs.Add(CreateRun(strike.Text, style.With(strike: true)));
                break;
            case HighlightInline highlight:
                runs.Add(CreateRun(highlight.Text, style.With(background: PdfCore.PdfColor.FromRgb(254, 243, 199))));
                break;
            case CodeSpanInline code:
                runs.Add(CreateRun(code.Text, style.With(background: PdfCore.PdfColor.FromRgb(241, 245, 249), color: PdfCore.PdfColor.FromRgb(30, 41, 59))));
                break;
            case LinkInline link:
                AddLinkRun(runs, link, style);
                break;
            case BoldSequenceInline boldSequence:
                foreach (IMarkdownInline nested in boldSequence.Inlines.Nodes) {
                    AddTextRuns(runs, nested, style.With(bold: true));
                }
                break;
            case ItalicSequenceInline italicSequence:
                foreach (IMarkdownInline nested in italicSequence.Inlines.Nodes) {
                    AddTextRuns(runs, nested, style.With(italic: true));
                }
                break;
            case BoldItalicSequenceInline boldItalicSequence:
                foreach (IMarkdownInline nested in boldItalicSequence.Inlines.Nodes) {
                    AddTextRuns(runs, nested, style.With(bold: true, italic: true));
                }
                break;
            case StrikethroughSequenceInline strikethroughSequence:
                foreach (IMarkdownInline nested in strikethroughSequence.Inlines.Nodes) {
                    AddTextRuns(runs, nested, style.With(strike: true));
                }
                break;
            case HighlightSequenceInline highlightSequence:
                foreach (IMarkdownInline nested in highlightSequence.Inlines.Nodes) {
                    AddTextRuns(runs, nested, style.With(background: PdfCore.PdfColor.FromRgb(254, 243, 199)));
                }
                break;
            case InsertedSequenceInline insertedSequence:
                foreach (IMarkdownInline nested in insertedSequence.Inlines.Nodes) {
                    AddTextRuns(runs, nested, style.With(underline: true));
                }
                break;
            case SuperscriptSequenceInline superscriptSequence:
                foreach (IMarkdownInline nested in superscriptSequence.Inlines.Nodes) {
                    AddTextRuns(runs, nested, style.With(baseline: PdfCore.PdfTextBaseline.Superscript));
                }
                break;
            case SubscriptSequenceInline subscriptSequence:
                foreach (IMarkdownInline nested in subscriptSequence.Inlines.Nodes) {
                    AddTextRuns(runs, nested, style.With(baseline: PdfCore.PdfTextBaseline.Subscript));
                }
                break;
            case HardBreakInline:
                runs.Add(PdfTextRun.LineBreak());
                break;
            case HtmlTagSequenceInline htmlTag:
                InlineStyle tagStyle = htmlTag.TagName switch {
                    "strong" or "b" => style.With(bold: true),
                    "em" or "i" => style.With(italic: true),
                    "u" or "ins" => style.With(underline: true),
                    "del" or "s" => style.With(strike: true),
                    "sup" => style.With(baseline: PdfCore.PdfTextBaseline.Superscript),
                    "sub" => style.With(baseline: PdfCore.PdfTextBaseline.Subscript),
                    "mark" => style.With(background: PdfCore.PdfColor.FromRgb(254, 243, 199)),
                    _ => style
                };
                foreach (IMarkdownInline nested in htmlTag.Inlines.Nodes) {
                    AddTextRuns(runs, nested, tagStyle);
                }
                break;
            case IInlineContainerMarkdownInline container when container.NestedInlines != null:
                foreach (IMarkdownInline nested in container.NestedInlines!.Nodes) {
                    AddTextRuns(runs, nested, style);
                }
                break;
            case IPlainTextMarkdownInline plain:
                var textBuilder = new StringBuilder();
                plain.AppendPlainText(textBuilder);
                runs.Add(CreateRun(textBuilder.ToString(), style));
                break;
        }
    }

    private static void AddLinkRun(List<PdfTextRun> runs, LinkInline link, InlineStyle style) {
        string label = string.IsNullOrEmpty(link.Text) ? link.Url : link.Text;
        PdfCore.PdfColor linkColor = style.LinkColor ?? PdfCore.PdfColor.FromRgb(37, 99, 235);
        bool underline = style.UnderlineLinks ?? true;
        if (TryGetBookmarkTarget(link.Url, out string? bookmark)) {
            runs.Add(CreateLinkRun(label, style, linkColor, underline, link.Title ?? label, uri: null, bookmark: bookmark));
            return;
        }

        string? absolute = NormalizeAbsoluteLink(link.Url);
        if (absolute != null) {
            runs.Add(CreateLinkRun(label, style, linkColor, underline, link.Title ?? label, uri: absolute, bookmark: null));
            return;
        }

        runs.Add(CreateRun(label, style));
    }

    private static PdfTextRun CreateLinkRun(string text, InlineStyle style, PdfCore.PdfColor color, bool underline, string contents, string? uri, string? bookmark) =>
        new PdfTextRun(
            text,
            bold: style.Bold,
            underline: underline,
            color: color,
            italic: style.Italic,
            strike: style.Strike,
            fontSize: style.FontSize,
            font: style.Font,
            linkUri: uri,
            linkContents: contents,
            baseline: style.Baseline,
            linkDestinationName: bookmark,
            backgroundColor: style.Background);

    private static PdfTextRun CreateRun(string text, InlineStyle style) {
        return new PdfTextRun(
            text,
            bold: style.Bold,
            underline: style.Underline,
            color: style.Color,
            italic: style.Italic,
            strike: style.Strike,
            fontSize: style.FontSize,
            font: style.Font,
            baseline: style.Baseline,
            backgroundColor: style.Background);
    }

    private static PdfCore.PdfParagraphBuilder ApplyStyle(PdfCore.PdfParagraphBuilder builder, InlineStyle style) {
        builder.Bold(style.Bold)
            .Italic(style.Italic)
            .Underline(style.Underline)
            .Strike(style.Strike)
            .Baseline(style.Baseline);

        if (style.FontSize.HasValue) {
            builder.FontSize(style.FontSize.Value);
        } else {
            builder.ResetFontSize();
        }

        if (style.Font.HasValue) {
            builder.Font(style.Font.Value);
        } else {
            builder.ResetFont();
        }

        if (style.Color.HasValue) {
            builder.Color(style.Color.Value);
        } else {
            builder.ResetColor();
        }

        if (style.Background.HasValue) {
            builder.BackgroundColor(style.Background.Value);
        } else {
            builder.ResetBackgroundColor();
        }

        return builder;
    }

    private static void AppendTextWithLineBreaks(PdfCore.PdfParagraphBuilder builder, string text) {
        string[] lines = (text ?? string.Empty).Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
        for (int i = 0; i < lines.Length; i++) {
            if (i > 0) {
                builder.LineBreak();
            }

            builder.Text(lines[i]);
        }
    }

    private static bool IsEmpty(InlineSequence sequence) {
        if (sequence.Nodes.Count == 0) {
            return true;
        }

        var builder = new StringBuilder();
        foreach (IMarkdownInline inline in sequence.Nodes) {
            if (inline is IPlainTextMarkdownInline plain) {
                plain.AppendPlainText(builder);
            }
        }

        return string.IsNullOrWhiteSpace(builder.ToString());
    }


    private readonly struct InlineStyle {
        public InlineStyle(
            bool bold,
            bool italic,
            bool underline,
            bool strike,
            PdfCore.PdfTextBaseline baseline,
            PdfCore.PdfColor? color,
            PdfCore.PdfColor? background,
            double? fontSize,
            PdfCore.PdfStandardFont? font,
            PdfCore.PdfColor? linkColor,
            bool? underlineLinks) {
            Bold = bold;
            Italic = italic;
            Underline = underline;
            Strike = strike;
            Baseline = baseline;
            Color = color;
            Background = background;
            FontSize = fontSize;
            Font = font;
            LinkColor = linkColor;
            UnderlineLinks = underlineLinks;
        }

        public static InlineStyle Default { get; } = new InlineStyle(false, false, false, false, PdfCore.PdfTextBaseline.Normal, null, null, null, null, null, null);

        public bool Bold { get; }

        public bool Italic { get; }

        public bool Underline { get; }

        public bool Strike { get; }

        public PdfCore.PdfTextBaseline Baseline { get; }

        public PdfCore.PdfColor? Color { get; }

        public PdfCore.PdfColor? Background { get; }

        public double? FontSize { get; }

        public PdfCore.PdfStandardFont? Font { get; }

        public PdfCore.PdfColor? LinkColor { get; }

        public bool? UnderlineLinks { get; }

        public InlineStyle With(
            bool? bold = null,
            bool? italic = null,
            bool? underline = null,
            bool? strike = null,
            PdfCore.PdfTextBaseline? baseline = null,
            PdfCore.PdfColor? color = null,
            PdfCore.PdfColor? background = null,
            double? fontSize = null,
            PdfCore.PdfStandardFont? font = null,
            PdfCore.PdfColor? linkColor = null,
            bool? underlineLinks = null) =>
            new InlineStyle(
                bold ?? Bold,
                italic ?? Italic,
                underline ?? Underline,
                strike ?? Strike,
                baseline ?? Baseline,
                color ?? Color,
                background ?? Background,
                fontSize ?? FontSize,
                font ?? Font,
                linkColor ?? LinkColor,
                underlineLinks ?? UnderlineLinks);
    }
}
