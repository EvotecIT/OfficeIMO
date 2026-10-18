using System;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Text;

namespace PSWriteOffice.Services.Pdf;

internal static class PdfRichTextRunBuilder
{
    internal static void ApplyText(
        PdfParagraphBuilder builder,
        string[] text,
        bool bold,
        bool italic,
        bool underline,
        bool strike,
        PdfTextBaseline baseline,
        PdfColor? color,
        PdfColor? backgroundColor,
        double? fontSize,
        PdfStandardFont? font,
        string? linkUri,
        string? linkDestinationName,
        string? linkContents)
    {
        if (text.Length == 0)
        {
            throw new PSArgumentException("Provide at least one text value.");
        }

        if (linkUri != null && linkDestinationName != null)
        {
            throw new PSArgumentException("A PDF text link can target either -LinkUri or -LinkDestinationName, not both.");
        }

        ApplyStyle(builder, bold, italic, underline, strike, baseline, color, backgroundColor, fontSize, font);
        foreach (var value in text)
        {
            AddText(builder, value, linkUri, linkDestinationName, linkContents, color, underline || linkUri != null || linkDestinationName != null);
        }
    }

    internal static void ApplyRuns(PdfParagraphBuilder builder, object[] runs)
    {
        foreach (var run in OfficeTextRunParser.ParseMany(runs))
        {
            ApplyRun(builder, run);
        }
    }

    internal static void ApplyRuns(PdfParagraphBuilder builder, object? runs)
    {
        foreach (var run in OfficeTextRunParser.ParseMany(runs))
        {
            ApplyRun(builder, run);
        }
    }

    internal static object[] ToRunArray(object? runs)
        => OfficeTextRunParser.ToRunArray(runs);

    internal static PdfTextRun[] ToTextRuns(object? runs)
        => OfficeTextRunParser.ParseMany(runs).Select(run => ToTextRun(run, TableLinkError)).ToArray();

    internal static PdfTextRun[] ToTextRuns(OfficeTextRunSpec[] runs)
        => runs.Select(run => ToTextRun(run, TableLinkError)).ToArray();

    internal static PdfTextRun[] ToCanvasTextRuns(object? runs)
        => OfficeTextRunParser.ParseMany(runs).Select(run => ToTextRun(run, CanvasLinkError)).ToArray();

    internal static PdfTextRun[] ToCanvasTextRuns(OfficeTextRunSpec[] runs)
        => runs.Select(run => ToTextRun(run, CanvasLinkError)).ToArray();

    private const string TableLinkError =
        "PDF table cell text runs do not support links yet. Use PdfText outside the table or remove LinkUri/LinkDestinationName from the table cell run.";

    private const string CanvasLinkError =
        "PDF canvas text runs do not support links. Use flowing PdfText for clickable text or remove LinkUri/LinkDestinationName from the canvas run.";

    private static void ApplyRun(PdfParagraphBuilder builder, OfficeTextRunSpec run)
    {
        if (run.IsLineBreak)
        {
            builder.LineBreak();
            return;
        }

        if (run.IsTab)
        {
            builder.Tab(
                GetEnum(run.TabLeader, PdfTabLeaderStyle.None),
                GetEnum(run.TabAlignment, PdfTabAlignment.Left));
            return;
        }

        var color = PdfCommandUtilities.ParseColor(run.Color);
        var backgroundColor = ParseOptionalBackgroundColor(run.BackgroundColor);
        var font = GetNullableEnum<PdfStandardFont>(run.FontName);
        var baseline = GetEnum(run.Baseline, PdfTextBaseline.Normal);

        if (run.LinkUri != null && run.LinkDestinationName != null)
        {
            throw new PSArgumentException("A PDF text run can target either LinkUri or LinkDestinationName, not both.");
        }

        ApplyStyle(builder, run.Bold, run.Italic, run.Underline, run.Strike, baseline, color, backgroundColor, run.FontSize, font);
        AddText(builder, run.Text, run.LinkUri, run.LinkDestinationName, run.LinkContents, color, run.Underline || run.LinkUri != null || run.LinkDestinationName != null);
    }

    private static PdfTextRun ToTextRun(OfficeTextRunSpec run, string linkError)
    {
        var text = run.IsLineBreak ? Environment.NewLine : run.IsTab ? "\t" : run.Text;
        var baseline = GetEnum(run.Baseline, PdfTextBaseline.Normal);
        if (!string.IsNullOrWhiteSpace(run.LinkUri) || !string.IsNullOrWhiteSpace(run.LinkDestinationName))
        {
            throw new PSArgumentException(linkError);
        }

        return new PdfTextRun(
            text,
            bold: run.Bold,
            underline: run.Underline,
            color: PdfCommandUtilities.ParseColor(run.Color),
            italic: run.Italic,
            strike: run.Strike,
            fontSize: run.FontSize,
            baseline: baseline,
            font: GetNullableEnum<PdfStandardFont>(run.FontName),
            tabLeader: GetEnum(run.TabLeader, PdfTabLeaderStyle.None),
            tabAlignment: GetEnum(run.TabAlignment, PdfTabAlignment.Left),
            backgroundColor: ParseOptionalBackgroundColor(run.BackgroundColor));
    }

    private static void AddText(PdfParagraphBuilder builder, string text, string? linkUri, string? linkDestinationName, string? linkContents, PdfColor? color, bool underline)
    {
        if (!string.IsNullOrWhiteSpace(linkUri))
        {
            builder.Link(text, linkUri!, color, underline, linkContents);
            return;
        }

        if (!string.IsNullOrWhiteSpace(linkDestinationName))
        {
            builder.LinkToBookmark(text, linkDestinationName!, color, underline, linkContents);
            return;
        }

        builder.Text(text);
    }

    private static void ApplyStyle(
        PdfParagraphBuilder builder,
        bool bold,
        bool italic,
        bool underline,
        bool strike,
        PdfTextBaseline baseline,
        PdfColor? color,
        PdfColor? backgroundColor,
        double? fontSize,
        PdfStandardFont? font)
    {
        builder.Bold(bold)
            .Italic(italic)
            .Underline(underline)
            .Strike(strike)
            .Baseline(baseline);

        if (color.HasValue)
        {
            builder.Color(color.Value);
        }
        else
        {
            builder.ResetColor();
        }

        if (backgroundColor.HasValue)
        {
            builder.BackgroundColor(backgroundColor.Value);
        }
        else
        {
            builder.ResetBackgroundColor();
        }

        if (fontSize.HasValue)
        {
            builder.FontSize(fontSize.Value);
        }
        else
        {
            builder.ResetFontSize();
        }

        if (font.HasValue)
        {
            builder.Font(font.Value);
        }
        else
        {
            builder.ResetFont();
        }
    }

    private static void ResetStyle(PdfParagraphBuilder builder)
    {
        ApplyStyle(builder, bold: false, italic: false, underline: false, strike: false, PdfTextBaseline.Normal, null, null, null, null);
    }

    private static TEnum GetEnum<TEnum>(string? value, TEnum defaultValue)
        where TEnum : struct
        => GetNullableEnum<TEnum>(value) ?? defaultValue;

    private static TEnum? GetNullableEnum<TEnum>(string? value)
        where TEnum : struct
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        return (TEnum)Enum.Parse(typeof(TEnum), value!, ignoreCase: true);
    }

    private static PdfColor? ParseOptionalBackgroundColor(string? color)
        => OfficeColorUtilities.IsNone(color) ? null : PdfCommandUtilities.ParseColor(color);
}
