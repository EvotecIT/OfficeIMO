using System;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Text;

namespace PSWriteOffice.Services.Pdf;

internal static class PdfRichTextRunBuilder
{
    internal static void ApplyText(
        PdfParagraph builder,
        string[] text,
        bool bold,
        bool italic,
        bool underline,
        bool strike,
        PdfTextBaseline baseline,
        global::OfficeIMO.OfficeColor? color,
        global::OfficeIMO.OfficeColor? backgroundColor,
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

    internal static void ApplyRuns(PdfParagraph builder, object[] runs)
    {
        foreach (var run in OfficeTextRunParser.ParseMany(runs))
        {
            ApplyRun(builder, run);
        }
    }

    internal static void ApplyRuns(PdfParagraph builder, object? runs)
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

    private static void ApplyRun(PdfParagraph builder, OfficeTextRunSpec run)
    {
        if (run.IsLineBreak)
        {
            builder.AddLineBreak();
            return;
        }

        if (run.IsTab)
        {
            builder.AddTab(
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
            fontSize: ((double?)(run.FontSize) is double physicalLength1 ? global::OfficeIMO.OfficeLength.Points(physicalLength1) : (global::OfficeIMO.OfficeLength?)null),
            baseline: baseline,
            standardFont: GetNullableEnum<PdfStandardFont>(run.FontName),
            tabLeader: GetEnum(run.TabLeader, PdfTabLeaderStyle.None),
            tabAlignment: GetEnum(run.TabAlignment, PdfTabAlignment.Left),
            backgroundColor: ParseOptionalBackgroundColor(run.BackgroundColor));
    }

    private static void AddText(PdfParagraph builder, string text, string? linkUri, string? linkDestinationName, string? linkContents, global::OfficeIMO.OfficeColor? color, bool underline)
    {
        if (!string.IsNullOrWhiteSpace(linkUri))
        {
            builder.AddLink(text, linkUri!, color, underline, linkContents);
            return;
        }

        if (!string.IsNullOrWhiteSpace(linkDestinationName))
        {
            builder.AddBookmarkLink(text, linkDestinationName!, color, underline, linkContents);
            return;
        }

        builder.AddText(text);
    }

    private static void ApplyStyle(
        PdfParagraph builder,
        bool bold,
        bool italic,
        bool underline,
        bool strike,
        PdfTextBaseline baseline,
        global::OfficeIMO.OfficeColor? color,
        global::OfficeIMO.OfficeColor? backgroundColor,
        double? fontSize,
        PdfStandardFont? font)
    {
        builder.RunDefaults.Bold = bold;
        builder.RunDefaults.Italic = italic;
        builder.RunDefaults.UnderlineStyle = (underline) ? global::OfficeIMO.Drawing.OfficeTextDecorationStyle.Single : global::OfficeIMO.Drawing.OfficeTextDecorationStyle.None;
        builder.RunDefaults.StrikeStyle = (strike) ? global::OfficeIMO.Drawing.OfficeTextDecorationStyle.Single : global::OfficeIMO.Drawing.OfficeTextDecorationStyle.None;
        builder.RunDefaults.Baseline = baseline;

        if (color.HasValue)
        {
            builder.RunDefaults.Color = color.Value;
        }
        else
        {
            builder.RunDefaults.Color = null;
        }

        if (backgroundColor.HasValue)
        {
            builder.RunDefaults.BackgroundColor = backgroundColor.Value;
        }
        else
        {
            builder.RunDefaults.BackgroundColor = null;
        }

        if (fontSize.HasValue)
        {
            builder.RunDefaults.FontSize = global::OfficeIMO.OfficeLength.Points(fontSize.Value);
        }
        else
        {
            builder.RunDefaults.FontSize = null;
        }

        if (font.HasValue)
        {
            builder.RunDefaults.StandardFont = font.Value; builder.RunDefaults.FontFamily = null;
        }
        else
        {
            builder.RunDefaults.StandardFont = null; builder.RunDefaults.FontFamily = null;
        }
    }

    private static void ResetStyle(PdfParagraph builder)
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

    private static global::OfficeIMO.OfficeColor? ParseOptionalBackgroundColor(string? color)
        => OfficeColorUtilities.IsNone(color) ? null : PdfCommandUtilities.ParseColor(color);
}
