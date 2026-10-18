using System;
using System.Management.Automation;
using System.Reflection;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Services.Text;
using A = DocumentFormat.OpenXml.Drawing;

namespace PSWriteOffice.Services.PowerPoint;

internal static class PowerPointTextRunService
{
    private static readonly PropertyInfo? OpenXmlRunProperty = typeof(PowerPointTextRun).GetProperty(
        "Run",
        BindingFlags.Instance | BindingFlags.NonPublic);

    internal static void ValidateRuns(object[] runs, bool allowHyperlinks)
        => ValidateRuns(OfficeTextRunParser.ParseMany(runs), allowHyperlinks);

    internal static void ApplyRuns(PowerPointTextBox textBox, object[] runs)
    {
        var parsedRuns = OfficeTextRunParser.ParseMany(runs);
        ValidateRuns(parsedRuns, allowHyperlinks: true);
        textBox.Text = string.Empty;
        var paragraph = textBox.Paragraphs.Count > 0 ? textBox.Paragraphs[0] : textBox.AddParagraph();
        ApplyRuns(paragraph, parsedRuns, allowHyperlinks: true);
    }

    internal static void ApplyRuns(PowerPointTableCell cell, object[] runs)
    {
        var parsedRuns = OfficeTextRunParser.ParseMany(runs);
        ValidateRuns(parsedRuns, allowHyperlinks: false);
        cell.Text = string.Empty;
        var first = true;
        foreach (var spec in parsedRuns)
        {
            var run = first && cell.Runs.Count > 0 && string.IsNullOrEmpty(cell.Runs[0].Text)
                ? cell.Runs[0]
                : cell.AddRun(ToPowerPointText(spec));
            if (first && string.IsNullOrEmpty(run.Text))
            {
                run.Text = ToPowerPointText(spec);
            }

            ApplyStyle(run, spec, allowHyperlinks: false);
            first = false;
        }
    }

    private static void ApplyRuns(PowerPointParagraph paragraph, OfficeTextRunSpec[] runs, bool allowHyperlinks)
    {
        var first = true;
        foreach (var spec in runs)
        {
            var run = first && paragraph.Runs.Count > 0 && string.IsNullOrEmpty(paragraph.Runs[0].Text)
                ? paragraph.Runs[0]
                : paragraph.AddRun(ToPowerPointText(spec));
            if (first && string.IsNullOrEmpty(run.Text))
            {
                run.Text = ToPowerPointText(spec);
            }

            ApplyStyle(run, spec, allowHyperlinks);
            first = false;
        }
    }

    private static string ToPowerPointText(OfficeTextRunSpec run)
        => run.IsLineBreak ? Environment.NewLine : run.IsTab ? "\t" : run.Text;

    private static void ApplyStyle(PowerPointTextRun run, OfficeTextRunSpec spec, bool allowHyperlinks)
    {
        run.Bold = spec.Bold;
        run.Italic = spec.Italic;
        run.Underline = spec.Underline;
        run.Strikethrough = spec.Strike;
        run.FontSize = spec.FontSize.HasValue
            ? (int)Math.Round(spec.FontSize.Value)
            : null;
        run.FontName = !string.IsNullOrWhiteSpace(spec.FontName)
            ? spec.FontName
            : null;
        run.Color = OfficeColorUtilities.ToRgbHex(spec.Color);
        run.HighlightColor = OfficeColorUtilities.IsNone(spec.BackgroundColor)
            ? null
            : OfficeColorUtilities.ToRgbHex(spec.BackgroundColor);
        ApplyBaseline(run, spec.Baseline);
        run.ClearHyperlink();

        if (!string.IsNullOrWhiteSpace(spec.LinkDestinationName))
        {
            throw new PSArgumentException("PowerPoint text runs support URI links, not named PDF/Word destinations.", nameof(spec.LinkDestinationName));
        }

        if (!string.IsNullOrWhiteSpace(spec.LinkUri))
        {
            if (!allowHyperlinks)
            {
                throw new PSArgumentException("PowerPoint table cell runs do not support hyperlinks yet.", nameof(spec.LinkUri));
            }

            run.SetHyperlink(spec.LinkUri!, spec.LinkContents);
        }
    }

    private static void ValidateRuns(OfficeTextRunSpec[] runs, bool allowHyperlinks)
    {
        foreach (var spec in runs)
        {
            ValidateRun(spec, allowHyperlinks);
        }
    }

    private static void ValidateRun(OfficeTextRunSpec spec, bool allowHyperlinks)
    {
        if (!string.IsNullOrWhiteSpace(spec.LinkDestinationName))
        {
            throw new PSArgumentException("PowerPoint text runs support URI links, not named PDF/Word destinations.", nameof(spec.LinkDestinationName));
        }

        if (!allowHyperlinks && !string.IsNullOrWhiteSpace(spec.LinkUri))
        {
            throw new PSArgumentException("PowerPoint table cell runs do not support hyperlinks yet.", nameof(spec.LinkUri));
        }
    }

    private static void ApplyBaseline(PowerPointTextRun run, string? baseline)
    {
        var value = ResolveBaseline(baseline);
        var openXmlRun = (A.Run?)OpenXmlRunProperty?.GetValue(run);
        if (openXmlRun == null)
        {
            if (value.HasValue)
            {
                throw new PSArgumentException("PowerPoint run baselines require access to the underlying OpenXML run.", nameof(baseline));
            }

            return;
        }

        var properties = openXmlRun.RunProperties ??= new A.RunProperties();
        properties.Baseline = value;
    }

    private static int? ResolveBaseline(string? baseline)
    {
        return OfficeTextRunParser.NormalizeKind(baseline) switch
        {
            "superscript" or "super" or "sup" => 30000,
            "subscript" or "sub" => -25000,
            _ => null
        };
    }
}
