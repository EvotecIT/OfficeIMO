using System.Globalization;
using OfficeIMO.Web.Converter.Models;
using OfficeIMO.Web.Converter.Services;

namespace OfficeIMO.Web.Converter.Engine;

/// <summary>Runs the format conversion routes and describes the result in visitor terms.</summary>
internal static class ConvertTool {
    internal static ToolResultDocument Run(ToolSession session, string routeId, string action, ToolOptions options) {
        ConversionRoute route = ConversionRouteCatalog.All.FirstOrDefault(candidate => string.Equals(candidate.Id, routeId, StringComparison.Ordinal))
            ?? throw new NotSupportedException($"The conversion '{routeId}' is not available in the browser.");
        return action switch {
            "convert" => Convert(session, route, options),
            "support" => SupportBundle(session, options),
            _ => throw new NotSupportedException($"The conversion tool has no '{action}' action.")
        };
    }

    /// <summary>
    /// PDF imports draw with the full pack. Conversions to PDF need the Japanese, Arabic and symbol fallbacks only
    /// when the document's text or fonts call for them; other conversions draw nothing.
    /// </summary>
    private static bool NeedsFallbackFonts(ToolSession session, ConversionRoute route, ToolOptions options) {
        if (route.Id.StartsWith("pdf-", StringComparison.Ordinal)) return true;
        if (!route.Id.EndsWith("-pdf", StringComparison.Ordinal)) return false;
        if (route.InputKind == ConversionInputKind.Text) return BrowserFallbackFonts.Needed(RequireText(options));
        SelectedDocument file = session.Input();
        return BrowserFallbackFonts.Needed(file.Bytes, file.Extension);
    }

    private static ToolResultDocument Convert(ToolSession session, ConversionRoute route, ToolOptions options) {
        BrowserPdfProfile profile = BrowserPdfProfileCatalog.Find(options.Text("profile", BrowserPdfProfileCatalog.Faithful.Id));
        bool overlay = options.Flag("overlay");
        if (NeedsFallbackFonts(session, route, options) && !BrowserFallbackFonts.IsAvailable) {
            throw new EngineNeedsAssemblyException(BrowserFallbackFonts.AssetName);
        }
        ConversionResult result = route.InputKind == ConversionInputKind.File
            ? session.Conversions.ConvertUpload(
                route,
                session.Input(),
                limitExcelRows: options.Text("rows") == "preview",
                profile,
                overlay,
                options.Text("slides", "editable"))
            : session.Conversions.ConvertText(route, RequireText(options), profile, overlay);
        session.LastConversion = result;

        ToolArtifact[] artifacts = session.ReplaceArtifacts(
            (new BrowserConversionArtifact(result.Bytes, result.FileName, result.ContentType), "primary"),
            (result.CompanionReport, "report"),
            (result.DebugOverlay, "overlay"));

        ConversionWarningView[] review = result.StructuredWarnings
            .Where(static warning => !string.Equals(warning.Severity, "Information", StringComparison.OrdinalIgnoreCase))
            .ToArray();
        ConversionWarningView[] information = result.StructuredWarnings.Except(review).ToArray();

        string output = OutputName(route.Target);
        bool pdfOutput = string.Equals(route.Target, "PDF", StringComparison.OrdinalIgnoreCase);
        // Page counts describe the output only when it is paginated; for PDF imports they count the source.
        int? outputPages = pdfOutput ? result.PageCount : null;
        ToolVerdict verdict;
        if (review.Length == 0) {
            verdict = new ToolVerdict(ToolTone.Good, Headline(output, outputPages), FidelityDetail(result.FidelityStatus));
        } else {
            int[] pages = review.Where(static warning => warning.PageNumber.HasValue).Select(static warning => warning.PageNumber!.Value).ToArray();
            string where = pages.Length == review.Length && pages.Length > 0 ? $"All on {ToolFormat.Pages(pages)}. " : pages.Length > 0 ? $"Mostly on {ToolFormat.Pages(pages)}. " : string.Empty;
            string impact = review.Any(static warning => warning.CanChangePagination)
                ? "Some can change line breaks or page breaks."
                : "Content was substituted or approximated; the file is still usable.";
            verdict = new ToolVerdict(
                ToolTone.Warn,
                Headline(output, outputPages) + " · " + ToolFormat.Count(review.Length, "thing to check", "things to check"),
                where + impact);
        }

        var facts = new List<ToolFact>();
        if (result.PageCount is int pageCount) facts.Add(new ToolFact(pdfOutput ? "Pages" : "Source pages", pageCount.ToString(CultureInfo.InvariantCulture)));
        facts.Add(new ToolFact("Size", ToolFormat.Bytes(result.Bytes.LongLength)));
        if (result.ConversionMilliseconds > 0) facts.Add(new ToolFact("Time", ToolFormat.Duration(result.ConversionMilliseconds)));
        if (pdfOutput) facts.Add(new ToolFact("PDF type", ProfileName(profile)));

        var items = new List<ToolItem>();
        int index = 0;
        foreach (ConversionWarningView warning in review.OrderBy(static warning => warning.PageNumber ?? int.MaxValue)) {
            items.Add(WarningItem(warning, index++, ToolState.Warning));
        }
        foreach (ConversionWarningView warning in information.OrderBy(static warning => warning.PageNumber ?? int.MaxValue)) {
            items.Add(WarningItem(warning, index++, ToolState.Info));
        }

        return new ToolResultDocument(true, verdict, facts.ToArray(), items.ToArray(), artifacts, Preview(result), result.ConversionMilliseconds, result.PeakRetainedMemoryBytes);
    }

    private static ToolResultDocument SupportBundle(ToolSession session, ToolOptions options) {
        ConversionResult result = session.LastConversion ?? throw new InvalidOperationException("Convert a file before preparing a bug report.");
        BrowserConversionArtifact bundle = session.Conversions.CreateSupportBundle(result, options.Flag("includeContent"));
        ToolArtifact[] artifacts = session.AppendArtifacts((bundle, "support"));
        return new ToolResultDocument(
            true,
            new ToolVerdict(ToolTone.Info, "Bug report files ready",
                options.Flag("includeContent")
                    ? "The ZIP includes your source file and the PDF because you chose to add them."
                    : "The ZIP holds fingerprints and diagnostics only. Your document content is not included."),
            [], [], artifacts, null, 0);
    }

    private static string RequireText(ToolOptions options) {
        string text = options.Text("text");
        if (string.IsNullOrWhiteSpace(text)) throw new InvalidOperationException("Paste or type some content first.");
        return text;
    }

    private static ToolItem WarningItem(ConversionWarningView warning, int index, string state) {
        string title = ToolFormat.Words(string.Equals(warning.Construct, warning.Code, StringComparison.OrdinalIgnoreCase) ? warning.Code : warning.Construct);
        // Engine codes are namespaced ("PdfVectorGraphics…"); the tool already says which format it is.
        foreach (string prefix in new[] { "Pdf ", "Word ", "Excel ", "Power point ", "Html ", "Markdown " }) {
            if (title.StartsWith(prefix, StringComparison.Ordinal) && title.Length > prefix.Length + 3) {
                title = char.ToUpperInvariant(title[prefix.Length]) + title[(prefix.Length + 1)..];
                break;
            }
        }
        string detail = warning.Message + (warning.CanChangePagination ? " This can change line breaks or page breaks." : string.Empty);
        return new ToolItem(
            "w" + index.ToString(CultureInfo.InvariantCulture),
            title,
            detail,
            state,
            warning.PageNumber.HasValue ? "Page " + warning.PageNumber.Value.ToString(CultureInfo.InvariantCulture) : "Whole document",
            warning.PageNumber);
    }

    private static ToolPreview? Preview(ConversionResult result) {
        if (string.Equals(result.ContentType, "application/pdf", StringComparison.OrdinalIgnoreCase)) {
            return new ToolPreview("pdf", 0, result.PageCount);
        }
        if (result.ContentType.StartsWith("image/", StringComparison.OrdinalIgnoreCase)) return new ToolPreview("image", 0);
        if (!string.IsNullOrWhiteSpace(result.HtmlPreview)) return new ToolPreview("html", 0, Html: result.HtmlPreview, Text: result.Text);
        // Office outputs carry only a status line as text; show it as a preview only for text formats.
        bool textOutput = result.ContentType.StartsWith("text/", StringComparison.OrdinalIgnoreCase);
        if (textOutput && !string.IsNullOrEmpty(result.Text)) return new ToolPreview("text", 0, Text: result.Text);
        return null;
    }

    private static string OutputName(string target) => target.ToUpperInvariant() switch {
        "PDF" => "PDF",
        "DOCX" => "Word document",
        "XLSX" => "Excel workbook",
        "PPTX" => "PowerPoint file",
        "HTML" => "HTML",
        "MD" or "MARKDOWN" => "Markdown",
        "PNG" => "Images",
        _ => target
    };

    private static string Headline(string output, int? pages) =>
        output + " ready" + (pages is int count ? " · " + ToolFormat.Count(count, "page") : string.Empty);

    private static string FidelityDetail(string? fidelity) => fidelity switch {
        null or "" or "Complete" or "Faithful" => "Everything converted without changes worth checking.",
        "FaithfulWithSubstitutions" => "Everything converted. A font you don't have here was swapped for a close match.",
        "Degraded" => "Some content was approximated or left out. The list below says what.",
        "Reconstructed" => "Rebuilt as editable content. Layout can differ from the original, so look it over.",
        "Visual" => "Each page is kept as a picture. It looks right, but text is not editable.",
        "Hybrid" => "Pages are kept as pictures with editable tables on top.",
        "Partial" => "Only part of the content was converted, as this mode intends.",
        _ => ToolFormat.Words(fidelity) + " conversion."
    };

    private static string ProfileName(BrowserPdfProfile profile) => profile.Kind switch {
        BrowserPdfProfileKind.Faithful => "Standard",
        BrowserPdfProfileKind.Portable => "Portable",
        BrowserPdfProfileKind.Archival => "Archive (PDF/A-2b)",
        BrowserPdfProfileKind.Accessible => "Accessible (tagged)",
        BrowserPdfProfileKind.Diagnostic => "Layout check",
        _ => profile.Label
    };
}
