using System.Globalization;
using OfficeIMO.Pdf;
using OfficeIMO.Web.Converter.Models;
using OfficeIMO.Web.Converter.Services;

namespace OfficeIMO.Web.Converter.Engine;

/// <summary>Runs the PDF workbench operations and the read-only checks the shell needs before them.</summary>
internal static class PdfTool {
    private sealed record RedactionReview(SelectedDocument File, string Text);
    internal static ToolResultDocument Run(ToolSession session, string toolId, string action, ToolOptions options) {
        PdfToolDefinition tool = PdfToolCatalog.All.FirstOrDefault(candidate => string.Equals(candidate.Id, toolId, StringComparison.Ordinal))
            ?? throw new NotSupportedException($"The PDF tool '{toolId}' is not available in the browser.");
        return action switch {
            "run" => Execute(session, tool, options),
            "find" => FindText(session, options),
            _ => throw new NotSupportedException($"The PDF tool has no '{action}' action.")
        };
    }

    internal static PdfProbeDocument Probe(ToolSession session, int slot) {
        SelectedDocument file = session.Input(slot);
        PdfDocumentPreflight preflight = PdfDocument.Preflight(file.Bytes, BrowserPdfPolicy.CreateReadOptions());
        bool encrypted = preflight.Probe.Security.HasEncryption;
        return new PdfProbeDocument(
            true,
            preflight.DocumentInfo?.PageCount ?? 0,
            encrypted,
            NeedsPassword: encrypted && !preflight.CanRead,
            preflight.CanManipulatePages,
            preflight.CanRead ? null : "This PDF needs a password before it can be read.");
    }

    private static ToolResultDocument Execute(ToolSession session, PdfToolDefinition tool, ToolOptions options) {
        if (tool.Kind == PdfToolKind.Redact && (session.LastReview is not RedactionReview review ||
            !ReferenceEquals(review.File, session.Input()) || review.Text != options.Text("text").Trim())) {
            throw new InvalidOperationException("The input or phrase changed. Find and review the matches again before removing text.");
        }
        PdfToolResult result = session.PdfTools.Execute(new PdfToolRequest(
            tool,
            session.Inputs,
            options.Text("pages", "all"),
            options.Number("perFile", 1),
            options.Number("rotation", 90),
            Enum.TryParse(options.Text("profile"), ignoreCase: true, out PdfOptimizationProfile profile) ? profile : PdfOptimizationProfile.Balanced,
            options.Text("userPassword"),
            options.Text("ownerPassword"),
            options.Text("text"),
            options.Flag("confirm")));

        // An inspection's answer is the fact sheet itself; its JSON is an extra download, not the result.
        ToolArtifact[] artifacts = session.ReplaceArtifacts((result.Artifact, tool.Kind == PdfToolKind.Inspect ? "report" : "primary"), (result.Report, "report"));
        IReadOnlyDictionary<string, string> details = result.Details ?? new Dictionary<string, string>();
        var facts = new List<ToolFact>();
        ToolVerdict verdict = tool.Kind switch {
            PdfToolKind.Inspect => InspectVerdict(details, facts),
            PdfToolKind.Optimize => OptimizeVerdict(details, facts),
            PdfToolKind.Compare => CompareVerdict(details, facts),
            PdfToolKind.Merge => new(ToolTone.Good, "Merged PDF ready" + PagesSuffix(result.PageCount), result.Summary),
            PdfToolKind.Split => new(ToolTone.Good, result.Summary.TrimEnd('.'), "The parts are packed together in one ZIP file."),
            PdfToolKind.Protect => new(ToolTone.Good, "Password added",
                "Anyone who opens this PDF now needs the password. Keep it somewhere safe: it can't be recovered."),
            PdfToolKind.Unlock => new(ToolTone.Good, "Password removed", "The copy opens without a password. Your original still has one."),
            PdfToolKind.Redact => new(ToolTone.Good, "Text removed and checked", result.Summary),
            _ => new(ToolTone.Good, "New PDF ready" + PagesSuffix(result.PageCount), result.Summary)
        };
        if (tool.Kind is not (PdfToolKind.Inspect or PdfToolKind.Optimize or PdfToolKind.Compare)) {
            if (result.PageCount is int pages) facts.Add(new ToolFact("Pages", pages.ToString(CultureInfo.InvariantCulture)));
            facts.Add(new ToolFact("Size", ToolFormat.Bytes(result.Artifact.Bytes.LongLength)));
        }

        // Success messages repeat the verdict in engine terms; only list what needs attention (plus inspection diagnostics).
        ToolItem[] items = result.Messages
            .Select((message, index) => new ToolItem("m" + index.ToString(CultureInfo.InvariantCulture), message.Title, message.Message, StateFromTone(message.ToneClass)))
            .Where(item => item.State != ToolState.Good)
            .Where(item => tool.Kind != PdfToolKind.Inspect || item.Title == "Diagnostic")
            .Where(item => tool.Kind != PdfToolKind.Compare || item.Title == "Structural difference")
            .ToArray();

        ToolPreview? preview = !result.PreviewInBrowser
            ? null
            : result.Artifact.ContentType switch {
                "application/pdf" => new ToolPreview("pdf", 0, result.PageCount),
                var type when type.StartsWith("text/html", StringComparison.OrdinalIgnoreCase) => new ToolPreview("frame", 0),
                _ => null
            };
        return new ToolResultDocument(true, verdict, facts.ToArray(), items, artifacts, preview, 0);
    }

    private static ToolResultDocument FindText(ToolSession session, ToolOptions options) {
        string text = options.Text("text").Trim();
        if (text.Length == 0) throw new InvalidOperationException("Type the text you want to remove.");
        PdfDocument source = BrowserPdfPolicy.Open(session.Input());
        var search = new PdfRedactionSearchOptions { MatchCase = false };
        search.AddLiteral(text);
        PdfRedactionPlan plan = source.Redactions.Search(search);
        session.LastReview = new RedactionReview(session.Input(), text);
        if (!plan.HasMatches) {
            return new ToolResultDocument(true,
                new ToolVerdict(ToolTone.Info, $"No matches for “{text}”", "Nothing would be removed. Check the spelling, or try a shorter phrase."),
                [new ToolFact("Matches", "0")], [], [], null, 0);
        }
        var pages = plan.Matches.GroupBy(static match => match.PageNumber).OrderBy(static group => group.Key).ToArray();
        ToolItem[] items = pages.Select(page => new ToolItem(
            "p" + page.Key.ToString(CultureInfo.InvariantCulture),
            "Page " + page.Key.ToString(CultureInfo.InvariantCulture),
            ToolFormat.Count(page.Count(), "match", "matches") + ": " + string.Join(", ", page.Select(static match => "“" + (match.Text ?? "…") + "”").Distinct().Take(4)),
            ToolState.Found,
            Page: page.Key)).ToArray();
        return new ToolResultDocument(true,
            new ToolVerdict(ToolTone.Warn,
                $"Found {ToolFormat.Count(plan.Matches.Count, "match", "matches")} on {ToolFormat.Pages(pages.Select(static page => page.Key))}",
                "These will be removed from the copy for good. Review them, then confirm."),
            [new ToolFact("Matches", plan.Matches.Count.ToString(CultureInfo.InvariantCulture)), new ToolFact("Pages", pages.Length.ToString(CultureInfo.InvariantCulture))],
            items, [], null, 0);
    }

    private static ToolVerdict InspectVerdict(IReadOnlyDictionary<string, string> details, List<ToolFact> facts) {
        bool Is(string key) => details.TryGetValue(key, out string? value) && bool.TryParse(value, out bool flag) && flag;
        int Count(string key) => details.TryGetValue(key, out string? value) && int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int count) ? count : 0;
        string? Value(string key) => details.TryGetValue(key, out string? value) ? value : null;

        int pages = Count("pages");
        if (Value("pages") is not null) facts.Add(new ToolFact("Pages", pages.ToString(CultureInfo.InvariantCulture)));
        if (Value("pdfVersion") is { } version) facts.Add(new ToolFact("PDF version", version));
        facts.Add(new ToolFact("Password", Is("encrypted") ? "Yes" : "No", Is("encrypted") ? ToolTone.Warn : null));
        facts.Add(new ToolFact("Form fields", Count("forms").ToString(CultureInfo.InvariantCulture)));
        facts.Add(new ToolFact("Signatures", Is("signatureMarkers") ? "Yes" : "No", Is("signatureMarkers") ? ToolTone.Info : null));
        facts.Add(new ToolFact("Scripts or actions", Is("activeContent") ? "Yes" : "No", Is("activeContent") ? ToolTone.Warn : null));
        facts.Add(new ToolFact("Attachments", Count("attachments").ToString(CultureInfo.InvariantCulture), Count("attachments") > 0 ? ToolTone.Info : null));
        facts.Add(new ToolFact("Tagged for accessibility", Is("tagged") ? "Yes" : "No"));
        facts.Add(new ToolFact("Comments and links", Count("annotations").ToString(CultureInfo.InvariantCulture)));

        if (!Is("canRead")) {
            return new ToolVerdict(ToolTone.Warn, "This PDF can't be fully read", "It probably needs a password. Unlock it first, then inspect it again.");
        }
        var notes = new List<string>();
        if (Is("encrypted")) notes.Add("It is password protected.");
        if (Is("activeContent")) notes.Add("It contains scripts or automatic actions.");
        if (Is("signatureMarkers")) notes.Add("It has digital signatures, which most edits will invalidate.");
        if (!Is("canRewrite")) notes.Add("Some tools can't rewrite it.");
        return new ToolVerdict(
            Is("activeContent") ? ToolTone.Warn : ToolTone.Good,
            pages > 0 ? $"A {pages.ToString(CultureInfo.InvariantCulture)}-page PDF" + (Value("pdfVersion") is { } v ? $", version {v}" : string.Empty) : "PDF inspected",
            notes.Count == 0 ? "Nothing unusual: no password, scripts or signatures." : string.Join(" ", notes));
    }

    private static ToolVerdict OptimizeVerdict(IReadOnlyDictionary<string, string> details, List<ToolFact> facts) {
        long Bytes(string key) => details.TryGetValue(key, out string? value) && long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out long bytes) ? bytes : 0;
        long before = Bytes("originalBytes"), after = Bytes("returnedBytes");
        facts.Add(new ToolFact("Before", ToolFormat.Bytes(before)));
        facts.Add(new ToolFact("After", ToolFormat.Bytes(after), after < before ? ToolTone.Good : null));
        if (after >= before || before == 0) {
            return new ToolVerdict(ToolTone.Info, "This PDF is already compact", "No smaller version was found without lowering quality, so the copy matches your original.");
        }
        double saved = 100d * (before - after) / before;
        facts.Add(new ToolFact("Saved", saved.ToString("0", CultureInfo.InvariantCulture) + "%", ToolTone.Good));
        return new ToolVerdict(ToolTone.Good,
            $"{ToolFormat.Bytes(before)} → {ToolFormat.Bytes(after)} ({saved.ToString("0", CultureInfo.InvariantCulture)}% smaller)",
            "Pages were not turned into images, so text stays sharp and selectable.");
    }

    private static ToolVerdict CompareVerdict(IReadOnlyDictionary<string, string> details, List<ToolFact> facts) {
        bool match = details.TryGetValue("isMatch", out string? value) && bool.TryParse(value, out bool flag) && flag;
        string pages = details.TryGetValue("comparedPages", out string? count) ? count : "0";
        string structural = details.TryGetValue("structuralDifferences", out string? differences) ? differences : "0";
        facts.Add(new ToolFact("Pages compared", pages));
        facts.Add(new ToolFact("Structure differences", structural, structural == "0" ? null : ToolTone.Warn));
        return match
            ? new ToolVerdict(ToolTone.Good, "The PDFs look the same", $"All {pages} compared pages match.")
            : new ToolVerdict(ToolTone.Warn, "The PDFs are different", "The comparison below highlights what changed on each page.");
    }

    private static string PagesSuffix(int? pages) => pages is int count ? " · " + ToolFormat.Count(count, "page") : string.Empty;

    private static string StateFromTone(string toneClass) => toneClass switch {
        "ocx-dot--good" => ToolState.Good,
        "ocx-dot--warn" => ToolState.Warning,
        "ocx-dot--bad" => ToolState.Bad,
        _ => ToolState.Info
    };
}
