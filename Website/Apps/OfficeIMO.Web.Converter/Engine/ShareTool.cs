using System.Globalization;
using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;
using OfficeIMO.Provenance;
using OfficeIMO.Security;
using OfficeIMO.Web.Converter.Models;
using OfficeIMO.Web.Converter.Services;
using OfficeIMO.Word;

namespace OfficeIMO.Web.Converter.Engine;

/// <summary>
/// Clean a file before sharing: one report of everything that would leave with the file, grouped by category, and one action
/// that writes a copy without the selected categories. Word, Excel and PowerPoint files go through the hidden-data API, which
/// reports and removes provenance as one of its categories, plus the Word and PowerPoint accessibility checkers. Images and
/// PDFs have no hidden-data API, so they take <see cref="OriginTool"/>'s provenance route.
/// </summary>
internal static class ShareTool {
    internal static readonly string[] OfficeExtensions = [".docx", ".xlsx", ".pptx"];
    internal const string PersonalGroup = "Personal and hidden data";
    internal const string AccessibilityGroup = "Accessibility";
    internal const string TableHeaders = "fix-table-headers";

    internal static ToolResultDocument Run(ToolSession session, string action, ToolOptions options) {
        SelectedDocument file = session.Input();
        if (!OfficeExtensions.Contains(file.Extension, StringComparer.OrdinalIgnoreCase)) return OriginTool.Run(session, action, options);
        return action switch {
            "inspect" => Inspect(session, file),
            "remove" => Remove(session, file, options),
            _ => throw new NotSupportedException($"The clean-before-sharing tool has no '{action}' action.")
        };
    }

    /// <summary>Category id, page title, group, and what selecting it does. Default categories start selected.</summary>
    private sealed record Category(OfficeHiddenDataKind Kind, string Id, string Title, string Explanation, bool Preselected, string Group = PersonalGroup);

    private static readonly Category[] Categories = [
        new(OfficeHiddenDataKind.DocumentProperties, "properties", "Document properties",
            "Author, last saved by, company, manager, template and editing time.", true),
        new(OfficeHiddenDataKind.CustomProperties, "custom-properties", "Custom properties",
            "Extra name and value pairs added by people or other apps.", true),
        new(OfficeHiddenDataKind.Comments, "comments", "Comments and notes",
            "Comments, replies and notes, with the text they point to left in place.", true),
        new(OfficeHiddenDataKind.PersonalIdentities, "people", "Names of people",
            "Names on comments and changes that stay in the file are replaced with \"Author\".", true),
        new(OfficeHiddenDataKind.TrackedChanges, "tracked-changes", "Tracked changes",
            "Selecting this accepts every change, as Accept All does in Word. Left unselected, changes stay for the reader to see.", false),
        new(OfficeHiddenDataKind.HiddenContent, "hidden-content", "Hidden content",
            "Hidden text, sheets, rows, columns, slides and shapes. Removing it changes what readers get, and deleting hidden rows or sheets can break formulas that use them.", false),
        new(OfficeHiddenDataKind.SpeakerNotes, "speaker-notes", "Speaker notes",
            "Notes under each slide.", true),
        new(OfficeHiddenDataKind.Thumbnail, "thumbnail", "Preview picture",
            "A small picture of the first page that file browsers show.", true),
        new(OfficeHiddenDataKind.CustomXml, "custom-xml", "Custom XML data",
            "Data stored by add-ins and document management systems. Bibliography sources and data bound to content controls stay.", true),
        new(OfficeHiddenDataKind.ImageMetadata, "image-metadata", "Photo details",
            "Camera, location (GPS) and editing details inside pictures. The pictures themselves don't change.", true),
        new(OfficeHiddenDataKind.PrinterSettings, "printer-settings", "Printer settings",
            "Printer names and setup saved with the file.", true),
        new(OfficeHiddenDataKind.ExternalReferences, "external-references", "Links to other files",
            "Paths to templates and linked workbooks. Links that formulas or refreshes need are reported and kept.", true),
        new(OfficeHiddenDataKind.DocumentVariables, "document-variables", "Document variables",
            "Values stored by macros and templates.", true),
        new(OfficeHiddenDataKind.RevisionIdentifiers, "revision-ids", "Editing session IDs",
            "Codes Word adds for each editing session, which can link copies of a file.", true),
        new(OfficeHiddenDataKind.Provenance, "provenance", "Origin records",
            "Records that say how content was made, such as a picture marked as AI-generated. Kept unless you select them, because they are a disclosure to the reader.", false,
            OriginTool.Group)
    ];

    private static Category For(OfficeHiddenDataKind kind) => Categories.First(category => category.Kind == kind);

    private static ToolResultDocument Inspect(ToolSession session, SelectedDocument file) {
        (OfficeHiddenDataReport report, Accessibility? accessibility) = Load(file, editable: false,
            static document => (document.Inspect(), document.InspectAccessibility()));
        session.LastReview = report;
        session.ReplaceArtifacts();

        var items = new List<ToolItem>();
        foreach (Category category in Categories) {
            IReadOnlyList<OfficeHiddenDataFinding> findings = report.GetFindings(category.Kind);
            if (findings.Count == 0) continue;
            string detail = category.Kind == OfficeHiddenDataKind.Provenance ? category.Explanation + " " + ProvenanceExamples(findings) : category.Explanation + " " + Examples(findings);
            items.Add(new ToolItem(category.Id, Title(category, findings.Count), detail, ToolState.Found, category.Group,
                Selectable: true, Selected: category.Preselected));
        }
        if (report.Findings.Any(static finding => finding.Kind == OfficeHiddenDataKind.Provenance && IsPackageCredential(finding))) {
            items.Add(new ToolItem("package-credential", "Content Credentials for the whole file",
                "These cover every byte of the file, so any change breaks them. Saving a cleaned copy always removes them, even if you keep the other origin records.",
                ToolState.Warning, OriginTool.Group));
        }
        Func<OfficeProvenanceEvidence, string?> assetOf = AssetResolver(report);
        OfficeProvenanceEvidence[] manifests = report.GetFindings(OfficeHiddenDataKind.Provenance)
            .Select(static finding => finding.Provenance).OfType<OfficeProvenanceEvidence>()
            .Where(static evidence => evidence.Manifest != null).Take(3).ToArray();
        for (int index = 0; index < manifests.Length; index++) {
            items.AddRange(OriginTool.Explain(manifests[index].Manifest!, index, manifests.Length, manifests[index], assetOf));
        }
        if (accessibility != null) items.AddRange(AccessibilityItems(accessibility));

        var facts = new List<ToolFact> { new("Format", FormatName(file.Extension)) };
        string? author = report.GetFindings(OfficeHiddenDataKind.DocumentProperties).Concat(report.GetFindings(OfficeHiddenDataKind.PersonalIdentities))
            .Select(static finding => finding.Value).FirstOrDefault(static value => !string.IsNullOrWhiteSpace(value));
        if (author != null) facts.Add(new ToolFact("Name in the file", author, ToolTone.Warn));
        if (report.GetFindings(OfficeHiddenDataKind.ImageMetadata).Any(static finding => Mentions(finding, "GPS"))) {
            facts.Add(new ToolFact("Photo location", "Stored in a picture", ToolTone.Warn));
        }
        if (report.GetFindings(OfficeHiddenDataKind.TrackedChanges).Count is int changes and > 0) {
            facts.Add(new ToolFact("Tracked changes", changes.ToString(CultureInfo.InvariantCulture), ToolTone.Warn));
        }
        if (report.GetFindings(OfficeHiddenDataKind.Provenance).Any(static finding => finding.Provenance?.DigitalSourceKind is
                OfficeProvenanceDigitalSourceKind.TrainedAlgorithmicMedia or OfficeProvenanceDigitalSourceKind.CompositeWithTrainedAlgorithmicMedia
            || finding.Provenance?.Manifest?.DeclaresGenerativeAi == true)) {
            facts.Add(new ToolFact("Generative AI", "Declared in an origin record", ToolTone.Warn));
        }
        if (accessibility != null) {
            facts.Add(new ToolFact("Accessibility", accessibility.Errors == 0 ? "No errors" : ToolFormat.Count(accessibility.Errors, "error"),
                accessibility.Errors == 0 ? ToolTone.Good : ToolTone.Warn));
        }

        int kinds = items.Count(static item => item.Selectable && item.Id != TableHeaders);
        OfficeProvenanceReport provenance = ProvenanceView(report);
        string origin = provenance.Evidence.Count > 0 ? " " + OriginTool.Summary(provenance, OriginTool.ManifestSummaries(provenance), assetOf) : string.Empty;
        ToolVerdict verdict = kinds > 0
            ? new ToolVerdict(ToolTone.Warn, "Found " + ToolFormat.Count(kinds, "kind") + " of data that would leave with this file",
                "Ticked items are removed from a copy. Tracked changes, hidden content and origin records stay unless you tick them, because removing them changes what readers see or what the file tells them." + origin)
            : new ToolVerdict(ToolTone.Good, "Nothing hidden found",
                "This file has no document properties, comments, hidden content, picture details or origin records that the tool looks for.");
        return new ToolResultDocument(true, verdict, facts.ToArray(), items.ToArray(), [], null, 0);
    }

    /// <summary>Accessibility findings for Word and PowerPoint. Only the safe Word fix is offered as a choice; the rest need a person.</summary>
    private static IEnumerable<ToolItem> AccessibilityItems(Accessibility accessibility) {
        if (accessibility.CanFixTableHeaders) {
            yield return new ToolItem(TableHeaders, "Mark table header rows",
                "Marks the first row of each data table as a header row, so screen readers announce column names. Tables that look like layout tables are left alone.",
                ToolState.Found, AccessibilityGroup, Selectable: true, Selected: false);
        }
        int shown = 0;
        foreach (IGrouping<string, (string Code, string Message, bool Error)> rule in accessibility.Findings.GroupBy(static finding => finding.Code)) {
            if (shown++ == 8) break;
            (string code, string message, bool error) = rule.First();
            string count = rule.Count() > 1 ? $" ({rule.Count()})" : string.Empty;
            yield return new ToolItem("a11y-" + code, AccessibilityTitle(code) + count, message + (rule.Count() > 1 ? " The same applies elsewhere in the file." : string.Empty),
                error ? ToolState.Warning : ToolState.Info, AccessibilityGroup);
        }
        if (accessibility.Findings.Count == 0) {
            yield return new ToolItem("a11y-ok", "No accessibility problems found", "Titles, language, alternative text, headings, tables, links and contrast passed the checks.", ToolState.Good, AccessibilityGroup);
        }
    }

    private static string AccessibilityTitle(string code) => code switch {
        "accessibility-missing-document-title" => "No document title",
        "accessibility-missing-language" => "No document language",
        "accessibility-missing-alternative-text" => "Pictures without alternative text",
        "accessibility-heading-level-skipped" => "Heading levels skipped",
        "accessibility-missing-headings" => "Long document without headings",
        "accessibility-missing-table-header" => "Tables without a header row",
        "accessibility-layout-table" => "Table used for layout",
        "accessibility-link-text-missing" or "accessibility-unclear-link-label" => "Unclear link text",
        "accessibility-low-contrast" => "Low text contrast",
        "accessibility-floating-object" => "Floating picture",
        "accessibility-blank-paragraph-spacing" => "Empty paragraphs used for spacing",
        "accessibility-missing-slide-title" => "Slides without a title",
        "accessibility-missing-shape-title" => "Shapes without a title",
        "accessibility-missing-link-tooltip" => "Links without a tooltip",
        "accessibility-chart-color-only-meaning" => "Chart relies on colour alone",
        _ => code
    };

    private static ToolResultDocument Remove(ToolSession session, SelectedDocument file, ToolOptions options) {
        IReadOnlySet<string> selected = options.List("remove");
        if (selected.Count == 0) throw new InvalidOperationException("Select at least one item to remove or fix.");
        OfficeHiddenDataKind kinds = OfficeHiddenDataKind.None;
        foreach (Category category in Categories) {
            if (selected.Contains(category.Id)) kinds |= category.Kind;
        }
        bool fixHeaders = selected.Contains(TableHeaders);
        var removal = new OfficeHiddenDataOptions {
            Kinds = kinds,
            TrackedChanges = (kinds & OfficeHiddenDataKind.TrackedChanges) != 0 ? OfficeTrackedChangesAction.AcceptAll : OfficeTrackedChangesAction.Keep
        };

        byte[]? output = null;
        int headersBefore = 0, headersAfter = 0;
        OfficeHiddenDataReport report = Load(file, editable: true, document => {
            OfficeHiddenDataReport result = document.Remove(removal);
            if (fixHeaders) (headersBefore, headersAfter) = document.FixTableHeaders();
            output = document.Save();
            return result;
        });
        OfficeHiddenDataReport after = Load(new SelectedDocument(file.Name, file.Extension, file.FormatLabel, output!.LongLength, output), editable: false,
            static document => document.Inspect());

        string outputName = Path.GetFileNameWithoutExtension(file.Name) + "-cleaned" + file.Extension;
        ToolArtifact[] artifacts = session.ReplaceArtifacts((new BrowserConversionArtifact(output, outputName, ContentType(file.Extension)), "primary"));

        var items = new List<ToolItem>();
        int changed = 0, kept = 0;
        foreach (Category category in Categories) {
            OfficeHiddenDataFinding[] findings = report.GetFindings(category.Kind).ToArray();
            if (findings.Length == 0) continue;
            OfficeHiddenDataFinding[] done = findings.Where(static finding => finding.Outcome is not (OfficeHiddenDataOutcome.Found or OfficeHiddenDataOutcome.Retained)).ToArray();
            OfficeHiddenDataFinding[] retained = findings.Where(static finding => finding.Outcome == OfficeHiddenDataOutcome.Retained).ToArray();
            changed += done.Length;
            kept += retained.Length;
            if (done.Length > 0) {
                items.Add(new ToolItem(category.Id, Title(category, done.Length), Outcomes(done), ToolState.Removed, category.Group));
            }
            if (retained.Length > 0) {
                bool byChoice = !selected.Contains(category.Id);
                string detail = byChoice
                    ? "Kept because you chose to keep it."
                    : "Kept: " + string.Join(" ", retained.Select(static finding => Plain(finding.Note)).Where(static note => !string.IsNullOrWhiteSpace(note)).Distinct().Take(3));
                items.Add(new ToolItem(category.Id + "-kept", Title(category, retained.Length), detail.TrimEnd(), byChoice ? ToolState.Kept : ToolState.Info, category.Group));
            }
        }
        if (fixHeaders) {
            int marked = Math.Max(0, headersBefore - headersAfter);
            items.Add(new ToolItem(TableHeaders, "Table header rows", marked > 0
                ? ToolFormat.Count(marked, "table") + " now " + (marked == 1 ? "has" : "have") + " a header row."
                : "No table needed a header row, or every remaining one looks like a layout table.", marked > 0 ? ToolState.Good : ToolState.Info, AccessibilityGroup));
        }
        bool watermark = (kinds & OfficeHiddenDataKind.Provenance) != 0 && ProvenanceView(report).Evidence
            .Select(static evidence => evidence.Manifest).OfType<OfficeC2paManifestSummary>().Any(static manifest => manifest.Actions.Any(static action => OriginTool.Watermarked(action.Action)));
        if (watermark) items.Add(new ToolItem("watermark", "Watermark in the content", OriginTool.WatermarkNote, ToolState.Warning, OriginTool.Group));
        items.Add(new ToolItem("content", "Visible content", (kinds & (OfficeHiddenDataKind.HiddenContent | OfficeHiddenDataKind.TrackedChanges)) == 0
            ? "Not changed. Only the items you ticked were touched."
            : "Changed only where you chose to remove hidden content or accept tracked changes.", ToolState.Kept));

        // Names on comments and changes that stay are anonymized, not removed, and retained items stay with a reason;
        // anything else in a selected category that the re-check still finds is reported as not removed.
        var explained = new HashSet<OfficeHiddenDataKind>(report.Findings.Where(static finding => finding.Outcome == OfficeHiddenDataOutcome.Retained).Select(static finding => finding.Kind));
        int stillThere = after.Findings.Count(finding => selected.Contains(For(finding.Kind).Id) && finding.Kind != OfficeHiddenDataKind.TrackedChanges &&
            !explained.Contains(finding.Kind) && !(finding.Kind == OfficeHiddenDataKind.PersonalIdentities && finding.Value == removal.AnonymousAuthor));
        const int retainedSelected = 0;
        string title = changed == 0 && !fixHeaders ? "Nothing was removed" : "Clean copy ready" + (changed > 0 ? " · " + ToolFormat.Count(changed, "item") + " removed or anonymized" : string.Empty);
        string detailText = stillThere > retainedSelected
            ? "We checked the copy again and some selected items are still there. See the list below."
            : kept > 0
                ? "We checked the copy again. The items shown as kept remain, each with the reason."
                : "We checked the copy again and found none of the selected items.";
        if (watermark) detailText += " The watermark a removed record mentions is part of the content and is still there.";
        var verdict = new ToolVerdict(changed == 0 && !fixHeaders || stillThere > retainedSelected ? ToolTone.Warn : ToolTone.Good, title, detailText);
        return new ToolResultDocument(true, verdict,
            [new ToolFact("Removed", changed.ToString(CultureInfo.InvariantCulture), changed > 0 ? ToolTone.Good : null),
             new ToolFact("Kept", kept.ToString(CultureInfo.InvariantCulture), kept > 0 ? ToolTone.Warn : null),
             new ToolFact("Size", ToolFormat.Bytes(output.LongLength))],
            items.ToArray(), artifacts, null, 0);
    }

    /// <summary>The hidden-data report's provenance findings, viewed as the provenance report the origin explanations read.</summary>
    private static OfficeProvenanceReport ProvenanceView(OfficeHiddenDataReport report) => new(OfficeProvenanceAssetFormat.ZipPackage,
        report.GetFindings(OfficeHiddenDataKind.Provenance).Select(static finding => finding.Provenance).OfType<OfficeProvenanceEvidence>().ToList());

    /// <summary>
    /// Hidden-data findings locate a picture's provenance as "/word/media/x.png/" + the carrier's own location;
    /// this maps each evidence back to its picture so the explanations can say which one a record belongs to.
    /// </summary>
    private static Func<OfficeProvenanceEvidence, string?> AssetResolver(OfficeHiddenDataReport report) {
        var assets = new Dictionary<OfficeProvenanceEvidence, string>(ReferenceEqualityComparer.Instance);
        foreach (OfficeHiddenDataFinding finding in report.GetFindings(OfficeHiddenDataKind.Provenance)) {
            if (finding.Provenance is not { } evidence) continue;
            string suffix = "/" + evidence.Location;
            if (finding.Location.Length > suffix.Length && finding.Location.EndsWith(suffix, StringComparison.Ordinal)) {
                assets[evidence] = finding.Location.Substring(0, finding.Location.Length - suffix.Length).TrimStart('/');
            }
        }
        return evidence => assets.TryGetValue(evidence, out string? asset) ? asset : null;
    }

    private static string ProvenanceExamples(IReadOnlyList<OfficeHiddenDataFinding> findings) {
        string[] places = findings.Select(static finding => finding.Provenance is { } evidence && finding.Location.EndsWith("/" + evidence.Location, StringComparison.Ordinal)
                ? "embedded asset " + finding.Location.Substring(0, finding.Location.Length - evidence.Location.Length - 1).TrimStart('/')
                : IsPackageCredential(finding) ? "the whole file" : finding.Location)
            .Distinct(StringComparer.Ordinal).Take(3).ToArray();
        return "Found in " + string.Join("; ", places) + (findings.Count > places.Length ? $" and {findings.Count - places.Length} more" : string.Empty) + ".";
    }

    private static string Title(Category category, int count) => count > 1 ? $"{category.Title} ({count})" : category.Title;

    /// <summary>A short, readable sample of what was found, such as the names or properties.</summary>
    private static string Examples(IReadOnlyList<OfficeHiddenDataFinding> findings) {
        string[] samples = findings.Select(static finding => finding.Value is { Length: > 0 } value ? $"{finding.Description}: {Clip(value)}" : finding.Description)
            .Distinct(StringComparer.Ordinal).Take(3).ToArray();
        string more = findings.Count > samples.Length ? $" and {findings.Count - samples.Length} more" : string.Empty;
        string text = "Found: " + string.Join("; ", samples) + more;
        return text.EndsWith('…') || text.EndsWith('.') ? text : text + ".";
    }

    private static string Outcomes(OfficeHiddenDataFinding[] findings) {
        var parts = findings.GroupBy(static finding => finding.Outcome).Select(static group => group.Key switch {
            OfficeHiddenDataOutcome.Anonymized => ToolFormat.Count(group.Count(), "name") + " replaced with an anonymous author",
            OfficeHiddenDataOutcome.Accepted => ToolFormat.Count(group.Count(), "change") + " accepted",
            OfficeHiddenDataOutcome.Rejected => ToolFormat.Count(group.Count(), "change") + " rejected",
            _ => ToolFormat.Count(group.Count(), "item") + " removed"
        });
        string notes = string.Join(" ", findings.Select(static finding => Plain(finding.Note)).Where(static note => !string.IsNullOrWhiteSpace(note)).Distinct().Take(2));
        return (string.Join(", ", parts) + "." + (notes.Length > 0 ? " " + notes : string.Empty)).Trim();
    }

    /// <summary>The engine's notes name API categories; the page names them the way its checkboxes do.</summary>
    private static string? Plain(string? note) => note?
        .Replace("select Provenance to remove both", "tick Content Credentials and AI labels to remove both", StringComparison.Ordinal)
        .Replace("select Provenance to remove it", "tick Content Credentials and AI labels to remove it", StringComparison.Ordinal)
        .Replace(" Use RemoveProvenance on the file to change only provenance.", string.Empty, StringComparison.Ordinal)
        .Replace("Provenance was not selected", "Content Credentials and AI labels were not ticked", StringComparison.Ordinal);

    private static bool Mentions(OfficeHiddenDataFinding finding, string text) =>
        finding.Description.Contains(text, StringComparison.OrdinalIgnoreCase);

    private static bool IsPackageCredential(OfficeHiddenDataFinding finding) =>
        finding.Location.Contains("content_credential.c2pa", StringComparison.OrdinalIgnoreCase);

    private static string? Clip(string? value) {
        if (string.IsNullOrWhiteSpace(value)) return null;
        string text = new string(value.Trim().Where(static ch => !char.IsControl(ch)).ToArray());
        return text.Length <= 60 ? text : text.Substring(0, 59) + "…";
    }

    /// <summary>One loaded Word, Excel or PowerPoint document behind the same three calls.</summary>
    /// <summary>Accessibility findings in one shape for Word and PowerPoint. <see cref="CanFixTableHeaders"/> is Word's safe fix.</summary>
    private sealed record Accessibility(IReadOnlyList<(string Code, string Message, bool Error)> Findings, bool CanFixTableHeaders) {
        internal int Errors => Findings.Count(static finding => finding.Error);
    }

    /// <summary>One loaded Word, Excel or PowerPoint document behind the same calls.</summary>
    private interface IShareDocument {
        OfficeHiddenDataReport Inspect();
        OfficeHiddenDataReport Remove(OfficeHiddenDataOptions options);
        Accessibility? InspectAccessibility();
        (int Before, int After) FixTableHeaders();
        byte[] Save();
    }

    private sealed class WordHost(WordDocument document) : IShareDocument {
        private const string MissingHeader = "accessibility-missing-table-header";
        public OfficeHiddenDataReport Inspect() => document.InspectHiddenData();
        public OfficeHiddenDataReport Remove(OfficeHiddenDataOptions options) => document.RemoveHiddenData(options);
        public Accessibility InspectAccessibility() {
            WordAccessibilityReport report = document.InspectAccessibility();
            return new Accessibility(report.Findings.Select(static finding => (finding.Code, finding.Message, finding.Severity == WordAccessibilitySeverity.Error)).ToList(),
                report.Findings.Any(static finding => finding.Code == MissingHeader));
        }
        public (int Before, int After) FixTableHeaders() {
            int before = document.InspectAccessibility().Findings.Count(static finding => finding.Code == MissingHeader);
            int after = document.FixAccessibility(new WordAccessibilityFixOptions { MarkTableHeaderRows = true }).Findings.Count(static finding => finding.Code == MissingHeader);
            return (before, after);
        }
        public byte[] Save() { using var stream = new MemoryStream(); document.Save(stream); return stream.ToArray(); }
    }

    private sealed class ExcelHost(ExcelDocument document) : IShareDocument {
        public OfficeHiddenDataReport Inspect() => document.InspectHiddenData();
        public OfficeHiddenDataReport Remove(OfficeHiddenDataOptions options) => document.RemoveHiddenData(options);
        public Accessibility? InspectAccessibility() => null;
        public (int Before, int After) FixTableHeaders() => (0, 0);
        public byte[] Save() { using var stream = new MemoryStream(); document.Save(stream); return stream.ToArray(); }
    }

    private sealed class PowerPointHost(PowerPointDocument document) : IShareDocument {
        public OfficeHiddenDataReport Inspect() => document.InspectHiddenData();
        public OfficeHiddenDataReport Remove(OfficeHiddenDataOptions options) => document.RemoveHiddenData(options);
        public Accessibility InspectAccessibility() => new(document.InspectAccessibility().Findings
            .Select(static finding => (finding.Code, finding.Message, finding.Severity == PowerPointAccessibilitySeverity.Error)).ToList(), false);
        public (int Before, int After) FixTableHeaders() => (0, 0);
        public byte[] Save() { using var stream = new MemoryStream(); document.Save(stream); return stream.ToArray(); }
    }

    private static T Load<T>(SelectedDocument file, bool editable, Func<IShareDocument, T> use) {
        if (editable && OfficePackageSignatureService.Inspect(file.Bytes).SignatureParts.Count > 0) {
            throw new NotSupportedException("This file is digitally signed. Removing hidden data would break the signature, so the tool won't change it. Clean the file before you sign it.");
        }
        // Editable loads work on a private expandable copy; the visitor's original bytes are never written.
        using var stream = new MemoryStream();
        stream.Write(file.Bytes, 0, file.Bytes.Length);
        stream.Position = 0;
        DocumentAccessMode access = editable ? DocumentAccessMode.ReadWrite : DocumentAccessMode.ReadOnly;
        OfficePackageSecurityOptions security = BrowserConversionService.CreateBrowserPackageSecurity();
        switch (file.Extension) {
            case ".docx": {
                using WordDocument document = WordDocument.Load(stream, new WordLoadOptions { AccessMode = access, PackageSecurity = security });
                return use(new WordHost(document));
            }
            case ".xlsx": {
                using ExcelDocument document = ExcelDocument.Load(stream, new ExcelLoadOptions { AccessMode = access, PackageSecurity = security });
                return use(new ExcelHost(document));
            }
            default: {
                using PowerPointDocument document = PowerPointDocument.Load(stream, new PowerPointLoadOptions { AccessMode = access, PackageSecurity = security });
                return use(new PowerPointHost(document));
            }
        }
    }

    private static string FormatName(string extension) => extension switch {
        ".docx" => "Word document",
        ".xlsx" => "Excel workbook",
        _ => "PowerPoint presentation"
    };

    private static string ContentType(string extension) => extension switch {
        ".docx" => "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        ".xlsx" => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        _ => "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    };
}
