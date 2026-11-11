using System.Globalization;
using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;
using OfficeIMO.Web.Converter.Services;
using OfficeIMO.Word;

namespace OfficeIMO.Web.Converter.Engine;

/// <summary>Finds metadata and hidden data in a Word, Excel or PowerPoint file, and writes a copy without the categories the visitor selects.</summary>
internal static class HiddenDataTool {
    internal static readonly string[] Extensions = [".docx", ".xlsx", ".pptx"];

    internal static ToolResultDocument Run(ToolSession session, string action, ToolOptions options) => action switch {
        "inspect" => Inspect(session),
        "remove" => Remove(session, options),
        _ => throw new NotSupportedException($"The hidden data tool has no '{action}' action.")
    };

    /// <summary>Category id, page title, and what selecting it does. Default categories start selected.</summary>
    private sealed record Category(OfficeHiddenDataKind Kind, string Id, string Title, string Explanation, bool Preselected);

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
        new(OfficeHiddenDataKind.Provenance, "provenance", "Origin records",
            "Content Credentials and AI source labels that say how content was made. Kept unless you select them.", false),
        new(OfficeHiddenDataKind.PrinterSettings, "printer-settings", "Printer settings",
            "Printer names and setup saved with the file.", true),
        new(OfficeHiddenDataKind.ExternalReferences, "external-references", "Links to other files",
            "Paths to templates and linked workbooks. Links that formulas or refreshes need are reported and kept.", true),
        new(OfficeHiddenDataKind.DocumentVariables, "document-variables", "Document variables",
            "Values stored by macros and templates.", true),
        new(OfficeHiddenDataKind.RevisionIdentifiers, "revision-ids", "Editing session IDs",
            "Codes Word adds for each editing session, which can link copies of a file.", true)
    ];

    private static Category For(OfficeHiddenDataKind kind) => Categories.First(category => category.Kind == kind);

    private static ToolResultDocument Inspect(ToolSession session) {
        SelectedDocument file = session.Input();
        EnsureSupported(file);
        OfficeHiddenDataReport report = Load(file, editable: false, static document => document.Inspect());
        session.LastReview = report;
        session.ReplaceArtifacts();

        var items = new List<ToolItem>();
        foreach (Category category in Categories) {
            IReadOnlyList<OfficeHiddenDataFinding> findings = report.GetFindings(category.Kind);
            if (findings.Count == 0) continue;
            items.Add(new ToolItem(category.Id, Title(category, findings.Count), category.Explanation + " " + Examples(findings),
                ToolState.Found, Selectable: true, Selected: category.Preselected));
        }
        if (report.Findings.Any(static finding => finding.Kind == OfficeHiddenDataKind.Provenance && IsPackageCredential(finding))) {
            items.Add(new ToolItem("package-credential", "Content Credentials for the whole file",
                "These cover every byte of the file, so any change breaks them. Saving a cleaned copy always removes them, even if you keep other origin records.",
                ToolState.Warning));
        }

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

        ToolVerdict verdict = report.HasHiddenData
            ? new ToolVerdict(ToolTone.Warn, "Found " + ToolFormat.Count(items.Count(static item => item.Selectable), "kind") + " of hidden data",
                "Selected items are removed from a copy. Tracked changes, hidden content and origin records are kept unless you select them, because removing them changes what readers see or what the file discloses.")
            : new ToolVerdict(ToolTone.Good, "No hidden data found",
                "This file has no document properties, comments, hidden content or picture details that the tool looks for.");
        return new ToolResultDocument(true, verdict, facts.ToArray(), items.ToArray(), [], null, 0);
    }

    private static ToolResultDocument Remove(ToolSession session, ToolOptions options) {
        SelectedDocument file = session.Input();
        EnsureSupported(file);
        IReadOnlySet<string> selected = options.List("remove");
        if (selected.Count == 0) throw new InvalidOperationException("Select at least one kind of hidden data to remove.");
        OfficeHiddenDataKind kinds = OfficeHiddenDataKind.None;
        foreach (Category category in Categories) {
            if (selected.Contains(category.Id)) kinds |= category.Kind;
        }
        var removal = new OfficeHiddenDataOptions {
            Kinds = kinds,
            TrackedChanges = (kinds & OfficeHiddenDataKind.TrackedChanges) != 0 ? OfficeTrackedChangesAction.AcceptAll : OfficeTrackedChangesAction.Keep
        };

        byte[]? output = null;
        OfficeHiddenDataReport report = Load(file, editable: true, document => {
            OfficeHiddenDataReport result = document.Remove(removal);
            output = document.Save();
            return result;
        });
        OfficeHiddenDataReport after = Load(new SelectedDocument(file.Name, file.Extension, file.Format, output!.LongLength, output), editable: false,
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
                items.Add(new ToolItem(category.Id, Title(category, done.Length), Outcomes(done), ToolState.Removed));
            }
            if (retained.Length > 0) {
                bool byChoice = !selected.Contains(category.Id);
                string detail = byChoice
                    ? "Kept because you chose to keep it."
                    : "Kept: " + string.Join(" ", retained.Select(static finding => finding.Note).Where(static note => !string.IsNullOrWhiteSpace(note)).Distinct().Take(3));
                items.Add(new ToolItem(category.Id + "-kept", Title(category, retained.Length), detail.TrimEnd(), byChoice ? ToolState.Kept : ToolState.Info));
            }
        }
        items.Add(new ToolItem("content", "Visible content", (kinds & (OfficeHiddenDataKind.HiddenContent | OfficeHiddenDataKind.TrackedChanges)) == 0
            ? "Not changed. Only the hidden data you selected was touched."
            : "Changed only where you chose to remove hidden content or accept tracked changes.", ToolState.Kept));

        int stillThere = after.Findings.Count(finding => selected.Contains(For(finding.Kind).Id) && finding.Kind != OfficeHiddenDataKind.TrackedChanges);
        int retainedSelected = report.Findings.Count(finding => finding.Outcome == OfficeHiddenDataOutcome.Retained && selected.Contains(For(finding.Kind).Id));
        string title = changed == 0 ? "Nothing was removed" : "Clean copy ready · " + ToolFormat.Count(changed, "item") + " removed or anonymized";
        string detailText = stillThere > retainedSelected
            ? "We checked the copy again and some selected items are still there. See the list below."
            : kept > 0
                ? "We checked the copy again. The items shown as kept remain, each with the reason."
                : "We checked the copy again and found none of the selected hidden data.";
        var verdict = new ToolVerdict(changed == 0 || stillThere > retainedSelected ? ToolTone.Warn : ToolTone.Good, title, detailText);
        return new ToolResultDocument(true, verdict,
            [new ToolFact("Removed", changed.ToString(CultureInfo.InvariantCulture), changed > 0 ? ToolTone.Good : null),
             new ToolFact("Kept", kept.ToString(CultureInfo.InvariantCulture), kept > 0 ? ToolTone.Warn : null),
             new ToolFact("Size", ToolFormat.Bytes(output.LongLength))],
            items.ToArray(), artifacts, null, 0);
    }

    private static string Title(Category category, int count) => count > 1 ? $"{category.Title} ({count})" : category.Title;

    /// <summary>A short, readable sample of what was found, such as the names or properties.</summary>
    private static string Examples(IReadOnlyList<OfficeHiddenDataFinding> findings) {
        string[] samples = findings.Select(static finding => finding.Value is { Length: > 0 } value ? $"{finding.Description}: {Clip(value)}" : finding.Description)
            .Distinct(StringComparer.Ordinal).Take(3).ToArray();
        string more = findings.Count > samples.Length ? $" and {findings.Count - samples.Length} more" : string.Empty;
        return "Found: " + string.Join("; ", samples) + more + ".";
    }

    private static string Outcomes(OfficeHiddenDataFinding[] findings) {
        var parts = findings.GroupBy(static finding => finding.Outcome).Select(static group => group.Key switch {
            OfficeHiddenDataOutcome.Anonymized => ToolFormat.Count(group.Count(), "name") + " replaced with an anonymous author",
            OfficeHiddenDataOutcome.Accepted => ToolFormat.Count(group.Count(), "change") + " accepted",
            OfficeHiddenDataOutcome.Rejected => ToolFormat.Count(group.Count(), "change") + " rejected",
            _ => ToolFormat.Count(group.Count(), "item") + " removed"
        });
        string notes = string.Join(" ", findings.Select(static finding => finding.Note).Where(static note => !string.IsNullOrWhiteSpace(note)).Distinct().Take(2));
        return (string.Join(", ", parts) + "." + (notes.Length > 0 ? " " + notes : string.Empty)).Trim();
    }

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
    private interface IHiddenDataDocument {
        OfficeHiddenDataReport Inspect();
        OfficeHiddenDataReport Remove(OfficeHiddenDataOptions options);
        byte[] Save();
    }

    private sealed class WordHost(WordDocument document) : IHiddenDataDocument {
        public OfficeHiddenDataReport Inspect() => document.InspectHiddenData();
        public OfficeHiddenDataReport Remove(OfficeHiddenDataOptions options) => document.RemoveHiddenData(options);
        public byte[] Save() { using var stream = new MemoryStream(); document.Save(stream); return stream.ToArray(); }
    }

    private sealed class ExcelHost(ExcelDocument document) : IHiddenDataDocument {
        public OfficeHiddenDataReport Inspect() => document.InspectHiddenData();
        public OfficeHiddenDataReport Remove(OfficeHiddenDataOptions options) => document.RemoveHiddenData(options);
        public byte[] Save() { using var stream = new MemoryStream(); document.Save(stream); return stream.ToArray(); }
    }

    private sealed class PowerPointHost(PowerPointDocument document) : IHiddenDataDocument {
        public OfficeHiddenDataReport Inspect() => document.InspectHiddenData();
        public OfficeHiddenDataReport Remove(OfficeHiddenDataOptions options) => document.RemoveHiddenData(options);
        public byte[] Save() { using var stream = new MemoryStream(); document.Save(stream); return stream.ToArray(); }
    }

    private static T Load<T>(SelectedDocument file, bool editable, Func<IHiddenDataDocument, T> use) {
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

    private static void EnsureSupported(SelectedDocument file) {
        if (!Extensions.Contains(file.Extension, StringComparer.OrdinalIgnoreCase)) {
            throw new NotSupportedException("Choose a Word (.docx), Excel (.xlsx) or PowerPoint (.pptx) file.");
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
