using OfficeIMO.Markdown;

namespace OfficeIMO.OneNote.Markdown;

/// <summary>One loss or compatibility diagnostic produced by semantic OneNote-to-Markdown projection.</summary>
public sealed class OneNoteMarkdownDiagnostic {
    internal OneNoteMarkdownDiagnostic(string code, OneNoteDiagnosticSeverity severity, string source, string message) {
        Code = code;
        Severity = severity;
        Source = source;
        Message = message;
    }

    /// <summary>Stable diagnostic code.</summary>
    public string Code { get; }

    /// <summary>Diagnostic severity.</summary>
    public OneNoteDiagnosticSeverity Severity { get; }

    /// <summary>Notebook, section, page, or source path associated with the diagnostic.</summary>
    public string Source { get; }

    /// <summary>Human-readable diagnostic message.</summary>
    public string Message { get; }
}

/// <summary>A Markdown document paired with explicit semantic-projection diagnostics.</summary>
public sealed class OneNoteMarkdownConversionResult {
    internal OneNoteMarkdownConversionResult(MarkdownDoc value, IReadOnlyList<OneNoteMarkdownDiagnostic> diagnostics) {
        Value = value;
        Diagnostics = diagnostics;
    }

    /// <summary>Projected Markdown document.</summary>
    public MarkdownDoc Value { get; }

    /// <summary>Source and projection diagnostics captured for this operation.</summary>
    public IReadOnlyList<OneNoteMarkdownDiagnostic> Diagnostics { get; }

    /// <summary>True when projection reported an approximation, omission, or error.</summary>
    public bool HasLoss => Diagnostics.Any(diagnostic => diagnostic.Severity != OneNoteDiagnosticSeverity.Information);
}

internal static class OneNoteMarkdownDiagnosticCollector {
    internal static IReadOnlyList<OneNoteMarkdownDiagnostic> Collect(OneNoteSection section, OneNoteMarkdownOptions options, ISet<OneNoteBinaryElement> resolvedAssets) {
        var diagnostics = new List<OneNoteMarkdownDiagnostic>();
        AddSourceDiagnostics(diagnostics, section.Diagnostics, SourceName(section.Name, "OneNote section"));
        AddOpaqueDiagnostic(diagnostics, section.UnknownObjects.Count, SourceName(section.Name, "OneNote section"));
        foreach (OneNotePage page in section.Pages) InspectPage(diagnostics, page, options, resolvedAssets);
        return diagnostics.AsReadOnly();
    }

    internal static IReadOnlyList<OneNoteMarkdownDiagnostic> Collect(OneNoteNotebook notebook, OneNoteMarkdownOptions options, ISet<OneNoteBinaryElement> resolvedAssets) {
        var diagnostics = new List<OneNoteMarkdownDiagnostic>();
        AddSourceDiagnostics(diagnostics, notebook.Diagnostics, SourceName(notebook.Name, "OneNote notebook"));
        AddOpaqueDiagnostic(diagnostics, notebook.UnknownObjects.Count, SourceName(notebook.Name, "OneNote notebook"));
        foreach (OneNoteSection section in notebook.Sections) InspectSection(diagnostics, section, options, resolvedAssets);
        foreach (OneNoteSectionGroup group in notebook.SectionGroups) InspectGroup(diagnostics, group, options, resolvedAssets);
        return diagnostics.AsReadOnly();
    }

    private static void InspectGroup(List<OneNoteMarkdownDiagnostic> diagnostics, OneNoteSectionGroup group, OneNoteMarkdownOptions options, ISet<OneNoteBinaryElement> resolvedAssets) {
        string source = SourceName(group.Name, "Section group");
        AddOpaqueDiagnostic(diagnostics, group.UnknownObjects.Count, source);
        foreach (OneNoteSection section in group.Sections) InspectSection(diagnostics, section, options, resolvedAssets);
        foreach (OneNoteSectionGroup child in group.SectionGroups) InspectGroup(diagnostics, child, options, resolvedAssets);
    }

    private static void InspectSection(List<OneNoteMarkdownDiagnostic> diagnostics, OneNoteSection section, OneNoteMarkdownOptions options, ISet<OneNoteBinaryElement> resolvedAssets) {
        string source = SourceName(section.Name, "OneNote section");
        AddSourceDiagnostics(diagnostics, section.Diagnostics, source);
        AddOpaqueDiagnostic(diagnostics, section.UnknownObjects.Count, source);
        foreach (OneNotePage page in section.Pages) InspectPage(diagnostics, page, options, resolvedAssets);
    }

    private static void InspectPage(List<OneNoteMarkdownDiagnostic> diagnostics, OneNotePage page, OneNoteMarkdownOptions options, ISet<OneNoteBinaryElement> resolvedAssets) {
        string source = SourceName(page.Title, "Untitled page");
        AddSourceDiagnostics(diagnostics, page.Diagnostics, source);
        AddOpaqueDiagnostic(diagnostics, page.UnknownObjects.Count, source);

        var state = new PageProjectionState();
        foreach (OneNoteOutline outline in page.Outlines) InspectElement(outline, options, resolvedAssets, state);
        foreach (OneNoteElement element in page.DirectContent) InspectElement(element, options, resolvedAssets, state);
        if (page.Width.HasValue || page.Height.HasValue) state.HasPositionedLayout = true;

        if (state.HasPositionedLayout) {
            diagnostics.Add(new OneNoteMarkdownDiagnostic(
                "ONENOTE_MARKDOWN_CANVAS_FLATTENED",
                OneNoteDiagnosticSeverity.Warning,
                source,
                "Free-form OneNote canvas placement was flattened into semantic document order."));
        }
        if (state.PlaceholderAssetCount > 0) {
            diagnostics.Add(new OneNoteMarkdownDiagnostic(
                "ONENOTE_MARKDOWN_ASSET_PLACEHOLDER",
                OneNoteDiagnosticSeverity.Warning,
                source,
                state.PlaceholderAssetCount + " image or binary asset(s) were represented by readable placeholders because no payload URI was resolved."));
        }
        if (state.LinkedBinaryCount > 0) {
            diagnostics.Add(new OneNoteMarkdownDiagnostic(
                "ONENOTE_MARKDOWN_BINARY_LINK_ONLY",
                OneNoteDiagnosticSeverity.Warning,
                source,
                state.LinkedBinaryCount + " attachment, recording, or ink payload(s) were projected as links rather than embedded page content."));
        }
        if (state.SimplifiedFormattingCount > 0) {
            diagnostics.Add(new OneNoteMarkdownDiagnostic(
                "ONENOTE_MARKDOWN_FORMATTING_SIMPLIFIED",
                OneNoteDiagnosticSeverity.Warning,
                source,
                state.SimplifiedFormattingCount + " formatted content item(s) use styling or metadata that Markdown cannot preserve faithfully."));
        }
        AddOpaqueDiagnostic(diagnostics, state.OpaqueItemCount, source);

        if (options.IncludeConflictPages) {
            foreach (OneNotePage conflict in page.ConflictPages) InspectPage(diagnostics, conflict, options, resolvedAssets);
        }
        if (options.IncludeVersionHistory) {
            foreach (OneNotePage version in page.VersionHistory) InspectPage(diagnostics, version, options, resolvedAssets);
        }
    }

    private static void InspectElement(OneNoteElement element, OneNoteMarkdownOptions options, ISet<OneNoteBinaryElement> resolvedAssets, PageProjectionState state) {
        OneNoteLayout? layout = element.Layout;
        if (layout != null && (layout.X.HasValue || layout.Y.HasValue || layout.Width.HasValue || layout.Height.HasValue)) {
            state.HasPositionedLayout = true;
        }
        state.OpaqueItemCount += element.UnknownProperties.Count;
        if (element.Tags.Count > 0 || element.Author != null) state.SimplifiedFormattingCount++;

        if (element is OneNoteOutline outline) {
            foreach (OneNoteElement child in outline.Children) InspectElement(child, options, resolvedAssets, state);
        } else if (element is OneNoteParagraph paragraph) {
            if (paragraph.Style.Alignment.HasValue || paragraph.Style.SpaceBefore.HasValue || paragraph.Style.SpaceAfter.HasValue || paragraph.Style.ExactLineSpacing.HasValue) {
                state.SimplifiedFormattingCount++;
            }
            foreach (OneNoteTextRun run in paragraph.Runs) {
                OneNoteTextStyle style = run.Style;
                if (!string.IsNullOrWhiteSpace(style.FontFamily) || style.FontSize.HasValue || style.ColorArgb.HasValue ||
                    style.HighlightColorArgb.HasValue || style.Underline == true || style.Superscript == true || style.Subscript == true) {
                    state.SimplifiedFormattingCount++;
                }
                state.OpaqueItemCount += run.UnknownProperties.Count;
            }
            foreach (OneNoteElement child in paragraph.Children) InspectElement(child, options, resolvedAssets, state);
        } else if (element is OneNoteTable table) {
            foreach (OneNoteTableRow row in table.Rows) {
                foreach (OneNoteTableCell cell in row.Cells) {
                    if (cell.ShadingColorArgb.HasValue) state.SimplifiedFormattingCount++;
                    state.OpaqueItemCount += cell.UnknownProperties.Count;
                    foreach (OneNoteElement child in cell.Content) InspectElement(child, options, resolvedAssets, state);
                }
            }
        } else if (element is OneNoteBinaryElement binary) {
            bool resolved = binary.Payload != null && resolvedAssets.Contains(binary);
            if (!resolved) state.PlaceholderAssetCount++;
            else if (binary is not OneNoteImage) state.LinkedBinaryCount++;
        }
    }

    private static void AddSourceDiagnostics(List<OneNoteMarkdownDiagnostic> target, IEnumerable<OneNoteDiagnostic> sourceDiagnostics, string fallbackSource) {
        foreach (OneNoteDiagnostic diagnostic in sourceDiagnostics) {
            target.Add(new OneNoteMarkdownDiagnostic(
                string.IsNullOrWhiteSpace(diagnostic.Code) ? "ONENOTE_SOURCE_DIAGNOSTIC" : diagnostic.Code,
                diagnostic.Severity,
                string.IsNullOrWhiteSpace(diagnostic.SourcePath) ? fallbackSource : diagnostic.SourcePath!,
                diagnostic.Message));
        }
    }

    private static void AddOpaqueDiagnostic(List<OneNoteMarkdownDiagnostic> diagnostics, int count, string source) {
        if (count <= 0) return;
        diagnostics.Add(new OneNoteMarkdownDiagnostic(
            "ONENOTE_MARKDOWN_OPAQUE_CONTENT_OMITTED",
            OneNoteDiagnosticSeverity.Warning,
            source,
            count + " opaque or unknown source item(s) cannot be represented by semantic Markdown."));
    }

    private static string SourceName(string? value, string fallback) => string.IsNullOrWhiteSpace(value) ? fallback : value!;

    private sealed class PageProjectionState {
        internal bool HasPositionedLayout { get; set; }
        internal int PlaceholderAssetCount { get; set; }
        internal int LinkedBinaryCount { get; set; }
        internal int SimplifiedFormattingCount { get; set; }
        internal int OpaqueItemCount { get; set; }
    }
}
