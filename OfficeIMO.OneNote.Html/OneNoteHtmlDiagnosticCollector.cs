using System.Collections.Generic;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.OneNote.Markdown;

namespace OfficeIMO.OneNote.Html;

internal static class OneNoteHtmlDiagnosticCollector {
    internal static IEnumerable<HtmlDiagnostic> Collect(OneNoteSection section, OneNoteMarkdownOptions options) {
        foreach (OneNotePage page in section.Pages) {
            foreach (HtmlDiagnostic diagnostic in InspectPage(page, options)) yield return diagnostic;
        }
    }

    internal static IEnumerable<HtmlDiagnostic> Collect(OneNoteNotebook notebook, OneNoteMarkdownOptions options) {
        foreach (OneNoteSection section in notebook.Sections) {
            foreach (HtmlDiagnostic diagnostic in Collect(section, options)) yield return diagnostic;
        }
        foreach (OneNoteSectionGroup group in notebook.SectionGroups) {
            foreach (HtmlDiagnostic diagnostic in InspectGroup(group, options, 0)) yield return diagnostic;
        }
    }

    private static IEnumerable<HtmlDiagnostic> InspectGroup(
        OneNoteSectionGroup group,
        OneNoteMarkdownOptions options,
        int depth) {
        if (depth >= options.MaxSectionGroupDepth) yield break;
        foreach (OneNoteSection section in group.Sections) {
            foreach (HtmlDiagnostic diagnostic in Collect(section, options)) yield return diagnostic;
        }
        foreach (OneNoteSectionGroup child in group.SectionGroups) {
            foreach (HtmlDiagnostic diagnostic in InspectGroup(child, options, depth + 1)) yield return diagnostic;
        }
    }

    private static IEnumerable<HtmlDiagnostic> InspectPage(OneNotePage page, OneNoteMarkdownOptions options) {
        int simplified = 0;
        foreach (OneNoteOutline outline in page.Outlines) simplified += CountUnrepresentedFormatting(outline, options, 0);
        foreach (OneNoteElement element in page.DirectContent) simplified += CountUnrepresentedFormatting(element, options, 0);
        if (simplified > 0) {
            yield return new HtmlDiagnostic(
                "OfficeIMO.OneNote.Html",
                "ONENOTE_HTML_FORMATTING_SIMPLIFIED",
                simplified + " formatted content item(s) use spacing or metadata that semantic HTML does not preserve faithfully.",
                HtmlDiagnosticSeverity.Warning,
                string.IsNullOrWhiteSpace(page.Title) ? "Untitled page" : page.Title,
                lossKind: OfficeConversionLossKind.Approximation);
        }

        if (options.IncludeConflictPages) {
            foreach (OneNotePage conflict in page.ConflictPages) {
                foreach (HtmlDiagnostic diagnostic in InspectPage(conflict, options)) yield return diagnostic;
            }
        }
        if (options.IncludeVersionHistory) {
            foreach (OneNotePage version in page.VersionHistory) {
                foreach (HtmlDiagnostic diagnostic in InspectPage(version, options)) yield return diagnostic;
            }
        }
    }

    private static int CountUnrepresentedFormatting(
        OneNoteElement element,
        OneNoteMarkdownOptions options,
        int depth) {
        if (depth >= options.MaxContentDepth) return 0;
        int count = element.Tags.Count > 0 || element.Author != null ? 1 : 0;
        if (element is OneNoteOutline outline) {
            foreach (OneNoteElement child in outline.Children) {
                count += CountUnrepresentedFormatting(child, options, depth + 1);
            }
        } else if (element is OneNoteParagraph paragraph) {
            if (paragraph.Style.SpaceBefore.HasValue || paragraph.Style.SpaceAfter.HasValue ||
                paragraph.Style.ExactLineSpacing.HasValue) count++;
            foreach (OneNoteElement child in paragraph.Children) {
                count += CountUnrepresentedFormatting(child, options, depth + 1);
            }
        } else if (element is OneNoteTable table) {
            foreach (OneNoteTableRow row in table.Rows) {
                foreach (OneNoteTableCell cell in row.Cells) {
                    foreach (OneNoteElement child in cell.Content) {
                        count += CountUnrepresentedFormatting(child, options, depth + 1);
                    }
                }
            }
        }
        return count;
    }
}
