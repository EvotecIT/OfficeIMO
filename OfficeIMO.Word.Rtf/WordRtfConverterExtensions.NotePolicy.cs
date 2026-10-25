using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Word.Rtf;

public static partial class WordRtfConverterExtensions {
    private static RtfNoteProjection? ResolveNoteProjection(RtfDocument? source, WordParagraph reference) {
        if (source == null) return null;
        WordSection? owner = reference.GetBodySection();
        for (int index = 0; index < reference._document.Sections.Count && index < source.Sections.Count; index++) {
            if (ReferenceEquals(reference._document.Sections[index]._sectionProperties, owner?._sectionProperties))
                return RtfNoteProjection.Resolve(source, source.Sections[index]);
        }
        return RtfNoteProjection.Resolve(source);
    }

    private static void AddRtfNotePolicyDiagnostics(RtfDocument source, RtfConversionReport report) {
        IEnumerable<RtfSection?> scopes = source.Sections.Cast<RtfSection?>();
        if (source.Sections.Count == 0) scopes = new RtfSection?[] { null };
        var positions = new HashSet<RtfEndnotePlacement>();
        foreach (RtfSection? section in scopes) {
            HashSet<RtfNote> notes = section == null ? RtfNoteReferenceCollector.Collect(source)
                : RtfNoteReferenceCollector.Collect(section.Blocks);
            RtfNoteProjection projection = RtfNoteProjection.Resolve(source, section);
            if (projection.FootnotesAreEndnotes && notes.Any(note => note.Kind == RtfNoteKind.Footnote) &&
                notes.Any(note => note.Kind == RtfNoteKind.Endnote)) {
                report.Add(RtfConversionSeverity.Warning, "RtfWordNoteGroupsFlattened",
                    "Legacy ordinary notes and explicit endnotes share one Word endnote group. The legacy ordinary-note settings are retained; independent group placement and numbering are not preserved.",
                    RtfConversionAction.Flattened, feature: "NoteGroups");
            }
            if (notes.Any(note => projection.ProjectKind(note.Kind) == RtfNoteKind.Endnote)) {
                positions.Add(projection.Settings.EndnotePlacement ?? RtfEndnotePlacement.DocumentEnd);
                if (projection.Settings.EndnoteRestart == RtfNoteNumberRestart.EachPage) {
                    report.Add(RtfConversionSeverity.Warning, "RtfWordEndnoteRestartFlattened",
                        "Word endnotes do not support restarting at each page; the page restart declaration is omitted.",
                        RtfConversionAction.Omitted, feature: "EndnoteRestart");
                }
            }
        }
        if (positions.Count > 1) report.Add(RtfConversionSeverity.Warning, "RtfWordEndnotePlacementFlattened",
            "Independent section endnote placements cannot be represented by Word's document-wide endnote placement. The last applied section placement is retained.",
            RtfConversionAction.Flattened, feature: "EndnotePlacement", count: positions.Count);
    }
}
