using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Legacy;

public static partial class LegacyWordImporter {
    private static IReadOnlyList<LegacyWordSection> Sections(LegacyWordModel model) {
        if (model.Sections.Count > 0) return model.Sections;
        var section = new LegacyWordSection();
        section.Blocks.AddRange(model.Paragraphs);
        return new[] { section };
    }

    private static void ProjectSections(LegacyWordModel model, WordDocument document,
        IDictionary<string, string> styleIds, ISet<string> usedStyleIds,
        IReadOnlyDictionary<string, LegacyWordStyle> recoveredStyles, CancellationToken cancellation) {
        IReadOnlyList<LegacyWordSection> sections = Sections(model);
        bool hasStories = sections.Any(section => section.HeadersAndFooters.Count > 0);
        bool oddEven = sections.Any(section => section.HeadersAndFooters.Any(story => story.Occurrence != LegacyWordHeaderFooterOccurrence.All));
        for (int index = 0; index < sections.Count; index++) {
            cancellation.ThrowIfCancellationRequested();
            LegacyWordSection source = sections[index];
            WordSection target = index == 0 ? document.Sections[0] :
                document.AddSection(source.StartsNewPage ? WordSectionBreakType.NextPage : WordSectionBreakType.Continuous);
            if (source.WidthPoints.HasValue) target.PageSettings.Width = checked((uint)Twips(source.WidthPoints.Value));
            if (source.HeightPoints.HasValue) target.PageSettings.Height = checked((uint)Twips(source.HeightPoints.Value));
            if (source.LeftPoints.HasValue) target.Margins.Left = checked((uint)Twips(source.LeftPoints.Value));
            if (source.RightPoints.HasValue) target.Margins.Right = checked((uint)Twips(source.RightPoints.Value));
            if (source.TopPoints.HasValue) target.Margins.Top = Twips(source.TopPoints.Value);
            if (source.BottomPoints.HasValue) target.Margins.Bottom = Twips(source.BottomPoints.Value);
            WordList? list = null;
            foreach (LegacyWordBlock block in source.Blocks) {
                cancellation.ThrowIfCancellationRequested();
                if (block is LegacyWordParagraph paragraph) {
                    WordParagraph projected;
                    if (paragraph.IsList) {
                        list ??= document.AddListBulleted();
                        projected = list.AddItem(string.Empty, Math.Max(0, Math.Min(8, paragraph.ListLevel)));
                    } else { list = null; projected = target.AddParagraph(); }
                    ProjectParagraph(paragraph, projected, document, styleIds, usedStyleIds, recoveredStyles, cancellation, model);
                } else if (block is LegacyWordTable table) {
                    list = null;
                    ProjectTable(table, target, model, document, styleIds, usedStyleIds, recoveredStyles, cancellation);
                }
            }
            if (source.Blocks.Count == 0) target.AddParagraph();
            if (hasStories) ProjectStories(source, target, oddEven, model, document, styleIds, usedStyleIds, recoveredStyles, cancellation);
        }
    }

    private static void ProjectTable(LegacyWordTable source, WordSection section, LegacyWordModel model, WordDocument document,
        IDictionary<string, string> styleIds, ISet<string> usedStyleIds,
        IReadOnlyDictionary<string, LegacyWordStyle> recoveredStyles, CancellationToken cancellation) {
        int columns = source.ColumnWidthsPoints.Count > 0 ? source.ColumnWidthsPoints.Count : source.Rows.Max(row => row.Cells.Count);
        WordTable table = section.AddTable(source.Rows.Count, columns, WordTableStyle.TableNormal);
        table.LayoutMode = WordTableLayoutMode.Fixed;
        for (int row = 0; row < source.Rows.Count; row++) {
            cancellation.ThrowIfCancellationRequested();
            table.Rows[row].RepeatHeaderRowAtTheTopOfEachPage = source.Rows[row].IsHeader;
            for (int column = 0; column < source.Rows[row].Cells.Count; column++) {
                LegacyWordTableCell sourceCell = source.Rows[row].Cells[column];
                WordTableCell cell = table.Rows[row].Cells[column];
                if (column < source.ColumnWidthsPoints.Count) {
                    cell.WidthType = WordTableWidthUnit.Dxa;
                    cell.Width = Twips(source.ColumnWidthsPoints[column]);
                }
                for (int paragraph = 0; paragraph < sourceCell.Paragraphs.Count; paragraph++) {
                    cancellation.ThrowIfCancellationRequested();
                    WordParagraph target = paragraph == 0 ? cell.Paragraphs[0] : cell.AddParagraph();
                    ProjectParagraph(sourceCell.Paragraphs[paragraph], target, document, styleIds, usedStyleIds, recoveredStyles, cancellation, model);
                }
            }
        }
    }

    private static void ProjectStories(LegacyWordSection source, WordSection target, bool oddEven, LegacyWordModel model,
        WordDocument document, IDictionary<string, string> styleIds, ISet<string> usedStyleIds,
        IReadOnlyDictionary<string, LegacyWordStyle> recoveredStyles, CancellationToken cancellation) {
        // Every section receives its own parts, including blanks for discontinued stories. This prevents
        // DOCX inheritance from keeping a cancelled header or sharing a changed header with an earlier section.
        target._sectionProperties.RemoveAllChildren<HeaderReference>();
        target._sectionProperties.RemoveAllChildren<FooterReference>();
        if (oddEven) target.DifferentOddAndEvenPages = true;
        foreach (bool footer in new[] { false, true }) {
            foreach (WordHeaderFooterType type in oddEven ? new[] { WordHeaderFooterType.Default, WordHeaderFooterType.Even } : new[] { WordHeaderFooterType.Default }) {
                cancellation.ThrowIfCancellationRequested();
                WordHeaderFooter part = footer ? target.GetOrCreateFooter(type) : target.GetOrCreateHeader(type);
                foreach (WordParagraph old in part.Paragraphs.ToList()) old.Remove();
                foreach (LegacyWordHeaderFooter story in source.HeadersAndFooters.Where(story => story.IsFooter == footer &&
                    (story.Occurrence == LegacyWordHeaderFooterOccurrence.All ||
                        story.Occurrence == (type == WordHeaderFooterType.Even ? LegacyWordHeaderFooterOccurrence.Even : LegacyWordHeaderFooterOccurrence.Odd)))) {
                    foreach (LegacyWordParagraph paragraph in story.Paragraphs)
                        ProjectParagraph(paragraph, part.AddParagraph(), document, styleIds, usedStyleIds, recoveredStyles, cancellation, model);
                }
                if (part.Paragraphs.Count == 0) part.AddParagraph();
            }
        }
    }

    private static void ProjectNote(LegacyWordModel model, int index, WordParagraph anchor, WordDocument document,
        IDictionary<string, string> styleIds, ISet<string> usedStyleIds,
        IReadOnlyDictionary<string, LegacyWordStyle> recoveredStyles, CancellationToken cancellation) {
        if (index < 0 || index >= model.Notes.Count) throw new InvalidDataException("A recovered note anchor is outside the note collection.");
        LegacyWordNote note = model.Notes[index];
        WordParagraph reference = note.Kind == LegacyWordNoteKind.Footnote ? anchor.AddFootNote(string.Empty) : anchor.AddEndNote(string.Empty);
        List<WordParagraph> paragraphs = (note.Kind == LegacyWordNoteKind.Footnote ? reference.FootNote!.Paragraphs : reference.EndNote!.Paragraphs)
            ?? throw new InvalidDataException("The projected note has no paragraphs.");
        WordParagraph target = paragraphs[0];
        for (int i = 0; i < note.Paragraphs.Count; i++) {
            cancellation.ThrowIfCancellationRequested();
            if (i > 0) target = target.AddParagraph();
            ProjectParagraph(note.Paragraphs[i], target, document, styleIds, usedStyleIds, recoveredStyles, cancellation, model);
        }
    }

    private static int Twips(double points) => checked((int)Math.Round(points * 20, MidpointRounding.AwayFromZero));
}
