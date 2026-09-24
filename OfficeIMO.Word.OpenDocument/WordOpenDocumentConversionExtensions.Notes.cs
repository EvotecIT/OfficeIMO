using OfficeIMO.OpenDocument;
using OfficeIMO.Word;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.OpenDocument;

public static partial class WordOpenDocumentConversionExtensions {
    private sealed class NoteMappingStats {
        internal int SeenWordFootnotes;
        internal int SeenWordEndnotes;
        internal int SeenOdtNotes;
        internal int ConvertedFootnotes;
        internal int ConvertedEndnotes;
        internal int UnsupportedFootnotes;
        internal int UnsupportedEndnotes;
        internal int ApproximatedBodies;
        internal int ApproximatedCitations;
        internal int UnsupportedBodyContent;
        internal int ApproximatedReferencePositions;
        internal int ApproximatedNumberingAndPlacement;
    }

    private static void CopyWordNotes(WordRunSnapshot run, OdtParagraph target, NoteMappingStats notes) {
        if (run.Footnote != null) {
            notes.SeenWordFootnotes++;
            CopyWordNote(run.Footnote.Paragraphs, OdtNoteKind.Footnote, target, notes);
            if (run.Text.Length > 0) notes.ApproximatedReferencePositions++;
        }
        if (run.Endnote != null) {
            notes.SeenWordEndnotes++;
            CopyWordNote(run.Endnote.Paragraphs, OdtNoteKind.Endnote, target, notes);
            if (run.Text.Length > 0) notes.ApproximatedReferencePositions++;
        }
    }

    private static void CopyWordNote(IReadOnlyList<WordParagraphSnapshot> paragraphs, OdtNoteKind kind,
        OdtParagraph target, NoteMappingStats notes) {
        if (paragraphs.Count == 0) {
            CountUnsupported(kind, notes);
            return;
        }
        OdtNote result = kind == OdtNoteKind.Footnote
            ? target.AddFootnote(paragraphs[0].Text)
            : target.AddEndnote(paragraphs[0].Text);
        for (int index = 1; index < paragraphs.Count; index++) result.AddParagraph(paragraphs[index].Text);
        if (paragraphs.Any(paragraph => HasNonPlainWordNoteContent(paragraph, kind))) notes.ApproximatedBodies++;
        if (paragraphs.Any(paragraph => paragraph.Runs.Any(run =>
            run.InlineImage != null || run.Footnote != null || run.Endnote != null))) notes.UnsupportedBodyContent++;
        if (kind == OdtNoteKind.Footnote) notes.ConvertedFootnotes++;
        else notes.ConvertedEndnotes++;
    }

    private static bool HasNonPlainWordNoteContent(WordParagraphSnapshot paragraph, OdtNoteKind kind) =>
        HasNonDefaultWordNoteStyle(paragraph, kind) ||
        paragraph.Runs.Any(run => run.Text.Length > 0 && (
            run.Bold || run.Italic || run.Underline || run.Strike || run.DoubleStrike ||
            run.IsHyperlink ||
            run.FontSizePoints.HasValue || run.FontFamily != null || run.ColorHex != null ||
            run.HighlightColor != null || run.RunShadingFillColorHex != null)) ||
        paragraph.Alignment != null || paragraph.IndentStartPoints.HasValue ||
        paragraph.IndentEndPoints.HasValue || paragraph.IndentFirstLinePoints.HasValue ||
        paragraph.ShadingFillColorHex != null || paragraph.BookmarkName != null;

    private static bool HasNonDefaultWordNoteStyle(WordParagraphSnapshot paragraph, OdtNoteKind kind) {
        string defaultId = kind == OdtNoteKind.Footnote ? "FootnoteText" : "EndnoteText";
        if (!string.IsNullOrWhiteSpace(paragraph.StyleId))
            return !string.Equals(paragraph.StyleId, defaultId, StringComparison.OrdinalIgnoreCase);
        if (string.IsNullOrWhiteSpace(paragraph.StyleName)) return false;
        string defaultName = kind == OdtNoteKind.Footnote ? "Footnote Text" : "Endnote Text";
        return !string.Equals(paragraph.StyleName, defaultName, StringComparison.OrdinalIgnoreCase) &&
            !string.Equals(paragraph.StyleName, defaultId, StringComparison.OrdinalIgnoreCase);
    }

    private static void CountWordNoteSettingsLoss(WordDocument source, NoteMappingStats notes) {
        int affectedNotes = 0;
        if (notes.ConvertedFootnotes > 0 && (source.Sections.Any(section => {
            WordFootnoteSettings settings = section.FootnoteSettings;
            return settings.Position.HasValue || settings.NumberingRestart.HasValue ||
                settings.StartNumber.HasValue || settings.NumberingFormat.HasValue;
        }) || HasDocumentNoteSettings<W.FootnoteProperties>(source))) affectedNotes += notes.ConvertedFootnotes;
        if (notes.ConvertedEndnotes > 0 && (source.Sections.Any(section => {
            WordEndnoteSettings settings = section.EndnoteSettings;
            return settings.Position.HasValue || settings.NumberingRestart.HasValue ||
                settings.StartNumber.HasValue || settings.NumberingFormat.HasValue;
        }) || HasDocumentNoteSettings<W.EndnoteProperties>(source))) affectedNotes += notes.ConvertedEndnotes;
        notes.ApproximatedNumberingAndPlacement = affectedNotes;
        if (notes.ConvertedFootnotes > 0 && HasCustomizedDefaultWordNoteStyle(source, "FootnoteText"))
            notes.ApproximatedBodies += notes.ConvertedFootnotes;
        if (notes.ConvertedEndnotes > 0 && HasCustomizedDefaultWordNoteStyle(source, "EndnoteText"))
            notes.ApproximatedBodies += notes.ConvertedEndnotes;
    }

    private static bool HasCustomizedDefaultWordNoteStyle(WordDocument source, string styleId) {
        W.Style? style = source.OpenXmlDocument.MainDocumentPart?.StyleDefinitionsPart?.Styles?
            .Elements<W.Style>().FirstOrDefault(candidate => string.Equals(candidate.StyleId?.Value,
                styleId, StringComparison.OrdinalIgnoreCase));
        if (style == null) return false;
        if (style.CustomStyle?.Value == true || style.BasedOn?.Val?.Value != "Normal") return true;
        W.StyleParagraphProperties? paragraph = style.StyleParagraphProperties;
        if (paragraph?.ChildElements.Count != 1 || paragraph.FirstChild is not W.SpacingBetweenLines spacing ||
            spacing.After?.Value != "0" || spacing.Line?.Value != "240" ||
            spacing.LineRule?.Value != W.LineSpacingRuleValues.Auto) return true;
        W.StyleRunProperties? run = style.StyleRunProperties;
        return run?.ChildElements.Count != 2 ||
            run.GetFirstChild<W.FontSize>()?.Val?.Value != "20" ||
            run.GetFirstChild<W.FontSizeComplexScript>()?.Val?.Value != "20";
    }

    private static bool HasDocumentNoteSettings<T>(WordDocument source) where T : DocumentFormat.OpenXml.OpenXmlElement =>
        source.OpenXmlDocument.MainDocumentPart?.DocumentSettingsPart?.Settings?
            .GetFirstChild<T>()?.ChildElements.Any(child => child is W.NumberingFormat or W.NumberingStart or
                W.NumberingRestart or W.FootnotePosition or W.EndnotePosition) == true;

    private static void CopyOdtNote(OdtNote source, WordParagraph target, NoteMappingStats notes) {
        notes.SeenOdtNotes++;
        if (!source.Kind.HasValue || !source.HasOnlyParagraphs || source.Paragraphs.Count == 0) {
            if (source.Kind == OdtNoteKind.Endnote) notes.UnsupportedEndnotes++;
            else notes.UnsupportedFootnotes++;
            return;
        }
        IReadOnlyList<OdtParagraph> paragraphs = source.Paragraphs;
        if (paragraphs.Any(paragraph => paragraph.InlineNodes.Any(node => node.Kind == OdtInlineNodeKind.Note))) {
            CountUnsupported(source.Kind.Value, notes);
            return;
        }
        if (paragraphs.Any(paragraph => paragraph.InlineNodes.Any(node =>
            node.Kind == OdtInlineNodeKind.Image || node.Kind == OdtInlineNodeKind.Other))) notes.UnsupportedBodyContent++;
        string text = string.Join("\n", paragraphs.Select(paragraph => paragraph.Text));
        if (source.Kind == OdtNoteKind.Footnote) {
            target.AddFootNote(text);
            notes.ConvertedFootnotes++;
            if (source.Citation != notes.ConvertedFootnotes.ToString(System.Globalization.CultureInfo.InvariantCulture))
                notes.ApproximatedCitations++;
        } else {
            target.AddEndNote(text);
            notes.ConvertedEndnotes++;
            if (source.Citation != notes.ConvertedEndnotes.ToString(System.Globalization.CultureInfo.InvariantCulture))
                notes.ApproximatedCitations++;
        }
        if (paragraphs.Count != 1 || paragraphs.Any(paragraph => paragraph.StyleName != null ||
            paragraph.InlineNodes.Any(node => node.Kind != OdtInlineNodeKind.Text))) notes.ApproximatedBodies++;
    }

    private static void CountUnsupported(OdtNoteKind kind, NoteMappingStats notes) {
        if (kind == OdtNoteKind.Footnote) notes.UnsupportedFootnotes++;
        else notes.UnsupportedEndnotes++;
    }

    private static void AddNoteMappings(OdfConversionReport report, NoteMappingStats notes) {
        AddCount(report, "footnotes", notes.ConvertedFootnotes);
        AddCount(report, "endnotes", notes.ConvertedEndnotes);
        if (notes.UnsupportedFootnotes > 0) report.Add("footnotes", OdfConversionMappingStatus.Unsupported,
            notes.UnsupportedFootnotes, "Note references without a supported body or class were omitted.");
        if (notes.UnsupportedEndnotes > 0) report.Add("endnotes", OdfConversionMappingStatus.Unsupported,
            notes.UnsupportedEndnotes, "Note references without a supported body were omitted.");
        if (notes.ApproximatedBodies > 0) report.Add("note-body-formatting", OdfConversionMappingStatus.Approximated,
            notes.ApproximatedBodies, "Note text was retained, but some note formatting or paragraph structure was flattened.");
        if (notes.UnsupportedBodyContent > 0) report.Add("note-body-content", OdfConversionMappingStatus.Unsupported,
            notes.UnsupportedBodyContent, "Note text was retained, but embedded media or unmodeled inline content were omitted.");
        if (notes.ApproximatedReferencePositions > 0) report.Add("note-reference-position", OdfConversionMappingStatus.Approximated,
            notes.ApproximatedReferencePositions, "A Word run contained both text and a note reference, so the reference was placed after its text.");
        if (notes.ApproximatedCitations > 0) report.Add("note-citations", OdfConversionMappingStatus.Approximated,
            notes.ApproximatedCitations, "Custom ODT citation text was replaced with Word's automatic note numbering.");
        if (notes.ApproximatedNumberingAndPlacement > 0) report.Add("note-numbering-placement", OdfConversionMappingStatus.Approximated,
            notes.ApproximatedNumberingAndPlacement,
            "Word note numbering format, start, restart, or placement settings were not carried into ODT.");
    }
}
