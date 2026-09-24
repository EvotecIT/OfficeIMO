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
        internal int UnsupportedHeaderFooterNotes;
        internal int UnsupportedOdtNoteConfigurations;
        internal int CurrentSectionIndex;
        internal WordDocument? WordSource;
        internal readonly Dictionary<int, int> FootnotesBySection = new Dictionary<int, int>();
        internal readonly Dictionary<int, int> EndnotesBySection = new Dictionary<int, int>();
    }

    private static void CopyWordNotes(WordRunSnapshot run, OdtParagraph target, NoteMappingStats notes) {
        if (run.Footnote != null) {
            notes.SeenWordFootnotes++;
            CopyWordNote(run.Footnote.Paragraphs, run.Footnote.ReferenceId, OdtNoteKind.Footnote, target, notes);
            if (HasAmbiguousWordNotePosition(run)) notes.ApproximatedReferencePositions++;
        }
        if (run.Endnote != null) {
            notes.SeenWordEndnotes++;
            CopyWordNote(run.Endnote.Paragraphs, run.Endnote.ReferenceId, OdtNoteKind.Endnote, target, notes);
            if (HasAmbiguousWordNotePosition(run)) notes.ApproximatedReferencePositions++;
        }
    }

    private static bool HasAmbiguousWordNotePosition(WordRunSnapshot run) =>
        run.Text.Length > 0 || run.PositionedImages.Count > 0 || run.NonTextBreaks?.Count > 0 ||
        (run.Footnote != null && run.Endnote != null);

    private static void CopyWordNote(IReadOnlyList<WordParagraphSnapshot> paragraphs, long? referenceId, OdtNoteKind kind,
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
        bool unsupportedInline = paragraphs.Any(paragraph => paragraph.Runs.Any(run =>
            run.InlineImage != null || run.Footnote != null || run.Endnote != null));
        bool unsupportedBlock = notes.WordSource != null &&
            HasUnsupportedWordNoteBlocks(notes.WordSource, referenceId, kind);
        if (unsupportedInline || unsupportedBlock) notes.UnsupportedBodyContent++;
        if (kind == OdtNoteKind.Footnote) {
            notes.ConvertedFootnotes++;
            Increment(notes.FootnotesBySection, notes.CurrentSectionIndex);
        } else {
            notes.ConvertedEndnotes++;
            Increment(notes.EndnotesBySection, notes.CurrentSectionIndex);
        }
    }

    private static void Increment(Dictionary<int, int> counts, int key) =>
        counts[key] = counts.TryGetValue(key, out int count) ? count + 1 : 1;

    private static bool HasUnsupportedWordNoteBlocks(WordDocument source, long? referenceId, OdtNoteKind kind) {
        if (!referenceId.HasValue) return false;
        DocumentFormat.OpenXml.OpenXmlElement? note = kind == OdtNoteKind.Footnote
            ? source.OpenXmlDocument.MainDocumentPart?.FootnotesPart?.Footnotes?
                .Elements<W.Footnote>().FirstOrDefault(item => item.Id?.Value == referenceId.Value)
            : source.OpenXmlDocument.MainDocumentPart?.EndnotesPart?.Endnotes?
                .Elements<W.Endnote>().FirstOrDefault(item => item.Id?.Value == referenceId.Value);
        return note != null && note.ChildElements.Any(child => child is not W.Paragraph);
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
        paragraph.SpaceAbovePoints.HasValue || paragraph.SpaceBelowPoints.HasValue ||
        paragraph.LineSpacingValue.HasValue || paragraph.LineSpacingRule != null ||
        paragraph.ShadingFillColorHex != null || paragraph.ShadingPattern.HasValue ||
        paragraph.LeftBorder != null || paragraph.RightBorder != null ||
        paragraph.TopBorder != null || paragraph.BottomBorder != null ||
        paragraph.TabStops.Count > 0 || paragraph.IsListItem || paragraph.IsRightToLeft ||
        paragraph.KeepWithNext || paragraph.KeepLinesTogether || paragraph.AvoidWidowAndOrphan ||
        paragraph.PageBreakBefore || paragraph.BookmarkName != null;

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
        bool documentFootnoteSettings = HasDocumentNoteSettings<W.FootnoteProperties>(source);
        bool documentEndnoteSettings = HasDocumentNoteSettings<W.EndnoteProperties>(source);
        foreach (KeyValuePair<int, int> entry in notes.FootnotesBySection) {
            if (documentFootnoteSettings || HasSectionFootnoteSettings(source.Sections[entry.Key].FootnoteSettings))
                affectedNotes += entry.Value;
        }
        foreach (KeyValuePair<int, int> entry in notes.EndnotesBySection) {
            if (documentEndnoteSettings || HasSectionEndnoteSettings(source.Sections[entry.Key].EndnoteSettings))
                affectedNotes += entry.Value;
        }
        notes.ApproximatedNumberingAndPlacement = affectedNotes;
        if (notes.ConvertedFootnotes > 0 && HasCustomizedDefaultWordNoteStyle(source, "FootnoteText"))
            notes.ApproximatedBodies += notes.ConvertedFootnotes;
        if (notes.ConvertedEndnotes > 0 && HasCustomizedDefaultWordNoteStyle(source, "EndnoteText"))
            notes.ApproximatedBodies += notes.ConvertedEndnotes;
    }

    private static bool HasSectionFootnoteSettings(WordFootnoteSettings settings) =>
        settings.Position.HasValue || settings.NumberingRestart.HasValue ||
        settings.StartNumber.HasValue || settings.NumberingFormat.HasValue;

    private static bool HasSectionEndnoteSettings(WordEndnoteSettings settings) =>
        settings.Position.HasValue || settings.NumberingRestart.HasValue ||
        settings.StartNumber.HasValue || settings.NumberingFormat.HasValue;

    private static void CountOdtNoteConfigurationLoss(OdtDocument source, NoteMappingStats notes) {
        if (!source.Package.ContainsEntry("styles.xml")) return;
        System.Xml.Linq.XNamespace text = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
        notes.UnsupportedOdtNoteConfigurations = source.Package.GetXml("styles.xml")
            .Descendants(text + "notes-configuration").Count();
    }

    private static bool HasCustomizedDefaultWordNoteStyle(WordDocument source, string styleId) {
        W.Styles? styles = source.OpenXmlDocument.MainDocumentPart?.StyleDefinitionsPart?.Styles;
        W.Style? style = styles?
            .Elements<W.Style>().FirstOrDefault(candidate => string.Equals(candidate.StyleId?.Value,
                styleId, StringComparison.OrdinalIgnoreCase));
        if (style == null) return false;
        W.Style? normal = styles?.Elements<W.Style>().FirstOrDefault(candidate =>
            string.Equals(candidate.StyleId?.Value, "Normal", StringComparison.OrdinalIgnoreCase));
        if (normal?.StyleParagraphProperties?.ChildElements.Count > 0 ||
            normal?.StyleRunProperties?.ChildElements.Count > 0) return true;
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

    private static void CountUnsupportedHeaderFooterNote(NoteMappingStats notes) {
        notes.SeenOdtNotes++;
        notes.UnsupportedHeaderFooterNotes++;
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
            notes.UnsupportedBodyContent, "Note text was retained, but embedded media or unsupported block or inline content was omitted.");
        if (notes.UnsupportedHeaderFooterNotes > 0) report.Add("note-headers-footers", OdfConversionMappingStatus.Unsupported,
            notes.UnsupportedHeaderFooterNotes, "ODT header and footer note references were omitted because Word does not permit them there.");
        if (notes.UnsupportedOdtNoteConfigurations > 0) report.Add("note-configuration", OdfConversionMappingStatus.Unsupported,
            notes.UnsupportedOdtNoteConfigurations, "ODT note numbering and placement configuration was not carried into Word.");
        if (notes.ApproximatedReferencePositions > 0) report.Add("note-reference-position", OdfConversionMappingStatus.Approximated,
            notes.ApproximatedReferencePositions,
            "A Word run contained a note reference with text, another note, a break, or an image; relative inline order may have changed.");
        if (notes.ApproximatedCitations > 0) report.Add("note-citations", OdfConversionMappingStatus.Approximated,
            notes.ApproximatedCitations, "Custom ODT citation text was replaced with Word's automatic note numbering.");
        if (notes.ApproximatedNumberingAndPlacement > 0) report.Add("note-numbering-placement", OdfConversionMappingStatus.Approximated,
            notes.ApproximatedNumberingAndPlacement,
            "Word note numbering format, start, restart, or placement settings were not carried into ODT.");
    }
}
