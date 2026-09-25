using OfficeIMO.OpenDocument;
using OfficeIMO.Word;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.OpenDocument;

public static partial class WordOpenDocumentConversionExtensions {
    private sealed class NoteMappingStats {
        internal NoteMappingStats(WordDocument? source = null) {
            WordSource = source;
            W.Body? body = source?.OpenXmlDocument.MainDocumentPart?.Document?.Body;
            if (body != null) {
                foreach (W.Run run in body.Descendants<W.Run>()) {
                    // A run may carry more than one reference of the same kind. The
                    // inspection snapshot currently exposes only its first one.
                    int extraFootnotes = Math.Max(0, run.Elements<W.FootnoteReference>().Count() - 1);
                    int extraEndnotes = Math.Max(0, run.Elements<W.EndnoteReference>().Count() - 1);
                    UnsupportedFootnotes += extraFootnotes;
                    UnsupportedEndnotes += extraEndnotes;
                    SeenWordFootnotes += extraFootnotes;
                    SeenWordEndnotes += extraEndnotes;
                    foreach (W.FootnoteReference reference in run.Elements<W.FootnoteReference>()) {
                        if (reference.Id != null) AddAnchor(FootnoteAnchors, reference.Id.Value, run);
                    }
                    foreach (W.EndnoteReference reference in run.Elements<W.EndnoteReference>()) {
                        if (reference.Id != null) AddAnchor(EndnoteAnchors, reference.Id.Value, run);
                    }
                }
                foreach (List<W.Run> anchors in FootnoteAnchors.Values) {
                    int repeatedAcrossRuns = Math.Max(0, anchors.Distinct().Count() - 1);
                    UnsupportedFootnotes += repeatedAcrossRuns;
                    SeenWordFootnotes += repeatedAcrossRuns;
                }
                foreach (List<W.Run> anchors in EndnoteAnchors.Values) {
                    int repeatedAcrossRuns = Math.Max(0, anchors.Distinct().Count() - 1);
                    UnsupportedEndnotes += repeatedAcrossRuns;
                    SeenWordEndnotes += repeatedAcrossRuns;
                }
            }
            if (source != null) {
                var seenFootnoteDefinitions = new HashSet<long>();
                foreach (W.Footnote note in source.OpenXmlDocument.MainDocumentPart?.FootnotesPart?.Footnotes?
                    .Elements<W.Footnote>() ?? Enumerable.Empty<W.Footnote>()) {
                    if (note.Type?.Value != W.FootnoteEndnoteValues.Separator &&
                        note.Type?.Value != W.FootnoteEndnoteValues.ContinuationSeparator &&
                        (note.Id == null || !FootnoteAnchors.ContainsKey(note.Id.Value) ||
                         !seenFootnoteDefinitions.Add(note.Id.Value)))
                        UnreferencedFootnoteDefinitions++;
                    if (note.Id != null && !FootnoteBodies.ContainsKey(note.Id.Value)) FootnoteBodies.Add(note.Id.Value, note);
                }
                var seenEndnoteDefinitions = new HashSet<long>();
                foreach (W.Endnote note in source.OpenXmlDocument.MainDocumentPart?.EndnotesPart?.Endnotes?
                    .Elements<W.Endnote>() ?? Enumerable.Empty<W.Endnote>()) {
                    if (note.Type?.Value != W.FootnoteEndnoteValues.Separator &&
                        note.Type?.Value != W.FootnoteEndnoteValues.ContinuationSeparator &&
                        (note.Id == null || !EndnoteAnchors.ContainsKey(note.Id.Value) ||
                         !seenEndnoteDefinitions.Add(note.Id.Value)))
                        UnreferencedEndnoteDefinitions++;
                    if (note.Id != null && !EndnoteBodies.ContainsKey(note.Id.Value)) EndnoteBodies.Add(note.Id.Value, note);
                }
            }
        }

        private static void AddAnchor(Dictionary<long, List<W.Run>> index, long id, W.Run run) {
            if (!index.TryGetValue(id, out List<W.Run>? anchors)) index.Add(id, anchors = new List<W.Run>());
            anchors.Add(run);
        }

        internal IReadOnlyList<W.Run> Anchors(OdtNoteKind kind, long? id) {
            if (!id.HasValue) return Array.Empty<W.Run>();
            Dictionary<long, List<W.Run>> index = kind == OdtNoteKind.Footnote ? FootnoteAnchors : EndnoteAnchors;
            return index.TryGetValue(id.Value, out List<W.Run>? runs) ? runs : Array.Empty<W.Run>();
        }

        internal DocumentFormat.OpenXml.OpenXmlElement? Body(OdtNoteKind kind, long? id) {
            if (!id.HasValue) return null;
            Dictionary<long, DocumentFormat.OpenXml.OpenXmlElement> index = kind == OdtNoteKind.Footnote
                ? FootnoteBodies : EndnoteBodies;
            return index.TryGetValue(id.Value, out DocumentFormat.OpenXml.OpenXmlElement? note) ? note : null;
        }

        private readonly Dictionary<long, List<W.Run>> FootnoteAnchors = new Dictionary<long, List<W.Run>>();
        private readonly Dictionary<long, List<W.Run>> EndnoteAnchors = new Dictionary<long, List<W.Run>>();
        private readonly Dictionary<long, DocumentFormat.OpenXml.OpenXmlElement> FootnoteBodies =
            new Dictionary<long, DocumentFormat.OpenXml.OpenXmlElement>();
        private readonly Dictionary<long, DocumentFormat.OpenXml.OpenXmlElement> EndnoteBodies =
            new Dictionary<long, DocumentFormat.OpenXml.OpenXmlElement>();
        internal int SeenWordFootnotes;
        internal int SeenWordEndnotes;
        internal int UnreferencedFootnoteDefinitions;
        internal int UnreferencedEndnoteDefinitions;
        internal int SeenOdtNotes;
        internal int ConvertedFootnotes;
        internal int ConvertedEndnotes;
        internal int UnsupportedFootnotes;
        internal int UnsupportedEndnotes;
        internal int ApproximatedBodies;
        internal int ApproximatedCitations;
        internal int UnsupportedBodyContent;
        internal int ApproximatedReferencePositions;
        internal int ApproximatedReferenceFormatting;
        internal int ApproximatedNumberingAndPlacement;
        internal int ApproximatedSeparators;
        internal int UnsupportedHeaderFooterNotes;
        internal int UnsupportedOdtNoteConfigurations;
        internal bool HasOdtDefaultNoteBodyFormatting;
        internal bool HasOdtDefaultNoteReferenceFormatting;
        internal int CurrentSectionIndex;
        internal WordDocument? WordSource;
        internal readonly Dictionary<int, int> FootnotesBySection = new Dictionary<int, int>();
        internal readonly Dictionary<int, int> EndnotesBySection = new Dictionary<int, int>();
        internal readonly HashSet<long> ProcessedFootnoteIds = new HashSet<long>();
        internal readonly HashSet<long> ProcessedEndnoteIds = new HashSet<long>();
    }

    private static void CopyWordNotes(WordRunSnapshot run, OdtParagraph target, NoteMappingStats notes) {
        if (run.Footnote != null) {
            notes.SeenWordFootnotes++;
            CopyWordNote(run.Footnote.Paragraphs, run.Footnote.ReferenceId, OdtNoteKind.Footnote, target, notes);
            if (HasAmbiguousWordNotePosition(run)) notes.ApproximatedReferencePositions++;
            if (HasCustomWordNoteMark(notes, run.Footnote.ReferenceId, OdtNoteKind.Footnote))
                notes.ApproximatedCitations++;
            if (HasStyledWordNoteAnchor(notes, run.Footnote.ReferenceId, OdtNoteKind.Footnote))
                notes.ApproximatedReferenceFormatting++;
        }
        if (run.Endnote != null) {
            notes.SeenWordEndnotes++;
            CopyWordNote(run.Endnote.Paragraphs, run.Endnote.ReferenceId, OdtNoteKind.Endnote, target, notes);
            if (HasAmbiguousWordNotePosition(run)) notes.ApproximatedReferencePositions++;
            if (HasCustomWordNoteMark(notes, run.Endnote.ReferenceId, OdtNoteKind.Endnote))
                notes.ApproximatedCitations++;
            if (HasStyledWordNoteAnchor(notes, run.Endnote.ReferenceId, OdtNoteKind.Endnote))
                notes.ApproximatedReferenceFormatting++;
        }
    }

    private static bool HasCustomWordNoteMark(NoteMappingStats notes, long? referenceId, OdtNoteKind kind) {
        IReadOnlyList<W.Run> runs = notes.Anchors(kind, referenceId);
        return kind == OdtNoteKind.Footnote
            ? runs.Any(run => run.Elements<W.FootnoteReference>().Any(reference =>
                reference.Id?.Value == referenceId && reference.CustomMarkFollows?.Value == true))
            : runs.Any(run => run.Elements<W.EndnoteReference>().Any(reference =>
                reference.Id?.Value == referenceId && reference.CustomMarkFollows?.Value == true));
    }

    private static bool HasStyledWordNoteAnchor(NoteMappingStats notes, long? referenceId, OdtNoteKind kind) {
        string defaultStyle = kind == OdtNoteKind.Footnote ? "FootnoteReference" : "EndnoteReference";
        return notes.Anchors(kind, referenceId).Any(run =>
            run.RunProperties?.ChildElements.Any(child =>
                child is not W.RunStyle style || style.Val?.Value != defaultStyle) == true);
    }

    private static bool HasAmbiguousWordNotePosition(WordRunSnapshot run) =>
        run.Text.Length > 0 || run.PositionedImages.Count > 0 || run.NonTextBreaks?.Count > 0 ||
        run.IsHyperlink || (run.Footnote != null && run.Endnote != null);

    private static bool HasOdtParagraphNoteReferenceFormatting(OdtParagraph paragraph) =>
        paragraph.Bold.HasValue || paragraph.Italic.HasValue || paragraph.Underline.HasValue ||
        paragraph.StrikeThrough.HasValue || paragraph.SmallCaps.HasValue ||
        paragraph.FontSize.HasValue || paragraph.TextPosition.HasValue ||
        paragraph.TextTransform.HasValue || paragraph.Color.HasValue ||
        paragraph.TextBackgroundColor.HasValue || paragraph.FontFamily != null ||
        paragraph.UnderlineStyle.HasValue || paragraph.UnderlineType.HasValue ||
        paragraph.LineThroughStyle.HasValue || paragraph.LineThroughType.HasValue;

    private static void CopyWordNote(IReadOnlyList<WordParagraphSnapshot> paragraphs, long? referenceId, OdtNoteKind kind,
        OdtParagraph target, NoteMappingStats notes) {
        if (referenceId.HasValue && !(kind == OdtNoteKind.Footnote
            ? notes.ProcessedFootnoteIds : notes.ProcessedEndnoteIds).Add(referenceId.Value)) {
            return;
        }
        if (paragraphs.Count == 0) {
            CountUnsupported(kind, notes);
            return;
        }
        OdtNote result = kind == OdtNoteKind.Footnote
            ? target.AddFootnote(paragraphs[0].Text)
            : target.AddEndnote(paragraphs[0].Text);
        for (int index = 1; index < paragraphs.Count; index++) result.AddParagraph(paragraphs[index].Text);
        if (paragraphs.Any(paragraph => HasNonPlainWordNoteContent(paragraph, kind)) ||
            HasStyledWordNoteBodyRun(notes, referenceId, kind) ||
            HasNonDefaultWordNoteReferenceMark(notes, referenceId, kind))
            notes.ApproximatedBodies++;
        bool unsupportedInline = paragraphs.Any(paragraph => paragraph.Runs.Any(run =>
            run.InlineImage != null || run.Footnote != null || run.Endnote != null)) ||
            HasUnsupportedWordNoteInline(notes, referenceId, kind);
        bool unsupportedBlock = HasUnsupportedWordNoteBlocks(notes, referenceId, kind);
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

    private static bool HasUnsupportedWordNoteBlocks(NoteMappingStats notes, long? referenceId, OdtNoteKind kind) {
        DocumentFormat.OpenXml.OpenXmlElement? note = notes.Body(kind, referenceId);
        return note != null && note.ChildElements.Any(child => child is not W.Paragraph);
    }

    private static bool HasUnsupportedWordNoteInline(NoteMappingStats notes, long? referenceId, OdtNoteKind kind) {
        DocumentFormat.OpenXml.OpenXmlElement? note = notes.Body(kind, referenceId);
        return note?.Elements<W.Paragraph>().Any(paragraph =>
            paragraph.ChildElements.Any(child => child is not W.ParagraphProperties and not W.Run and
                not W.Hyperlink)) == true ||
            note?.Descendants<W.Run>().Any(run => run.ChildElements.Any(child =>
            child is not W.RunProperties and not W.Text and not W.FootnoteReferenceMark and
                not W.EndnoteReferenceMark and not W.TabChar and not W.CarriageReturn and
                not W.NoBreakHyphen and not W.SoftHyphen &&
                (child is not W.Break lineBreak || lineBreak.Type != null &&
                    lineBreak.Type.Value != W.BreakValues.TextWrapping))) == true;
    }

    private static bool HasNonDefaultWordNoteReferenceMark(NoteMappingStats notes, long? referenceId, OdtNoteKind kind) {
        DocumentFormat.OpenXml.OpenXmlElement? note = notes.Body(kind, referenceId);
        string defaultStyle = kind == OdtNoteKind.Footnote ? "FootnoteReference" : "EndnoteReference";
        if (note == null) return false;
        int markCount = kind == OdtNoteKind.Footnote
            ? note.Descendants<W.FootnoteReferenceMark>().Count()
            : note.Descendants<W.EndnoteReferenceMark>().Count();
        if (markCount != 1) return true;
        DocumentFormat.OpenXml.OpenXmlElement? firstVisible = note.Elements<W.Paragraph>()
            .SelectMany(paragraph => paragraph.Descendants<W.Run>())
            .SelectMany(run => run.ChildElements)
            .FirstOrDefault(child => child is W.FootnoteReferenceMark or W.EndnoteReferenceMark or
                W.TabChar or W.Break || child is W.Text text && !string.IsNullOrEmpty(text.Text));
        if (kind == OdtNoteKind.Footnote && firstVisible is not W.FootnoteReferenceMark ||
            kind == OdtNoteKind.Endnote && firstVisible is not W.EndnoteReferenceMark) return true;
        return note?.Descendants<W.Run>().Any(run =>
            run.ChildElements.Any(child => kind == OdtNoteKind.Footnote
                ? child is W.FootnoteReferenceMark : child is W.EndnoteReferenceMark) &&
            run.RunProperties?.ChildElements.Any(child =>
                child is not W.RunStyle style || style.Val?.Value != defaultStyle) == true) == true;
    }

    private static bool HasStyledWordNoteBodyRun(NoteMappingStats notes, long? referenceId, OdtNoteKind kind) {
        DocumentFormat.OpenXml.OpenXmlElement? note = notes.Body(kind, referenceId);
        return note?.Descendants<W.Run>().Any(run =>
            (!run.ChildElements.Any(child => kind == OdtNoteKind.Footnote
                ? child is W.FootnoteReferenceMark : child is W.EndnoteReferenceMark) ||
             run.ChildElements.OfType<W.Text>().Any(text => !string.IsNullOrEmpty(text.Text))) &&
            run.RunProperties?.GetFirstChild<W.RunStyle>() != null) == true;
    }

    private static bool HasNonPlainWordNoteContent(WordParagraphSnapshot paragraph, OdtNoteKind kind) =>
        HasNonDefaultWordNoteStyle(paragraph, kind) ||
        paragraph.Runs.Any(run => run.NonTextBreaks?.Values.Any(kind => kind != WordBreakType.TextWrapping) == true ||
            run.Text.Length > 0 && (
            run.Bold || run.Italic || run.Underline || run.Strike || run.DoubleStrike ||
            run.IsHyperlink ||
            run.FontSizePoints.HasValue || run.FontFamily != null || run.ColorHex != null ||
            run.HighlightColor != null || run.RunShadingFillColorHex != null ||
            run.RunShadingPattern.HasValue || run.VerticalTextAlignment != null || run.CapsStyle != null)) ||
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
        bool documentFootnoteSettings = HasDocumentNoteSettings<W.FootnoteDocumentWideProperties>(source);
        bool documentEndnoteSettings = HasDocumentNoteSettings<W.EndnoteDocumentWideProperties>(source);
        foreach (KeyValuePair<int, int> entry in notes.FootnotesBySection) {
            if (documentFootnoteSettings || HasSectionFootnoteSettings(source.Sections[entry.Key].FootnoteSettings))
                affectedNotes += entry.Value;
        }
        foreach (KeyValuePair<int, int> entry in notes.EndnotesBySection) {
            if (documentEndnoteSettings || HasSectionEndnoteSettings(source.Sections[entry.Key].EndnoteSettings))
                affectedNotes += entry.Value;
        }
        bool hasUnreferencedConfiguration = documentFootnoteSettings || documentEndnoteSettings ||
            source.Sections.Any(section => HasSectionFootnoteSettings(section.FootnoteSettings) ||
                HasSectionEndnoteSettings(section.EndnoteSettings));
        notes.ApproximatedNumberingAndPlacement = hasUnreferencedConfiguration
            ? Math.Max(1, affectedNotes) : affectedNotes;
        notes.ApproximatedSeparators += CountCustomizedWordSeparators(
            source.OpenXmlDocument.MainDocumentPart?.FootnotesPart?.Footnotes?.Elements<W.Footnote>()) +
            CountCustomizedWordSeparatorReferences(source, OdtNoteKind.Footnote);
        notes.ApproximatedSeparators += CountCustomizedWordSeparators(
            source.OpenXmlDocument.MainDocumentPart?.EndnotesPart?.Endnotes?.Elements<W.Endnote>()) +
            CountCustomizedWordSeparatorReferences(source, OdtNoteKind.Endnote);
        if (notes.ConvertedFootnotes > 0 && HasCustomizedDefaultWordNoteStyle(source, "FootnoteText"))
            notes.ApproximatedBodies += notes.ConvertedFootnotes;
        if (notes.ConvertedFootnotes > 0 && HasCustomizedDefaultWordNoteReferenceStyle(source, "FootnoteReference"))
            notes.ApproximatedReferenceFormatting += notes.ConvertedFootnotes;
        if (notes.ConvertedEndnotes > 0 && HasCustomizedDefaultWordNoteStyle(source, "EndnoteText"))
            notes.ApproximatedBodies += notes.ConvertedEndnotes;
        if (notes.ConvertedEndnotes > 0 && HasCustomizedDefaultWordNoteReferenceStyle(source, "EndnoteReference"))
            notes.ApproximatedReferenceFormatting += notes.ConvertedEndnotes;
        if (HasAuthoredWordDocumentDefaults(source))
            notes.ApproximatedBodies += notes.ConvertedFootnotes + notes.ConvertedEndnotes;
        if (HasAuthoredWordDocumentRunDefaults(source))
            notes.ApproximatedReferenceFormatting += notes.ConvertedFootnotes + notes.ConvertedEndnotes;
    }

    private static bool HasAuthoredWordDocumentDefaults(WordDocument source) {
        W.DocDefaults? defaults = source.OpenXmlDocument.MainDocumentPart?.StyleDefinitionsPart?.Styles?.DocDefaults;
        W.ParagraphPropertiesBaseStyle? paragraph = defaults?.ParagraphPropertiesDefault?.ParagraphPropertiesBaseStyle;
        if (HasAuthoredWordDocumentRunDefaults(source) ||
            paragraph?.ChildElements.Any(child => child is not W.SpacingBetweenLines) == true)
            return true;
        W.SpacingBetweenLines? spacing = paragraph?.GetFirstChild<W.SpacingBetweenLines>();
        return spacing != null && (spacing.After?.Value != "160" || spacing.Line?.Value != "259" ||
            spacing.LineRule?.Value != W.LineSpacingRuleValues.Auto);
    }

    private static bool HasAuthoredWordDocumentRunDefaults(WordDocument source) {
        W.RunPropertiesBaseStyle? run = source.OpenXmlDocument.MainDocumentPart?.StyleDefinitionsPart?
            .Styles?.DocDefaults?.RunPropertiesDefault?.RunPropertiesBaseStyle;
        if (run?.ChildElements.Any(child => child is not W.RunFonts and not W.FontSize and
            not W.FontSizeComplexScript and not W.Languages) == true) return true;
        W.RunFonts? fonts = run?.GetFirstChild<W.RunFonts>();
        if (fonts != null && (fonts.AsciiTheme?.Value != W.ThemeFontValues.MinorHighAnsi ||
            fonts.HighAnsiTheme?.Value != W.ThemeFontValues.MinorHighAnsi ||
            fonts.EastAsiaTheme?.Value != W.ThemeFontValues.MinorHighAnsi ||
            fonts.ComplexScriptTheme?.Value != W.ThemeFontValues.MinorBidi ||
            fonts.Ascii != null || fonts.HighAnsi != null || fonts.EastAsia != null ||
            fonts.ComplexScript != null)) return true;
        W.FontSize? size = run?.GetFirstChild<W.FontSize>();
        W.FontSizeComplexScript? complexSize = run?.GetFirstChild<W.FontSizeComplexScript>();
        if (size != null && size.Val?.Value != "22" ||
            complexSize != null && complexSize.Val?.Value != "22") return true;
        W.Languages? languages = run?.GetFirstChild<W.Languages>();
        if (languages != null && (languages.Val?.Value != "en-US" ||
            languages.EastAsia?.Value != "en-US" || languages.Bidi?.Value != "ar-SA")) return true;
        return false;
    }

    private static int CountCustomizedWordSeparators<T>(IEnumerable<T>? candidates)
        where T : DocumentFormat.OpenXml.OpenXmlElement =>
        candidates?.Count(note => {
            bool separator = note is W.Footnote footnote &&
                (footnote.Type?.Value == W.FootnoteEndnoteValues.Separator ||
                 footnote.Type?.Value == W.FootnoteEndnoteValues.ContinuationSeparator) ||
                note is W.Endnote endnote &&
                (endnote.Type?.Value == W.FootnoteEndnoteValues.Separator ||
                 endnote.Type?.Value == W.FootnoteEndnoteValues.ContinuationSeparator);
            if (!separator) return false;
            W.Paragraph[] paragraphs = note.Elements<W.Paragraph>().ToArray();
            if (paragraphs.Length != 1 || note.ChildElements.Count != 1) return true;
            W.Paragraph paragraph = paragraphs[0];
            W.Run[] runs = paragraph.Elements<W.Run>().ToArray();
            if (runs.Length != 1 || paragraph.ChildElements.Count != runs.Length +
                (paragraph.ParagraphProperties == null ? 0 : 1)) return true;
            W.Run run = runs[0];
            if (run.ChildElements.Count != 1 || run.FirstChild is not W.SeparatorMark and not W.ContinuationSeparatorMark)
                return true;
            W.SpacingBetweenLines? spacing = paragraph.ParagraphProperties?.GetFirstChild<W.SpacingBetweenLines>();
            return paragraph.ParagraphProperties != null &&
                (paragraph.ParagraphProperties.ChildElements.Count != 1 || spacing == null ||
                 spacing.After?.Value != "0" || spacing.Line?.Value != "240" ||
                 spacing.LineRule?.Value != W.LineSpacingRuleValues.Auto);
        }) ?? 0;

    private static int CountCustomizedWordSeparatorReferences(WordDocument source, OdtNoteKind kind) {
        W.Settings? settings = source.OpenXmlDocument.MainDocumentPart?.DocumentSettingsPart?.Settings;
        IEnumerable<DocumentFormat.OpenXml.OpenXmlElement> properties = kind == OdtNoteKind.Footnote
            ? settings?.Elements<W.FootnoteDocumentWideProperties>() ?? Enumerable.Empty<DocumentFormat.OpenXml.OpenXmlElement>()
            : settings?.Elements<W.EndnoteDocumentWideProperties>() ?? Enumerable.Empty<DocumentFormat.OpenXml.OpenXmlElement>();
        properties = properties.Concat(source.OpenXmlDocument.MainDocumentPart?.Document?.Body?
            .Descendants<W.SectionProperties>().SelectMany(section => kind == OdtNoteKind.Footnote
                ? section.Elements<W.FootnoteProperties>().Cast<DocumentFormat.OpenXml.OpenXmlElement>()
                : section.Elements<W.EndnoteProperties>().Cast<DocumentFormat.OpenXml.OpenXmlElement>())
            ?? Enumerable.Empty<DocumentFormat.OpenXml.OpenXmlElement>());
        return kind == OdtNoteKind.Footnote
            ? properties.SelectMany(property => property.Elements<W.FootnoteSpecialReference>())
                .Count(reference => reference.Id?.Value is not -1 and not 0)
            : properties.SelectMany(property => property.Elements<W.EndnoteSpecialReference>())
                .Count(reference => reference.Id?.Value is not -1 and not 0);
    }

    private static bool HasSectionFootnoteSettings(WordFootnoteSettings settings) =>
        settings.Position.HasValue || settings.NumberingRestart.HasValue ||
        settings.StartNumber.HasValue || settings.NumberingFormat.HasValue;

    private static bool HasSectionEndnoteSettings(WordEndnoteSettings settings) =>
        settings.Position.HasValue || settings.NumberingRestart.HasValue ||
        settings.StartNumber.HasValue || settings.NumberingFormat.HasValue;

    private static void CountOdtNoteConfigurationLoss(OdtDocument source, NoteMappingStats notes) {
        System.Xml.Linq.XNamespace text = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
        System.Xml.Linq.XNamespace style = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
        foreach (string part in new[] { "styles.xml", "content.xml" }) {
            if (!source.Package.ContainsEntry(part)) continue;
            System.Xml.Linq.XDocument document = source.Package.GetXml(part);
            notes.UnsupportedOdtNoteConfigurations += document.Descendants(text + "notes-configuration").Count();
            notes.UnsupportedOdtNoteConfigurations += document.Descendants(style + "footnote-sep").Count();
        }
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

    private static bool HasCustomizedDefaultWordNoteReferenceStyle(WordDocument source, string styleId) {
        W.Styles? styles = source.OpenXmlDocument.MainDocumentPart?.StyleDefinitionsPart?.Styles;
        W.Style? style = styles?
            .Elements<W.Style>().FirstOrDefault(candidate => string.Equals(candidate.StyleId?.Value,
                styleId, StringComparison.OrdinalIgnoreCase));
        if (style == null) return false;
        if (style.CustomStyle?.Value == true || style.BasedOn?.Val?.Value != "DefaultParagraphFont") return true;
        W.Style? baseStyle = styles?.Elements<W.Style>().FirstOrDefault(candidate =>
            string.Equals(candidate.StyleId?.Value, "DefaultParagraphFont", StringComparison.OrdinalIgnoreCase));
        if (baseStyle?.StyleRunProperties?.HasChildren == true || baseStyle?.BasedOn != null) return true;
        W.StyleRunProperties? properties = style.StyleRunProperties;
        return properties?.ChildElements.Count != 1 ||
            properties.GetFirstChild<W.VerticalTextAlignment>()?.Val?.Value != W.VerticalPositionValues.Superscript;
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
        if (paragraphs.Any(paragraph => ContainsInlineKind(paragraph.InlineNodes, OdtInlineNodeKind.Note))) {
            CountUnsupported(source.Kind.Value, notes);
            return;
        }
        if (paragraphs.Any(paragraph => ContainsUnsupportedNoteBodyInline(paragraph.InlineNodes)))
            notes.UnsupportedBodyContent++;
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
        if (notes.HasOdtDefaultNoteBodyFormatting || paragraphs.Count != 1 || paragraphs.Any(paragraph => paragraph.StyleName != null ||
            paragraph.InlineNodes.Any(node => node.Kind != OdtInlineNodeKind.Text))) notes.ApproximatedBodies++;
    }

    private static bool ContainsInlineKind(IReadOnlyList<OdtInlineNode> nodes, OdtInlineNodeKind kind) =>
        nodes.Any(node => node.Kind == kind || ContainsInlineKind(node.Children, kind));

    private static bool ContainsUnsupportedNoteBodyInline(IReadOnlyList<OdtInlineNode> nodes) =>
        nodes.Any(node => node.Kind is OdtInlineNodeKind.Image or OdtInlineNodeKind.Other ||
            ContainsUnsupportedNoteBodyInline(node.Children));

    private static bool HasOdtDefaultNoteBodyFormatting(OdtDocument source) {
        if (!source.Package.ContainsEntry("styles.xml")) return false;
        System.Xml.Linq.XNamespace style = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
        return source.Package.GetXml("styles.xml").Descendants(style + "default-style")
            .Where(element => (string?)element.Attribute(style + "family") is "paragraph" or "text")
            .SelectMany(element => element.Elements())
            .Any(element => (element.Name == style + "paragraph-properties" ||
                             element.Name == style + "text-properties") &&
                            (element.HasAttributes || element.HasElements));
    }

    private static bool HasOdtDefaultNoteReferenceFormatting(OdtDocument source) {
        if (!source.Package.ContainsEntry("styles.xml")) return false;
        System.Xml.Linq.XNamespace style = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
        return source.Package.GetXml("styles.xml").Descendants(style + "default-style")
            .Where(element => (string?)element.Attribute(style + "family") is "paragraph" or "text")
            .SelectMany(element => element.Elements(style + "text-properties"))
            .Any(element => element.HasAttributes || element.HasElements);
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
        if (notes.UnreferencedFootnoteDefinitions > 0) report.Add("source-footnotes",
            OdfConversionMappingStatus.Unsupported, notes.UnreferencedFootnoteDefinitions,
            "Word footnote definitions without a body reference or with duplicate IDs were omitted.");
        if (notes.UnreferencedEndnoteDefinitions > 0) report.Add("source-endnotes",
            OdfConversionMappingStatus.Unsupported, notes.UnreferencedEndnoteDefinitions,
            "Word endnote definitions without a body reference or with duplicate IDs were omitted.");
        if (notes.UnsupportedFootnotes > 0) report.Add("footnotes", OdfConversionMappingStatus.Unsupported,
            notes.UnsupportedFootnotes, "Note references with unsupported bodies, classes, or repeated references to one Word note were omitted.");
        if (notes.UnsupportedEndnotes > 0) report.Add("endnotes", OdfConversionMappingStatus.Unsupported,
            notes.UnsupportedEndnotes, "Note references with unsupported bodies or repeated references to one Word note were omitted.");
        if (notes.ApproximatedBodies > 0) report.Add("note-body-formatting", OdfConversionMappingStatus.Approximated,
            notes.ApproximatedBodies, "Note text was retained, but some body formatting, paragraph structure, or reference mark placement was flattened.");
        if (notes.UnsupportedBodyContent > 0) report.Add("note-body-content", OdfConversionMappingStatus.Unsupported,
            notes.UnsupportedBodyContent, "Note text was retained, but embedded media or unsupported block or inline content was omitted.");
        if (notes.UnsupportedHeaderFooterNotes > 0) report.Add("note-headers-footers", OdfConversionMappingStatus.Unsupported,
            notes.UnsupportedHeaderFooterNotes, "ODT header and footer note references were omitted because Word does not permit them there.");
        if (notes.UnsupportedOdtNoteConfigurations > 0) report.Add("note-configuration", OdfConversionMappingStatus.Unsupported,
            notes.UnsupportedOdtNoteConfigurations, "ODT note numbering, placement, or footnote separator configuration was not carried into Word.");
        if (notes.ApproximatedReferencePositions > 0) report.Add("note-reference-position", OdfConversionMappingStatus.Approximated,
            notes.ApproximatedReferencePositions,
            "A Word run contained a note reference with text, another note, a break, an image, or a hyperlink; inline context or order may have changed.");
        if (notes.ApproximatedReferenceFormatting > 0) report.Add("note-reference-formatting", OdfConversionMappingStatus.Approximated,
            notes.ApproximatedReferenceFormatting, "Direct or inherited note-reference formatting was replaced by the destination's default note reference.");
        if (notes.ApproximatedCitations > 0) report.Add("note-citations", OdfConversionMappingStatus.Approximated,
            notes.ApproximatedCitations, "Custom Word or ODT note marks were replaced with automatic numbering.");
        if (notes.ApproximatedSeparators > 0) report.Add("note-separators", OdfConversionMappingStatus.Approximated,
            notes.ApproximatedSeparators, "Customized Word note separators were not carried into ODT.");
        if (notes.ApproximatedNumberingAndPlacement > 0) report.Add("note-numbering-placement", OdfConversionMappingStatus.Approximated,
            notes.ApproximatedNumberingAndPlacement,
            "Word note numbering format, start, restart, or placement settings were not carried into ODT.");
    }
}
