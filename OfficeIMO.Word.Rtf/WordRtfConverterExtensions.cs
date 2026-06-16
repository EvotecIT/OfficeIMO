using System.Linq;
using System.Collections.Generic;
using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Rtf;

/// <summary>
/// Extension methods for converting between <see cref="WordDocument"/> and <see cref="RtfDocument"/>.
/// </summary>
public static partial class WordRtfConverterExtensions {
    /// <summary>Converts a Word document to the semantic RTF model.</summary>
    public static RtfDocument ToRtfDocument(this WordDocument document) {
        if (document == null) throw new ArgumentNullException(nameof(document));

        RtfDocument rtf = RtfDocument.Create();
        CopyDocumentInfo(document, rtf);
        CopyCustomMetadata(document, rtf);
        CopyDefaultLanguage(document, rtf);
        CopyDocumentSettings(document, rtf);
        var revisionAuthorIndexes = new Dictionary<string, int>(StringComparer.Ordinal);
        CopyHeaderFooters(document, rtf, revisionAuthorIndexes);
        if (ShouldExportSections(document)) {
            CopySections(document, rtf, revisionAuthorIndexes);
        } else {
            CopyPageSetup(document, rtf);
            CopyWordElements(document.Elements, rtf, revisionAuthorIndexes);
        }

        return rtf;
    }

    /// <summary>Creates a Word document from a semantic RTF document.</summary>
    public static WordDocument ToWordDocument(this RtfDocument rtfDocument) {
        if (rtfDocument == null) throw new ArgumentNullException(nameof(rtfDocument));

        WordDocument document = WordDocument.Create();
        ApplyDocumentInfo(rtfDocument, document);
        ApplyCustomMetadata(rtfDocument, document);
        ApplyDefaultLanguage(rtfDocument, document);
        ApplyDocumentSettings(rtfDocument, document);
        ApplyHeaderFooters(rtfDocument, document);
        if (rtfDocument.Sections.Count > 0) {
            ApplySections(rtfDocument, document);
        } else {
            ApplyPageSetup(rtfDocument, document);
            foreach (IRtfBlock block in rtfDocument.Blocks) {
                if (block is RtfParagraph paragraph) {
                    AppendParagraph(document, paragraph, rtfDocument);
                } else if (block is RtfTable table) {
                    AppendTable(document, table, rtfDocument);
                } else if (block is RtfImage image) {
                    AppendImage(document, image);
                }
            }
        }

        return document;
    }

    private static void CopyHeaderFooters(WordDocument document, RtfDocument rtf, Dictionary<string, int> revisionAuthorIndexes) {
        CopyHeaderFooter(document.Header?.Default, rtf, rtf.AddHeader, RtfHeaderFooterKind.Header, revisionAuthorIndexes);
        CopyHeaderFooter(document.Header?.First, rtf, rtf.AddHeader, RtfHeaderFooterKind.FirstHeader, revisionAuthorIndexes);
        CopyHeaderFooter(document.Header?.Even, rtf, rtf.AddHeader, RtfHeaderFooterKind.LeftHeader, revisionAuthorIndexes);
        CopyHeaderFooter(document.Footer?.Default, rtf, rtf.AddFooter, RtfHeaderFooterKind.Footer, revisionAuthorIndexes);
        CopyHeaderFooter(document.Footer?.First, rtf, rtf.AddFooter, RtfHeaderFooterKind.FirstFooter, revisionAuthorIndexes);
        CopyHeaderFooter(document.Footer?.Even, rtf, rtf.AddFooter, RtfHeaderFooterKind.LeftFooter, revisionAuthorIndexes);
    }

    private static void CopyHeaderFooter(WordHeaderFooter? source, RtfDocument rtf, Func<RtfHeaderFooterKind, RtfHeaderFooter> addDestination, RtfHeaderFooterKind kind, Dictionary<string, int> revisionAuthorIndexes) {
        if (source == null || source.Paragraphs.Count == 0) {
            return;
        }

        RtfHeaderFooter destination = addDestination(kind);
        foreach (WordParagraph wordParagraph in source.Paragraphs.GroupBy(paragraph => paragraph._paragraph).Select(group => group.First())) {
            RtfParagraph paragraph = destination.AddParagraph();
            CopyTabStops(wordParagraph, paragraph);
            CopyParagraphFormatting(wordParagraph, paragraph, rtf);
            AppendFormattedRuns(wordParagraph, paragraph, rtf, revisionAuthorIndexes);
        }
    }

    private static void CopyTabStops(WordParagraph source, RtfParagraph destination) {
        foreach (WordTabStop tabStop in source.TabStops) {
            destination.AddTabStop(tabStop.Position, ToRtfTabAlignment(tabStop.Alignment), ToRtfTabLeader(tabStop.Leader));
        }
    }

    private static RtfTabAlignment ToRtfTabAlignment(TabStopValues alignment) {
        if (alignment == TabStopValues.Center) return RtfTabAlignment.Center;
        if (alignment == TabStopValues.Right) return RtfTabAlignment.Right;
        if (alignment == TabStopValues.Decimal) return RtfTabAlignment.Decimal;
        if (alignment == TabStopValues.Bar) return RtfTabAlignment.Bar;
        return RtfTabAlignment.Left;
    }

    private static RtfTabLeader ToRtfTabLeader(TabStopLeaderCharValues leader) {
        if (leader == TabStopLeaderCharValues.Dot) return RtfTabLeader.Dots;
        if (leader == TabStopLeaderCharValues.MiddleDot) return RtfTabLeader.MiddleDots;
        if (leader == TabStopLeaderCharValues.Hyphen) return RtfTabLeader.Hyphen;
        if (leader == TabStopLeaderCharValues.Underscore) return RtfTabLeader.Underline;
        if (leader == TabStopLeaderCharValues.Heavy) return RtfTabLeader.ThickLine;
        return RtfTabLeader.None;
    }

    private static void ApplyHeaderFooters(RtfDocument rtfDocument, WordDocument document) {
        foreach (RtfHeaderFooter headerFooter in rtfDocument.HeaderFooters) {
            WordHeaderFooter target = GetWordHeaderFooter(document, headerFooter.Kind);
            foreach (RtfParagraph paragraph in headerFooter.Paragraphs) {
                AppendParagraph(target, paragraph, rtfDocument);
            }
        }
    }

    private static WordHeaderFooter GetWordHeaderFooter(WordDocument document, RtfHeaderFooterKind kind) {
        switch (kind) {
            case RtfHeaderFooterKind.FirstHeader:
                return document.HeaderFirstOrCreate;
            case RtfHeaderFooterKind.LeftHeader:
                return document.HeaderEvenOrCreate;
            case RtfHeaderFooterKind.RightHeader:
            case RtfHeaderFooterKind.Header:
                return document.HeaderDefaultOrCreate;
            case RtfHeaderFooterKind.FirstFooter:
                return document.FooterFirstOrCreate;
            case RtfHeaderFooterKind.LeftFooter:
                return document.FooterEvenOrCreate;
            case RtfHeaderFooterKind.RightFooter:
            case RtfHeaderFooterKind.Footer:
                return document.FooterDefaultOrCreate;
            default:
                throw new ArgumentOutOfRangeException(nameof(kind), kind, "Unsupported RTF header or footer kind.");
        }
    }

    private static void AppendFormattedRuns(WordParagraph wordParagraph, RtfParagraph paragraph, RtfDocument rtfDocument, Dictionary<string, int> revisionAuthorIndexes) {
        bool hasRuns = false;
        RtfRun? previousRun = null;
        var complexFields = new Stack<ComplexFieldCapture>();
        var bookmarkNames = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (OpenXmlElement element in wordParagraph._paragraph.ChildElements) {
            switch (element) {
                case BookmarkStart bookmarkStart when bookmarkStart.Name?.Value != null && bookmarkStart.Id?.Value != null:
                    bookmarkNames[bookmarkStart.Id.Value] = bookmarkStart.Name.Value;
                    paragraph.AddBookmarkStart(bookmarkStart.Name.Value);
                    break;
                case BookmarkEnd bookmarkEnd when bookmarkEnd.Id?.Value != null && bookmarkNames.TryGetValue(bookmarkEnd.Id.Value, out string? bookmarkName):
                    paragraph.AddBookmarkEnd(bookmarkName);
                    break;
                case Run runElement:
                    if (TryAppendComplexFieldRun(wordParagraph, runElement, paragraph, rtfDocument, revisionAuthorIndexes, complexFields)) {
                        previousRun = null;
                        hasRuns = true;
                        break;
                    }

                    hasRuns |= AppendWordRun(new WordParagraph(wordParagraph._document, wordParagraph._paragraph, runElement), paragraph, ref previousRun, rtfDocument, revisionAuthorIndexes);
                    break;
                case InsertedRun insertedRun:
                    int? insertedAuthorIndex = GetOrAddRevisionAuthorIndex(rtfDocument, insertedRun.Author?.Value, revisionAuthorIndexes);
                    foreach (Run childRun in insertedRun.Elements<Run>()) {
                        hasRuns |= AppendWordRun(new WordParagraph(wordParagraph._document, wordParagraph._paragraph, childRun), paragraph, ref previousRun, rtfDocument, revisionAuthorIndexes, RtfRevisionKind.Inserted, insertedAuthorIndex);
                    }

                    break;
                case DeletedRun deletedRun:
                    int? deletedAuthorIndex = GetOrAddRevisionAuthorIndex(rtfDocument, deletedRun.Author?.Value, revisionAuthorIndexes);
                    foreach (Run childRun in deletedRun.Elements<Run>()) {
                        hasRuns |= AppendWordRun(new WordParagraph(wordParagraph._document, wordParagraph._paragraph, childRun), paragraph, ref previousRun, rtfDocument, revisionAuthorIndexes, RtfRevisionKind.Deleted, deletedAuthorIndex);
                    }

                    break;
                case Hyperlink hyperlink:
                    foreach (Run childRun in hyperlink.Elements<Run>()) {
                        var wordRun = new WordParagraph(wordParagraph._document, wordParagraph._paragraph, childRun) {
                            _hyperlink = hyperlink
                        };
                        hasRuns |= AppendWordRun(wordRun, paragraph, ref previousRun, rtfDocument, revisionAuthorIndexes);
                    }

                    break;
                case SimpleField simpleField:
                    AppendSimpleField(wordParagraph, simpleField, paragraph, rtfDocument, revisionAuthorIndexes);
                    hasRuns = true;
                    break;
            }
        }

        if (!hasRuns && !string.IsNullOrEmpty(wordParagraph.Text)) {
            paragraph.AddText(wordParagraph.Text);
        }

        AppendWordComments(wordParagraph, paragraph, rtfDocument, revisionAuthorIndexes);
    }

    private static bool TryAppendComplexFieldRun(WordParagraph wordParagraph, Run runElement, RtfParagraph paragraph, RtfDocument rtfDocument, Dictionary<string, int> revisionAuthorIndexes, Stack<ComplexFieldCapture> captures) {
        FieldChar? fieldChar = runElement.Elements<FieldChar>().FirstOrDefault();
        if (fieldChar?.FieldCharType?.Value == FieldCharValues.Begin) {
            captures.Push(new ComplexFieldCapture());
            return true;
        }

        if (captures.Count == 0) {
            return false;
        }

        ComplexFieldCapture capture = captures.Peek();
        if (fieldChar?.FieldCharType?.Value == FieldCharValues.Separate) {
            capture.CapturingResult = true;
            return true;
        }

        if (fieldChar?.FieldCharType?.Value == FieldCharValues.End) {
            captures.Pop();
            if (captures.Count == 0) {
                CompleteComplexField(paragraph, capture);
            } else {
                CompleteNestedComplexField(captures.Peek(), capture);
            }

            return true;
        }

        FieldCode? fieldCode = runElement.Elements<FieldCode>().FirstOrDefault();
        if (fieldCode != null && !capture.CapturingResult) {
            capture.Instruction.Append(fieldCode.Text);
            return true;
        }

        if (capture.CapturingResult) {
            RtfRun? previousRun = capture.PreviousRun;
            AppendWordRun(new WordParagraph(wordParagraph._document, wordParagraph._paragraph, runElement), capture.Result, ref previousRun, rtfDocument, revisionAuthorIndexes);
            capture.PreviousRun = previousRun;
            return true;
        }

        return true;
    }

    private static void CompleteComplexField(RtfParagraph paragraph, ComplexFieldCapture capture) {
        RtfField field = paragraph.AddField(capture.Instruction.ToString().Trim());
        CopyInlines(capture.Result, field.Result);
    }

    private static void CompleteNestedComplexField(ComplexFieldCapture parent, ComplexFieldCapture nested) {
        if (!parent.CapturingResult) {
            parent.Instruction.Append(nested.Instruction.ToString().Trim());
            return;
        }

        RtfField field = parent.Result.AddField(nested.Instruction.ToString().Trim());
        CopyInlines(nested.Result, field.Result);
        parent.PreviousRun = null;
    }

    private static void CopyInlines(RtfParagraph source, RtfParagraph destination) {
        foreach (IRtfInline inline in source.Inlines) {
            switch (inline) {
                case RtfRun run:
                    CopyRun(run, destination.AddText(run.Text));
                    break;
                case RtfBookmarkMarker marker when marker.Kind == RtfBookmarkMarkerKind.Start:
                    destination.AddBookmarkStart(marker.Name);
                    break;
                case RtfBookmarkMarker marker:
                    destination.AddBookmarkEnd(marker.Name);
                    break;
                case RtfField nestedField:
                    RtfField field = destination.AddField(nestedField.Instruction);
                    field.HyperlinkField = nestedField.HyperlinkField?.Clone();
                    CopyInlines(nestedField.Result, field.Result);
                    break;
                case RtfGeneratedText generatedText:
                    destination.AddGeneratedText(generatedText.Kind, generatedText.FallbackText).Note = generatedText.Note;
                    break;
                case RtfImage image:
                    RtfImage copy = destination.AddImage(image.Format, image.Data);
                    CopyImage(image, copy);
                    break;
                case RtfBreak rtfBreak:
                    destination.AddBreak(rtfBreak.Kind);
                    break;
            }
        }
    }

    private static void CopyRun(RtfRun source, RtfRun destination) {
        destination.Bold = source.Bold;
        destination.Italic = source.Italic;
        destination.UnderlineStyle = source.UnderlineStyle;
        destination.Strike = source.Strike;
        destination.DoubleStrike = source.DoubleStrike;
        destination.Hidden = source.Hidden;
        destination.Outline = source.Outline;
        destination.Shadow = source.Shadow;
        destination.Emboss = source.Emboss;
        destination.Imprint = source.Imprint;
        destination.CapsStyle = source.CapsStyle;
        destination.VerticalPosition = source.VerticalPosition;
        destination.FontSize = source.FontSize;
        destination.FontId = source.FontId;
        destination.ForegroundColorIndex = source.ForegroundColorIndex;
        destination.HighlightColorIndex = source.HighlightColorIndex;
        destination.UnderlineColorIndex = source.UnderlineColorIndex;
        destination.CharacterSpacingTwips = source.CharacterSpacingTwips;
        destination.CharacterScalePercent = source.CharacterScalePercent;
        destination.KerningHalfPoints = source.KerningHalfPoints;
        destination.CharacterOffsetHalfPoints = source.CharacterOffsetHalfPoints;
        destination.Direction = source.Direction;
        destination.LanguageId = source.LanguageId;
        destination.StyleId = source.StyleId;
        destination.Hyperlink = source.Hyperlink;
        destination.Note = source.Note;
        destination.RevisionKind = source.RevisionKind;
        destination.RevisionAuthorIndex = source.RevisionAuthorIndex;
        destination.RevisionTimestampValue = source.RevisionTimestampValue;
    }

    private static void AppendSimpleField(WordParagraph wordParagraph, SimpleField simpleField, RtfParagraph paragraph, RtfDocument rtfDocument, Dictionary<string, int> revisionAuthorIndexes) {
        string instruction = simpleField.Instruction?.Value ?? string.Empty;
        RtfField field = paragraph.AddField(instruction.Trim());
        RtfRun? previousRun = null;
        foreach (Run childRun in simpleField.Elements<Run>()) {
            AppendWordRun(new WordParagraph(wordParagraph._document, wordParagraph._paragraph, childRun), field.Result, ref previousRun, rtfDocument, revisionAuthorIndexes);
        }
    }

    private static bool AppendWordRun(WordParagraph wordRun, RtfParagraph paragraph, ref RtfRun? previousRun, RtfDocument rtfDocument, Dictionary<string, int> revisionAuthorIndexes, RtfRevisionKind revisionKind = RtfRevisionKind.None, int? revisionAuthorIndex = null) {
        if (wordRun.Break != null) {
            paragraph.AddBreak(ToRtfBreakKind(wordRun.Break.BreakType));
            previousRun = null;
            return true;
        }

        if (wordRun.FootNote != null) {
            if (previousRun != null) {
                previousRun.Note = CopyNote(wordRun.FootNote.Paragraphs, RtfNoteKind.Footnote, rtfDocument, revisionAuthorIndexes);
            }

            return false;
        }

        if (wordRun.EndNote != null) {
            if (previousRun != null) {
                previousRun.Note = CopyNote(wordRun.EndNote.Paragraphs, RtfNoteKind.Endnote, rtfDocument, revisionAuthorIndexes);
            }

            return false;
        }

        if (wordRun.IsImage && wordRun.Image != null) {
            RtfImage? image = CreateRtfImage(wordRun);
            if (image == null) {
                return false;
            }

            RtfImage copy = paragraph.AddImage(image.Format, image.Data);
            CopyImage(image, copy);
            previousRun = null;
            return true;
        }

        string? text = wordRun.Text;
        if (string.IsNullOrEmpty(text)) {
            return false;
        }

        string? hyperlink = wordRun.IsHyperLink && wordRun.Hyperlink != null ? wordRun.Hyperlink.Uri?.ToString() : null;
        RtfRun run = paragraph.AddText(text);
        ApplyWordRunFormatting(wordRun, run, rtfDocument);
        if (!string.IsNullOrWhiteSpace(hyperlink) &&
            Uri.TryCreate(hyperlink, UriKind.RelativeOrAbsolute, out Uri? hyperlinkUri)) {
            run.Hyperlink = hyperlinkUri;
        }

        run.RevisionKind = revisionKind;
        run.RevisionAuthorIndex = revisionAuthorIndex;

        previousRun = run;
        return true;
    }

    private static RtfBreakKind ToRtfBreakKind(BreakValues? breakType) {
        if (breakType == BreakValues.Page) return RtfBreakKind.Page;
        if (breakType == BreakValues.Column) return RtfBreakKind.Column;
        return RtfBreakKind.Line;
    }

    private static RtfNote CopyNote(List<WordParagraph>? paragraphs, RtfNoteKind kind, RtfDocument rtfDocument, Dictionary<string, int> revisionAuthorIndexes) {
        var note = new RtfNote(kind);
        if (paragraphs == null) {
            return note;
        }

        foreach (WordParagraph wordParagraph in paragraphs.GroupBy(paragraph => paragraph._paragraph).Select(group => group.First())) {
            RtfParagraph paragraph = note.AddParagraph();
            CopyParagraphFormatting(wordParagraph, paragraph, rtfDocument);
            AppendFormattedRuns(wordParagraph, paragraph, rtfDocument, revisionAuthorIndexes);
        }

        return note;
    }

    private static void AppendParagraph(WordDocument document, RtfParagraph paragraph, RtfDocument rtfDocument) {
        WordParagraph wordParagraph = document.AddParagraph();
        ApplyTabStops(wordParagraph, paragraph);
        ApplyParagraphFormatting(wordParagraph, paragraph, rtfDocument);
        AppendRuns(wordParagraph, paragraph, rtfDocument);
    }

    private static void AppendParagraph(WordHeaderFooter headerFooter, RtfParagraph paragraph, RtfDocument rtfDocument) {
        WordParagraph wordParagraph = headerFooter.AddParagraph();
        ApplyTabStops(wordParagraph, paragraph);
        ApplyParagraphFormatting(wordParagraph, paragraph, rtfDocument);
        AppendRuns(wordParagraph, paragraph, rtfDocument);
    }

    private static void ApplyTabStops(WordParagraph destination, RtfParagraph source) {
        foreach (RtfTabStop tabStop in source.TabStops) {
            destination.AddTabStop(tabStop.PositionTwips, ToWordTabAlignment(tabStop.Alignment), ToWordTabLeader(tabStop.Leader));
        }
    }

    private static TabStopValues ToWordTabAlignment(RtfTabAlignment alignment) {
        switch (alignment) {
            case RtfTabAlignment.Center:
                return TabStopValues.Center;
            case RtfTabAlignment.Right:
                return TabStopValues.Right;
            case RtfTabAlignment.Decimal:
                return TabStopValues.Decimal;
            case RtfTabAlignment.Bar:
                return TabStopValues.Bar;
            default:
                return TabStopValues.Left;
        }
    }

    private static TabStopLeaderCharValues ToWordTabLeader(RtfTabLeader leader) {
        switch (leader) {
            case RtfTabLeader.Dots:
                return TabStopLeaderCharValues.Dot;
            case RtfTabLeader.MiddleDots:
                return TabStopLeaderCharValues.MiddleDot;
            case RtfTabLeader.Hyphen:
                return TabStopLeaderCharValues.Hyphen;
            case RtfTabLeader.Underline:
                return TabStopLeaderCharValues.Underscore;
            case RtfTabLeader.ThickLine:
                return TabStopLeaderCharValues.Heavy;
            case RtfTabLeader.EqualSign:
                return TabStopLeaderCharValues.MiddleDot;
            default:
                return TabStopLeaderCharValues.None;
        }
    }

    private static void AppendRuns(WordParagraph wordParagraph, RtfParagraph paragraph, RtfDocument? rtfDocument) {
        var openBookmarks = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (IRtfInline inline in paragraph.Inlines) {
            if (inline is RtfBookmarkMarker marker) {
                AppendBookmarkMarker(wordParagraph, marker, openBookmarks);
                continue;
            }

            if (inline is RtfField field) {
                AppendField(wordParagraph, field, rtfDocument);
                continue;
            }

            if (inline is RtfGeneratedText generatedText) {
                AppendGeneratedText(wordParagraph, generatedText, rtfDocument);
                continue;
            }

            if (inline is RtfBreak rtfBreak) {
                wordParagraph.AddBreak(ToWordBreakKind(rtfBreak.Kind));
                continue;
            }

            if (inline is RtfImage image) {
                AppendImage(wordParagraph, image);
                continue;
            }

            if (!(inline is RtfRun run)) {
                continue;
            }

            WordParagraph wordRun;
            if (run.RevisionKind == RtfRevisionKind.Inserted) {
                wordParagraph.AddInsertedText(run.Text, GetRevisionAuthorName(rtfDocument, run.RevisionAuthorIndex));
                continue;
            }

            if (run.RevisionKind == RtfRevisionKind.Deleted) {
                wordParagraph.AddDeletedText(run.Text, GetRevisionAuthorName(rtfDocument, run.RevisionAuthorIndex));
                continue;
            }

            if (run.Hyperlink != null) {
                wordRun = wordParagraph.AddHyperLink(run.Text, run.Hyperlink, addStyle: true);
            } else {
                wordRun = wordParagraph.AddText(run.Text);
            }

            ApplyRtfRunFormatting(run, wordRun, rtfDocument);

            if (run.Note != null) {
                AppendNote(wordRun, run.Note, rtfDocument);
            }
        }
    }

    private static void AppendGeneratedText(WordParagraph wordParagraph, RtfGeneratedText generatedText, RtfDocument? rtfDocument) {
        if (generatedText.Kind == RtfGeneratedTextKind.NoteReference) {
            WordParagraph wordRun = wordParagraph.AddText(generatedText.FallbackText ?? string.Empty);
            if (generatedText.Note != null) {
                AppendNote(wordRun, generatedText.Note, rtfDocument);
            }

            return;
        }

        var field = new RtfField(ToWordFieldInstruction(generatedText.Kind));
        if (!string.IsNullOrEmpty(generatedText.FallbackText)) {
            field.AddText(generatedText.FallbackText!);
        }

        AppendField(wordParagraph, field, rtfDocument);
        if (generatedText.Note != null) {
            WordParagraph wordRun = wordParagraph.AddText(generatedText.FallbackText ?? string.Empty);
            AppendNote(wordRun, generatedText.Note, rtfDocument);
        }
    }

    private static string ToWordFieldInstruction(RtfGeneratedTextKind kind) {
        switch (kind) {
            case RtfGeneratedTextKind.SectionNumber:
                return "SECTION";
            case RtfGeneratedTextKind.CurrentDate:
            case RtfGeneratedTextKind.CurrentDateLong:
            case RtfGeneratedTextKind.CurrentDateAbbreviated:
                return "DATE";
            case RtfGeneratedTextKind.CurrentTime:
                return "TIME";
            default:
                return "PAGE";
        }
    }

    private static BreakValues? ToWordBreakKind(RtfBreakKind kind) {
        switch (kind) {
            case RtfBreakKind.Page:
            case RtfBreakKind.SoftPage:
                return BreakValues.Page;
            case RtfBreakKind.Column:
                return BreakValues.Column;
            default:
                return null;
        }
    }

    private static void SetHiddenWordRun(WordParagraph wordRun) {
        if (wordRun._run == null) {
            return;
        }

        wordRun._run.RunProperties ??= new RunProperties();
        wordRun._run.RunProperties.Vanish = new Vanish();
    }

    private static void AppendField(WordParagraph wordParagraph, RtfField field, RtfDocument? rtfDocument) {
        if (field.Hyperlink != null) {
            AppendHyperlinkField(wordParagraph, field, rtfDocument);
            return;
        }

        var simpleField = new SimpleField { Instruction = field.Instruction };
        var resultParagraph = new WordParagraph(wordParagraph._document, newParagraph: true, newRun: false);
        AppendRuns(resultParagraph, field.Result, rtfDocument);
        foreach (OpenXmlElement child in resultParagraph._paragraph.ChildElements) {
            simpleField.Append(child.CloneNode(true));
        }

        wordParagraph._paragraph.Append(simpleField);
    }

    private static void AppendHyperlinkField(WordParagraph wordParagraph, RtfField field, RtfDocument? rtfDocument) {
        var resultParagraph = new WordParagraph(wordParagraph._document, newParagraph: true, newRun: false);
        AppendRuns(resultParagraph, field.Result, rtfDocument);
        IEnumerable<WordParagraph> runs = resultParagraph._paragraph.Elements<Run>()
            .Select(run => new WordParagraph(wordParagraph._document, resultParagraph._paragraph, run));

        string tooltip = field.HyperlinkField?.ScreenTip ?? string.Empty;
        WordHyperLink.AddHyperLink(wordParagraph, runs, field.Hyperlink!, addStyle: true, tooltip: tooltip);
    }

    private static void AppendBookmarkMarker(WordParagraph wordParagraph, RtfBookmarkMarker marker, Dictionary<string, string> openBookmarks) {
        if (marker.Kind == RtfBookmarkMarkerKind.Start) {
            string id = wordParagraph._document.BookmarkId.ToString();
            openBookmarks[marker.Name] = id;
            wordParagraph._paragraph.Append(new BookmarkStart { Name = marker.Name, Id = id });
            return;
        }

        if (openBookmarks.TryGetValue(marker.Name, out string? bookmarkId)) {
            wordParagraph._paragraph.Append(new BookmarkEnd { Id = bookmarkId });
            openBookmarks.Remove(marker.Name);
        }
    }

    private static void AppendNote(WordParagraph wordRun, RtfNote note, RtfDocument? rtfDocument) {
        if (note.Kind == RtfNoteKind.Annotation) {
            AppendAnnotationComment(wordRun, note, rtfDocument);
            return;
        }

        if (note.Kind != RtfNoteKind.Footnote && note.Kind != RtfNoteKind.Endnote) {
            return;
        }

        RtfParagraph? firstParagraph = note.Paragraphs.FirstOrDefault();
        WordParagraph noteParagraph = firstParagraph != null
            ? CreateDetachedWordParagraph(wordRun._document, firstParagraph, rtfDocument)
            : new WordParagraph(wordRun._document, newParagraph: true, newRun: true);
        WordParagraph referenceRun = note.Kind == RtfNoteKind.Endnote
            ? WordEndNote.AddEndNote(wordRun._document, wordRun, noteParagraph)
            : WordFootNote.AddFootNote(wordRun._document, wordRun, noteParagraph);
        AppendAdditionalNoteParagraphs(referenceRun, note, rtfDocument);
    }

    private static WordParagraph CreateDetachedWordParagraph(WordDocument document, RtfParagraph paragraph, RtfDocument? rtfDocument) {
        var wordParagraph = new WordParagraph(document, newParagraph: true, newRun: false);
        if (rtfDocument != null) {
            ApplyParagraphFormatting(wordParagraph, paragraph, rtfDocument);
        }

        AppendRuns(wordParagraph, paragraph, rtfDocument);
        return wordParagraph;
    }

    private static void AppendAdditionalNoteParagraphs(WordParagraph referenceRun, RtfNote note, RtfDocument? rtfDocument) {
        if (note.Paragraphs.Count <= 1) {
            return;
        }

        long? referenceId = note.Kind == RtfNoteKind.Endnote
            ? referenceRun.EndNote?.ReferenceId
            : referenceRun.FootNote?.ReferenceId;
        if (!referenceId.HasValue) {
            return;
        }

        OpenXmlCompositeElement? noteElement = note.Kind == RtfNoteKind.Endnote
            ? referenceRun._document._wordprocessingDocument.MainDocumentPart?.EndnotesPart?.Endnotes?
                .ChildElements.OfType<Endnote>()
                .FirstOrDefault(item => item.Id?.Value == referenceId.Value)
            : referenceRun._document._wordprocessingDocument.MainDocumentPart?.FootnotesPart?.Footnotes?
                .ChildElements.OfType<Footnote>()
                .FirstOrDefault(item => item.Id?.Value == referenceId.Value);
        if (noteElement == null) return;

        foreach (RtfParagraph paragraph in note.Paragraphs.Skip(1)) {
            WordParagraph wordParagraph = CreateDetachedWordParagraph(referenceRun._document, paragraph, rtfDocument);
            noteElement.Append(wordParagraph._paragraph);
        }
    }

    private sealed class ComplexFieldCapture {
        public System.Text.StringBuilder Instruction { get; } = new System.Text.StringBuilder();

        public RtfParagraph Result { get; } = new RtfParagraph();

        public bool CapturingResult { get; set; }

        public RtfRun? PreviousRun { get; set; }
    }
}
