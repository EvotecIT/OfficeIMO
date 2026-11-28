using System.Globalization;
using System.Text;

namespace OfficeIMO.Word.Rtf;

/// <summary>
/// Extension methods for converting between <see cref="WordDocument"/> and <see cref="RtfDocument"/>.
/// </summary>
public static partial class WordRtfConverterExtensions {
    internal static RtfDocument ConvertToRtfDocument(this WordDocument document) {
        ArgumentNullException.ThrowIfNull(document);

        RtfDocument rtf = RtfDocument.Create();
        CopyDocumentInfo(document, rtf);
        CopyCustomMetadata(document, rtf);
        CopyDefaultLanguage(document, rtf);
        CopyDocumentSettings(document, rtf);
        CopyWordStylesAndNumbering(document, rtf);
        var revisionAuthorIndexes = new Dictionary<string, int>(StringComparer.Ordinal);
        if (ShouldExportSections(document)) {
            CopySections(document, rtf, revisionAuthorIndexes);
        } else {
            CopyHeaderFooters(document, rtf, revisionAuthorIndexes);
            CopyPageSetup(document, rtf);
            CopyWordElements(document.Elements, rtf, revisionAuthorIndexes);
        }

        return rtf;
    }

    private static WordDocument ToWordDocumentCore(RtfDocument rtfDocument, RtfConversionReport? report) {
        RtfTableTraversalGuard.ValidateDocument(rtfDocument);
        WordDocument document = WordDocument.Create();
        ApplyDocumentInfo(rtfDocument, document);
        ApplyCustomMetadata(rtfDocument, document);
        ApplyDefaultLanguage(rtfDocument, document);
        ApplyDocumentSettings(rtfDocument, document);
        ApplyRtfStylesAndNumbering(rtfDocument, document);
        RtfWordNumberingPlan numbering = RtfWordNumberingPlan.Create(rtfDocument, document, report);
        if (rtfDocument.Sections.Count > 0) {
            ApplySections(rtfDocument, document, numbering);
        } else {
            ApplyHeaderFooters(rtfDocument, document, numbering);
            ApplyPageSetup(rtfDocument, document);
            foreach (IRtfBlock block in rtfDocument.Blocks) {
                if (block is RtfParagraph paragraph) {
                    AppendParagraph(document, paragraph, rtfDocument, numbering);
                } else if (block is RtfTable table) {
                    AppendTable(document, table, rtfDocument, numbering);
                } else if (block is RtfImage image) {
                    AppendImage(document, image);
                }
            }
        }

        return document;
    }

    private static void CopyHeaderFooters(WordDocument document, RtfDocument rtf, Dictionary<string, int> revisionAuthorIndexes) {
        CopyHeaderFooter(document.PageSetup.GetHeader(WordHeaderFooterType.Default), rtf, rtf.AddHeader, RtfHeaderFooterKind.Header, revisionAuthorIndexes);
        CopyHeaderFooter(document.PageSetup.GetHeader(WordHeaderFooterType.First), rtf, rtf.AddHeader, RtfHeaderFooterKind.FirstHeader, revisionAuthorIndexes);
        CopyHeaderFooter(document.PageSetup.GetHeader(WordHeaderFooterType.Even), rtf, rtf.AddHeader, RtfHeaderFooterKind.LeftHeader, revisionAuthorIndexes);
        CopyHeaderFooter(document.PageSetup.GetFooter(WordHeaderFooterType.Default), rtf, rtf.AddFooter, RtfHeaderFooterKind.Footer, revisionAuthorIndexes);
        CopyHeaderFooter(document.PageSetup.GetFooter(WordHeaderFooterType.First), rtf, rtf.AddFooter, RtfHeaderFooterKind.FirstFooter, revisionAuthorIndexes);
        CopyHeaderFooter(document.PageSetup.GetFooter(WordHeaderFooterType.Even), rtf, rtf.AddFooter, RtfHeaderFooterKind.LeftFooter, revisionAuthorIndexes);
    }

    private static void CopyHeaderFooter(WordHeaderFooter? source, RtfDocument rtf, Func<RtfHeaderFooterKind, RtfHeaderFooter> addDestination, RtfHeaderFooterKind kind, Dictionary<string, int> revisionAuthorIndexes) {
        if (source == null) {
            return;
        }

        RtfHeaderFooter destination = addDestination(kind);
        foreach (WordParagraph wordParagraph in source.Paragraphs.GroupBy(paragraph => paragraph.Node()).Select(group => group.First())) {
            RtfParagraph paragraph = destination.AddParagraph();
            CopyParagraphFormatting(wordParagraph, paragraph, rtf);
            AppendFormattedRuns(wordParagraph, paragraph, rtf, revisionAuthorIndexes);
        }
    }

    private static void CopyTabStops(WordParagraph source, RtfParagraph destination) {
        foreach (WordTabStop tabStop in source.TabStops) {
            // RTF has no corresponding clear-stop declaration in this model;
            // conversion diagnostics report the unsupported source control.
            if (tabStop.Alignment == WordTabAlignment.Clear) continue;
            destination.AddTabStop(checked((int)tabStop.Position.ToTwips()), ToRtfTabAlignment(tabStop.Alignment.ToOpenXmlToken()), ToRtfTabLeader(tabStop.Leader.ToOpenXmlToken()));
        }
    }

    // Tab alignments and leaders are w:val and w:leader tokens.
    private static RtfTabAlignment ToRtfTabAlignment(string alignment) {
        if (alignment == "center") return RtfTabAlignment.Center;
        if (alignment == "right") return RtfTabAlignment.Right;
        if (alignment == "decimal") return RtfTabAlignment.Decimal;
        if (alignment == "bar") return RtfTabAlignment.Bar;
        return RtfTabAlignment.Left;
    }

    private static RtfTabLeader ToRtfTabLeader(string leader) {
        if (leader == "dot") return RtfTabLeader.Dots;
        if (leader == "middleDot") return RtfTabLeader.MiddleDots;
        if (leader == "hyphen") return RtfTabLeader.Hyphen;
        if (leader == "underscore") return RtfTabLeader.Underline;
        if (leader == "heavy") return RtfTabLeader.ThickLine;
        return RtfTabLeader.None;
    }

    private static void ApplyHeaderFooters(RtfDocument rtfDocument, WordDocument document, RtfWordNumberingPlan numbering) {
        foreach (RtfHeaderFooter headerFooter in rtfDocument.HeaderFooters) {
            WordHeaderFooter target = GetWordHeaderFooter(document, headerFooter.Kind);
            foreach (RtfParagraph paragraph in headerFooter.Paragraphs) {
                AppendParagraph(target, paragraph, rtfDocument, numbering);
            }
        }
    }

    private static WordHeaderFooter GetWordHeaderFooter(WordDocument document, RtfHeaderFooterKind kind) {
        switch (kind) {
            case RtfHeaderFooterKind.FirstHeader:
                return document.PageSetup.FirstPageHeader;
            case RtfHeaderFooterKind.LeftHeader:
                return document.PageSetup.EvenPageHeader;
            case RtfHeaderFooterKind.RightHeader:
            case RtfHeaderFooterKind.Header:
                return document.PageSetup.Header;
            case RtfHeaderFooterKind.FirstFooter:
                return document.PageSetup.FirstPageFooter;
            case RtfHeaderFooterKind.LeftFooter:
                return document.PageSetup.EvenPageFooter;
            case RtfHeaderFooterKind.RightFooter:
            case RtfHeaderFooterKind.Footer:
                return document.PageSetup.Footer;
            default:
                throw new ArgumentOutOfRangeException(nameof(kind), kind, "Unsupported RTF header or footer kind.");
        }
    }

    private static void AppendFormattedRuns(WordParagraph wordParagraph, RtfParagraph paragraph, RtfDocument rtfDocument, Dictionary<string, int> revisionAuthorIndexes) {
        bool hasRuns = false;
        RtfRun? previousRun = null;
        var complexFields = new Stack<ComplexFieldCapture>();
        var bookmarkNames = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (OfficeXmlNode element in wordParagraph.Node().Elements()) {
            if (!element.IsModelElement()) continue;
            if (element.NamespaceId == WordXml.M) {
                if (element.LocalName is "oMath" or "oMathPara") {
                    AppendEquationField(paragraph, element, complexFields);
                    hasRuns = true;
                    previousRun = null;
                }
                continue;
            }
            if (element.NamespaceId != WordXml.W) continue;
            switch (element.LocalName) {
                case "bookmarkStart" when element.WAttr("name") != null && element.WAttr("id") != null:
                    bookmarkNames[element.WAttr("id")!] = element.WAttr("name")!;
                    paragraph.AddBookmarkStart(element.WAttr("name")!);
                    break;
                case "bookmarkEnd" when element.WAttr("id") != null && bookmarkNames.TryGetValue(element.WAttr("id")!, out string? bookmarkName):
                    paragraph.AddBookmarkEnd(bookmarkName);
                    break;
                case "r":
                    if (IsCommentReferenceOnlyRun(element)) {
                        break;
                    }

                    if (TryAppendComplexFieldRun(wordParagraph, element, paragraph, rtfDocument, revisionAuthorIndexes, complexFields)) {
                        previousRun = null;
                        hasRuns = true;
                        break;
                    }

                    hasRuns |= AppendWordRun(wordParagraph._document.RunView(wordParagraph.Node(), element), paragraph, ref previousRun, rtfDocument, revisionAuthorIndexes);
                    break;
                case "ins":
                case "moveTo":
                    hasRuns |= AppendRevisionContent(wordParagraph, element, paragraph, ref previousRun, rtfDocument, revisionAuthorIndexes, complexFields, RtfRevisionKind.Inserted, element.WAttr("author"));
                    break;
                case "del":
                case "moveFrom":
                    hasRuns |= AppendRevisionContent(wordParagraph, element, paragraph, ref previousRun, rtfDocument, revisionAuthorIndexes, complexFields, RtfRevisionKind.Deleted, element.WAttr("author"));
                    break;
                case "hyperlink":
                    hasRuns |= AppendHyperlinkContent(
                        wordParagraph,
                        element,
                        element,
                        paragraph,
                        ref previousRun,
                        rtfDocument,
                        revisionAuthorIndexes,
                        complexFields);
                    break;
                case "sdt":
                case "sdtContent":
                    hasRuns |= AppendInlineContainerContent(
                        wordParagraph,
                        element,
                        paragraph,
                        ref previousRun,
                        rtfDocument,
                        revisionAuthorIndexes,
                        complexFields);
                    break;
                case "fldSimple":
                    AppendSimpleField(
                        wordParagraph,
                        element,
                        paragraph,
                        rtfDocument,
                        revisionAuthorIndexes,
                        complexFields);
                    hasRuns = true;
                    break;
            }
        }

        if (!hasRuns && !string.IsNullOrEmpty(wordParagraph.Text)) {
            paragraph.AddText(wordParagraph.Text);
        }

        AppendWordComments(wordParagraph, paragraph, rtfDocument, revisionAuthorIndexes);
    }

    private static bool AppendHyperlinkContent(
        WordParagraph wordParagraph,
        OfficeXmlNode container,
        OfficeXmlNode hyperlink,
        RtfParagraph paragraph,
        ref RtfRun? previousRun,
        RtfDocument rtfDocument,
        Dictionary<string, int> revisionAuthorIndexes,
        Stack<ComplexFieldCapture> complexFields,
        RtfRevisionKind revisionKind = RtfRevisionKind.None,
        int? revisionAuthorIndex = null) {
        ComplexFieldCapture? activeCapture = complexFields.Count > 0 ? complexFields.Peek() : null;
        if (activeCapture != null && !activeCapture.CapturingResult) {
            return false;
        }

        RtfParagraph destination = activeCapture?.Result ?? paragraph;
        RtfField hyperlinkField = destination.AddField(CreateHyperlinkFieldInstruction(wordParagraph, hyperlink));
        RtfRun? hyperlinkPreviousRun = null;
        bool hasContent = AppendInlineContainerContent(
            wordParagraph,
            container,
            hyperlinkField.Result,
            ref hyperlinkPreviousRun,
            rtfDocument,
            revisionAuthorIndexes,
            new Stack<ComplexFieldCapture>(),
            revisionKind,
            revisionAuthorIndex);
        if (activeCapture != null) {
            activeCapture.PreviousRun = null;
        }
        previousRun = null;
        return hasContent;
    }

    private static bool AppendInlineContainerContent(
        WordParagraph wordParagraph,
        OfficeXmlNode container,
        RtfParagraph paragraph,
        ref RtfRun? previousRun,
        RtfDocument rtfDocument,
        Dictionary<string, int> revisionAuthorIndexes,
        Stack<ComplexFieldCapture> complexFields,
        RtfRevisionKind revisionKind = RtfRevisionKind.None,
        int? revisionAuthorIndex = null) {
        bool hasContent = false;
        foreach (OfficeXmlNode child in container.Elements()) {
            if (!child.IsModelElement()) continue;
            if (child.NamespaceId == WordXml.M) {
                if (child.LocalName is "oMath" or "oMathPara") {
                    AppendEquationField(paragraph, child, complexFields, revisionKind, revisionAuthorIndex);
                    previousRun = null;
                    hasContent = true;
                }
                continue;
            }
            if (child.NamespaceId != WordXml.W) continue;
            switch (child.LocalName) {
                case "r":
                    if (TryAppendComplexFieldRun(
                        wordParagraph,
                        child,
                        paragraph,
                        rtfDocument,
                        revisionAuthorIndexes,
                        complexFields,
                        revisionKind,
                        revisionAuthorIndex)) {
                        previousRun = null;
                        hasContent = true;
                        break;
                    }

                    hasContent |= AppendWordRun(
                        wordParagraph._document.RunView(wordParagraph.Node(), child),
                        paragraph,
                        ref previousRun,
                        rtfDocument,
                        revisionAuthorIndexes,
                        revisionKind,
                        revisionAuthorIndex);
                    break;
                case "fldSimple":
                    AppendSimpleField(
                        wordParagraph,
                        child,
                        paragraph,
                        rtfDocument,
                        revisionAuthorIndexes,
                        complexFields,
                        revisionKind,
                        revisionAuthorIndex);
                    previousRun = null;
                    hasContent = true;
                    break;
                case "del":
                case "moveFrom":
                    break;
                case "ins":
                case "moveTo":
                    hasContent |= AppendRevisionContent(
                        wordParagraph,
                        child,
                        paragraph,
                        ref previousRun,
                        rtfDocument,
                        revisionAuthorIndexes,
                        complexFields,
                        RtfRevisionKind.Inserted,
                        child.WAttr("author"));
                    break;
                case "sdt":
                case "sdtContent":
                    hasContent |= AppendInlineContainerContent(
                        wordParagraph,
                        child,
                        paragraph,
                        ref previousRun,
                        rtfDocument,
                        revisionAuthorIndexes,
                        complexFields,
                        revisionKind,
                        revisionAuthorIndex);
                    break;
                case "hyperlink":
                    hasContent |= AppendHyperlinkContent(
                        wordParagraph,
                        child,
                        child,
                        paragraph,
                        ref previousRun,
                        rtfDocument,
                        revisionAuthorIndexes,
                        complexFields,
                        revisionKind,
                        revisionAuthorIndex);
                    break;
            }
        }

        return hasContent;
    }

    private static string CreateHyperlinkFieldInstruction(WordParagraph wordParagraph, OfficeXmlNode hyperlink) {
        WordHyperLink wordHyperlink = wordParagraph._document.HyperlinkView(wordParagraph.Node(), hyperlink);
        return new RtfHyperlinkFieldInfo {
            Target = wordHyperlink.Uri,
            SubAddress = hyperlink.WAttr("anchor"),
            ScreenTip = hyperlink.WAttr("tooltip"),
            TargetFrame = hyperlink.WAttr("tgtFrame")
        }.ToInstruction();
    }

    private static bool AppendRevisionContent(
        WordParagraph wordParagraph,
        OfficeXmlNode revision,
        RtfParagraph paragraph,
        ref RtfRun? previousRun,
        RtfDocument rtfDocument,
        Dictionary<string, int> revisionAuthorIndexes,
        Stack<ComplexFieldCapture> complexFields,
        RtfRevisionKind revisionKind,
        string? author) {
        bool hasContent = false;
        int? authorIndex = GetOrAddRevisionAuthorIndex(rtfDocument, author, revisionAuthorIndexes);
        foreach (OfficeXmlNode child in revision.Elements()) {
            if (!child.IsModelElement()) continue;
            if (child.NamespaceId == WordXml.M) {
                if (child.LocalName is "oMath" or "oMathPara") {
                    AppendEquationField(paragraph, child, complexFields, revisionKind, authorIndex);
                    previousRun = null;
                    hasContent = true;
                }
                continue;
            }
            if (child.NamespaceId != WordXml.W) continue;
            switch (child.LocalName) {
                case "r":
                    if (TryAppendComplexFieldRun(
                        wordParagraph,
                        child,
                        paragraph,
                        rtfDocument,
                        revisionAuthorIndexes,
                        complexFields,
                        revisionKind,
                        authorIndex)) {
                        previousRun = null;
                        hasContent = true;
                        break;
                    }

                    hasContent |= AppendWordRun(
                        wordParagraph._document.RunView(wordParagraph.Node(), child),
                        paragraph,
                        ref previousRun,
                        rtfDocument,
                        revisionAuthorIndexes,
                        revisionKind,
                        authorIndex);
                    break;
                case "fldSimple":
                    AppendSimpleField(
                        wordParagraph,
                        child,
                        paragraph,
                        rtfDocument,
                        revisionAuthorIndexes,
                        complexFields,
                        revisionKind,
                        authorIndex);
                    previousRun = null;
                    hasContent = true;
                    break;
                case "hyperlink":
                    hasContent |= AppendHyperlinkContent(
                        wordParagraph,
                        child,
                        child,
                        paragraph,
                        ref previousRun,
                        rtfDocument,
                        revisionAuthorIndexes,
                        complexFields,
                        revisionKind,
                        authorIndex);
                    break;
                case "sdt":
                case "sdtContent":
                    hasContent |= AppendInlineContainerContent(
                        wordParagraph,
                        child,
                        paragraph,
                        ref previousRun,
                        rtfDocument,
                        revisionAuthorIndexes,
                        complexFields,
                        revisionKind,
                        authorIndex);
                    break;
            }
        }

        return hasContent;
    }

    private static bool IsCommentReferenceOnlyRun(OfficeXmlNode run) =>
        WordConverterNodes.ChildNodesIncludingMarkup(run).All(element =>
            element.IsModelElement() && (element.IsW("rPr") || element.IsW("commentReference")));

    private static bool TryAppendComplexFieldRun(
        WordParagraph wordParagraph,
        OfficeXmlNode runElement,
        RtfParagraph paragraph,
        RtfDocument rtfDocument,
        Dictionary<string, int> revisionAuthorIndexes,
        Stack<ComplexFieldCapture> captures,
        RtfRevisionKind revisionKind = RtfRevisionKind.None,
        int? revisionAuthorIndex = null) {
        OfficeXmlNode fieldChar = runElement.ModelChild("fldChar");
        string? fieldCharType = fieldChar.WAttr("fldCharType");
        if (fieldCharType == "begin") {
            captures.Push(new ComplexFieldCapture(revisionKind != RtfRevisionKind.None));
            return true;
        }

        if (captures.Count == 0) {
            return false;
        }

        ComplexFieldCapture capture = captures.Peek();
        OfficeXmlNode fieldCode = runElement.ModelChild("instrText");
        if (revisionKind != RtfRevisionKind.None && (fieldChar.Exists || fieldCode.Exists)) {
            capture.RestrictToEquationFields = true;
        }
        if (fieldCharType == "separate") {
            capture.CapturingResult = true;
            return true;
        }

        if (fieldCharType == "end") {
            captures.Pop();
            if (captures.Count == 0) {
                CompleteComplexField(paragraph, capture);
            } else {
                CompleteNestedComplexField(captures.Peek(), capture);
            }

            return true;
        }

        if (fieldCode.Exists && !capture.CapturingResult) {
            capture.Instruction.Append(fieldCode.LeafText());
            return true;
        }

        if (capture.CapturingResult) {
            RtfRun? previousRun = capture.PreviousRun;
            AppendWordRun(
                wordParagraph._document.RunView(wordParagraph.Node(), runElement),
                capture.Result,
                ref previousRun,
                rtfDocument,
                revisionAuthorIndexes,
                revisionKind,
                revisionAuthorIndex);
            capture.PreviousRun = previousRun;
            return true;
        }

        return true;
    }

    private static void CompleteComplexField(RtfParagraph paragraph, ComplexFieldCapture capture) {
        if (capture.RestrictToEquationFields && !IsEquationFieldInstruction(capture.Instruction.ToString())) {
            CopyInlines(capture.Result, paragraph);
            return;
        }

        RtfField field = paragraph.AddField(capture.Instruction.ToString().Trim());
        CopyInlines(capture.Result, field.Result);
    }

    private static void CompleteNestedComplexField(ComplexFieldCapture parent, ComplexFieldCapture nested) {
        if (nested.RestrictToEquationFields && !IsEquationFieldInstruction(nested.Instruction.ToString())) {
            if (parent.CapturingResult) {
                CopyInlines(nested.Result, parent.Result);
                parent.PreviousRun = null;
            }
            return;
        }

        if (!parent.CapturingResult) {
            parent.Instruction.Append(nested.Instruction.ToString().Trim());
            return;
        }

        RtfField field = parent.Result.AddField(nested.Instruction.ToString().Trim());
        CopyInlines(nested.Result, field.Result);
        parent.PreviousRun = null;
    }

    private static bool IsEquationFieldInstruction(string instruction) {
        string trimmed = (instruction ?? string.Empty).TrimStart();
        if (!trimmed.StartsWith("EQ", StringComparison.OrdinalIgnoreCase)) return false;
        return trimmed.Length == 2 || char.IsWhiteSpace(trimmed[2]) || trimmed[2] == '\\';
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
        destination.Strikethrough = source.Strikethrough;
        destination.DoubleStrikethrough = source.DoubleStrikethrough;
        destination.DirectHidden = source.DirectHidden;
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

    private static void AppendSimpleField(
        WordParagraph wordParagraph,
        OfficeXmlNode simpleField,
        RtfParagraph paragraph,
        RtfDocument rtfDocument,
        Dictionary<string, int> revisionAuthorIndexes,
        Stack<ComplexFieldCapture> complexFields,
        RtfRevisionKind revisionKind = RtfRevisionKind.None,
        int? revisionAuthorIndex = null) {
        ComplexFieldCapture? activeCapture = complexFields.Count > 0 ? complexFields.Peek() : null;
        if (activeCapture != null && !activeCapture.CapturingResult) {
            return;
        }

        string instruction = simpleField.WAttr("instr") ?? string.Empty;
        RtfParagraph destination = activeCapture?.Result ?? paragraph;
        if (revisionKind != RtfRevisionKind.None && !IsEquationFieldInstruction(instruction)) {
            RtfRun? flattenedPreviousRun = null;
            AppendInlineContainerContent(
                wordParagraph,
                simpleField,
                destination,
                ref flattenedPreviousRun,
                rtfDocument,
                revisionAuthorIndexes,
                new Stack<ComplexFieldCapture>(),
                revisionKind,
                revisionAuthorIndex);
            if (activeCapture != null) {
                activeCapture.PreviousRun = flattenedPreviousRun;
            }
            return;
        }

        RtfField field = destination.AddField(instruction.Trim());
        RtfRun? previousRun = null;
        AppendInlineContainerContent(
            wordParagraph,
            simpleField,
            field.Result,
            ref previousRun,
            rtfDocument,
            revisionAuthorIndexes,
            new Stack<ComplexFieldCapture>(),
            revisionKind,
            revisionAuthorIndex);
        if (activeCapture != null) {
            activeCapture.PreviousRun = null;
        }
    }

    private static void AppendEquationField(
        RtfParagraph paragraph,
        OfficeXmlNode mathElement,
        Stack<ComplexFieldCapture> captures,
        RtfRevisionKind revisionKind = RtfRevisionKind.None,
        int? revisionAuthorIndex = null) {
        if (captures.Count == 0) {
            AppendEquationField(paragraph, mathElement, revisionKind, revisionAuthorIndex);
            return;
        }

        ComplexFieldCapture capture = captures.Peek();
        if (!capture.CapturingResult) {
            return;
        }

        AppendEquationField(capture.Result, mathElement, revisionKind, revisionAuthorIndex);
        capture.PreviousRun = null;
    }

    private static void AppendEquationField(
        RtfParagraph paragraph,
        OfficeXmlNode mathElement,
        RtfRevisionKind revisionKind = RtfRevisionKind.None,
        int? revisionAuthorIndex = null) {
        RtfField field = paragraph.AddEquationField(
            WordConverterNodes.MathEquationFieldInstruction(mathElement),
            WordConverterNodes.MathText(mathElement));
        foreach (RtfRun resultRun in field.Result.Inlines.OfType<RtfRun>()) {
            resultRun.RevisionKind = revisionKind;
            resultRun.RevisionAuthorIndex = revisionAuthorIndex;
        }
    }

    private static bool AppendWordRunText(WordParagraph wordRun, RtfParagraph paragraph, ref RtfRun? previousRun, RtfDocument rtfDocument, RtfRevisionKind revisionKind, int? revisionAuthorIndex) {
        string? text = GetWordRunText(wordRun, revisionKind);
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

    private static string GetWordRunText(WordParagraph wordRun, RtfRevisionKind revisionKind) {
        if (revisionKind != RtfRevisionKind.Deleted || wordRun.RunNode().IsNone) {
            return wordRun.Text;
        }

        var builder = new StringBuilder();
        foreach (OfficeXmlNode child in wordRun.ProjectedRunChildNodes()) {
            if (!child.IsModelElement()) continue;
            if (child.IsW("t") || child.IsW("delText")) {
                builder.Append(child.LeafText());
            } else if (child.IsW("tab")) {
                builder.Append('\t');
            } else if (child.IsW("br")) {
                string? type = child.WAttr("type");
                builder.Append(type is null || type == "textWrapping" ? '\n' : '\u2028');
            }
        }

        return builder.ToString();
    }

    // Break kinds are w:type tokens.
    private static RtfBreakKind ToRtfBreakKind(string? breakType) {
        if (breakType == "page") return RtfBreakKind.Page;
        if (breakType == "column") return RtfBreakKind.Column;
        return RtfBreakKind.Line;
    }

    private static RtfNote CopyNote(List<WordParagraph>? paragraphs, RtfNoteKind kind, RtfDocument rtfDocument, Dictionary<string, int> revisionAuthorIndexes) {
        var note = new RtfNote(kind);
        if (paragraphs == null) {
            return note;
        }

        foreach (WordParagraph wordParagraph in paragraphs.GroupBy(paragraph => paragraph.Node()).Select(group => group.First())) {
            RtfParagraph paragraph = note.AddParagraph();
            CopyParagraphFormatting(wordParagraph, paragraph, rtfDocument);
            AppendFormattedRuns(wordParagraph, paragraph, rtfDocument, revisionAuthorIndexes);
        }

        return note;
    }

    private static void AppendParagraph(WordDocument document, RtfParagraph paragraph, RtfDocument rtfDocument, RtfWordNumberingPlan numbering) {
        WordParagraph wordParagraph = document.AddParagraph();
        ApplyParagraphFormatting(wordParagraph, paragraph, rtfDocument, numbering);
        AppendRuns(wordParagraph, paragraph, rtfDocument, numbering);
    }

    private static void AppendParagraph(WordHeaderFooter headerFooter, RtfParagraph paragraph, RtfDocument rtfDocument, RtfWordNumberingPlan numbering) {
        WordParagraph wordParagraph = headerFooter.AddParagraph();
        ApplyParagraphFormatting(wordParagraph, paragraph, rtfDocument, numbering);
        AppendRuns(wordParagraph, paragraph, rtfDocument, numbering);
    }

    private static void ApplyTabStops(WordParagraph destination, RtfParagraph source, RtfDocument document) {
        if (source.StyleId.HasValue && HasRtfStyle(document, source.StyleId.Value, RtfStyleKind.Paragraph)) {
            // Word merges tabs along the retained style chain, while an RTF
            // tab list replaces that chain's list. Inventory declarations, not
            // only the effective nearest style, so an ancestor cannot add ink
            // or move a tab absent from the effective RTF paragraph.
            var activePositions = new HashSet<int>(source.TabStops.Select(stop => stop.PositionTwips));
            var visitedStyles = new HashSet<int>();
            var clearedPositions = new HashSet<int>();
            int? styleId = source.StyleId;
            while (styleId.HasValue && visitedStyles.Add(styleId.Value)) {
                RtfStyle? style = document.Styles.LastOrDefault(item =>
                    item.Id == styleId.Value && item.Kind == RtfStyleKind.Paragraph);
                if (style == null) break;
                foreach (RtfTabStop tabStop in style.TabStops) {
                    int position = tabStop.PositionTwips;
                    if (!activePositions.Contains(position) && clearedPositions.Add(position)) {
                        destination.AddTabStop(OfficeLength.Emus((long)position * 635), WordTabAlignment.Clear);
                    }
                }
                styleId = style.BasedOnStyleId;
            }
        }
        foreach (RtfTabStop tabStop in source.TabStops) {
            destination.AddTabStop(OfficeLength.Emus((long)tabStop.PositionTwips * 635), ToWordTabAlignment(tabStop.Alignment), ToWordTabLeader(tabStop.Leader));
        }
    }

    private static WordTabAlignment ToWordTabAlignment(RtfTabAlignment alignment) {
        switch (alignment) {
            case RtfTabAlignment.Center:
                return WordTabAlignment.Center;
            case RtfTabAlignment.Right:
                return WordTabAlignment.Right;
            case RtfTabAlignment.Decimal:
                return WordTabAlignment.Decimal;
            case RtfTabAlignment.Bar:
                return WordTabAlignment.Bar;
            default:
                return WordTabAlignment.Left;
        }
    }

    private static WordTabLeader ToWordTabLeader(RtfTabLeader leader) {
        switch (leader) {
            case RtfTabLeader.Dots:
                return WordTabLeader.Dot;
            case RtfTabLeader.MiddleDots:
                return WordTabLeader.MiddleDot;
            case RtfTabLeader.Hyphen:
                return WordTabLeader.Hyphen;
            case RtfTabLeader.Underline:
                return WordTabLeader.Underscore;
            case RtfTabLeader.ThickLine:
                return WordTabLeader.Heavy;
            case RtfTabLeader.EqualSign:
                return WordTabLeader.MiddleDot;
            default:
                return WordTabLeader.None;
        }
    }

    private static void AppendRuns(WordParagraph wordParagraph, RtfParagraph paragraph, RtfDocument? rtfDocument, RtfWordNumberingPlan numbering) {
        numbering.AppendFallback(wordParagraph, paragraph);
        var openBookmarks = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (IRtfInline inline in paragraph.Inlines) {
            if (inline is RtfBookmarkMarker marker) {
                AppendBookmarkMarker(wordParagraph, marker, openBookmarks);
                continue;
            }

            if (inline is RtfField field) {
                AppendField(wordParagraph, field, rtfDocument, numbering);
                continue;
            }

            if (inline is RtfGeneratedText generatedText) {
                AppendGeneratedText(wordParagraph, generatedText, rtfDocument, numbering);
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
                wordRun = wordParagraph.AddHyperLink(run.Text, run.Hyperlink, addStyle: true).Paragraph;
            } else {
                wordRun = wordParagraph.AddText(run.Text);
            }

            ApplyRtfRunFormatting(rtfDocument != null && rtfDocument.Styles.Count > 0 ? rtfDocument.GetRunFormatting(paragraph, run) : run, wordRun, rtfDocument);

            if (run.Note != null) {
                AppendNote(wordRun, run.Note, rtfDocument, numbering);
            }
        }
    }

    private static void AppendGeneratedText(WordParagraph wordParagraph, RtfGeneratedText generatedText, RtfDocument? rtfDocument, RtfWordNumberingPlan numbering) {
        if (generatedText.Kind == RtfGeneratedTextKind.NoteReference) {
            WordParagraph wordRun = wordParagraph.AddText(generatedText.FallbackText ?? string.Empty);
            if (generatedText.Note != null) {
                AppendNote(wordRun, generatedText.Note, rtfDocument, numbering);
            }

            return;
        }

        var field = new RtfField(ToWordFieldInstruction(generatedText.Kind));
        if (!string.IsNullOrEmpty(generatedText.FallbackText)) {
            field.AddText(generatedText.FallbackText!);
        }

        AppendField(wordParagraph, field, rtfDocument, numbering);
        if (generatedText.Note != null) {
            WordParagraph wordRun = wordParagraph.AddText(generatedText.FallbackText ?? string.Empty);
            AppendNote(wordRun, generatedText.Note, rtfDocument, numbering);
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

    private static WordBreakType? ToWordBreakKind(RtfBreakKind kind) {
        switch (kind) {
            case RtfBreakKind.Page:
            case RtfBreakKind.SoftPage:
                return WordBreakType.Page;
            case RtfBreakKind.Column:
                return WordBreakType.Column;
            default:
                return null;
        }
    }

    private static void SetHiddenWordRun(WordParagraph wordRun, bool hidden) {
        if (wordRun.RunNode().IsNone) {
            return;
        }

        WordXmlEdits.ReplaceChild(GetOrAddRunProperties(wordRun), "vanish", WRPrView.Order, FormatOnOff(hidden));
    }

    private static void AppendField(WordParagraph wordParagraph, RtfField field, RtfDocument? rtfDocument, RtfWordNumberingPlan numbering) {
        if (field.HyperlinkField != null) {
            AppendHyperlinkField(wordParagraph, field, rtfDocument, numbering);
            return;
        }

        OfficeXmlNode paragraph = wordParagraph.Node();
        OfficeXmlNode simpleField = paragraph.CreateElement(WordXml.W, "fldSimple");
        simpleField.SetAttribute(WordXml.W, "instr", field.Instruction);
        var resultParagraph = new WordParagraph(wordParagraph._document, newParagraph: true, newRun: false);
        AppendRuns(resultParagraph, field.Result, rtfDocument, numbering);
        foreach (OfficeXmlNode child in resultParagraph.Node().Elements()) {
            if (child.IsW("pPr")) {
                continue;
            }
            simpleField.Append(child.Clone());
        }

        paragraph.Append(simpleField);
    }

    private static void AppendHyperlinkField(WordParagraph wordParagraph, RtfField field, RtfDocument? rtfDocument, RtfWordNumberingPlan numbering) {
        var resultParagraph = new WordParagraph(wordParagraph._document, newParagraph: true, newRun: false);
        AppendRuns(resultParagraph, field.Result, rtfDocument, numbering);
        List<OfficeXmlNode> inlineContent = CreateHyperlinkInlineContent(resultParagraph.Node());
        string tooltip = field.HyperlinkField?.ScreenTip ?? string.Empty;
        WordParagraph? linkedParagraph = null;
        Uri? hyperlinkTarget = field.Hyperlink;
        if (hyperlinkTarget != null) {
            linkedParagraph = WordConverterNodes.AddHyperLinkContent(
                wordParagraph,
                inlineContent,
                hyperlinkTarget,
                addStyle: true,
                tooltip: tooltip);
        } else if (!string.IsNullOrEmpty(field.HyperlinkField?.SubAddress)) {
            linkedParagraph = WordConverterNodes.AddHyperLinkContent(
                wordParagraph,
                inlineContent,
                field.HyperlinkField!.SubAddress!,
                addStyle: true,
                tooltip: tooltip);
        } else {
            OfficeXmlNode paragraph = wordParagraph.Node();
            foreach (OfficeXmlNode child in inlineContent) {
                paragraph.Append(child.Clone());
            }
            return;
        }

        WordHyperLink? linkedHyperlink = linkedParagraph?.Hyperlink;
        if (linkedHyperlink != null && !string.IsNullOrEmpty(field.HyperlinkField?.SubAddress)) {
            linkedHyperlink.Anchor = field.HyperlinkField!.SubAddress;
        }
        string? targetFrameValue = field.HyperlinkField?.TargetFrame;
        if (linkedHyperlink != null &&
            !string.IsNullOrEmpty(targetFrameValue) &&
            Enum.TryParse<WordHyperlinkTargetFrame>(targetFrameValue, true, out WordHyperlinkTargetFrame targetFrame)) {
            linkedHyperlink.TargetFrame = targetFrame;
        }
    }

    private static List<OfficeXmlNode> CreateHyperlinkInlineContent(OfficeXmlNode paragraph) {
        var inlineContent = new List<OfficeXmlNode>();
        foreach (OfficeXmlNode child in paragraph.Elements()) {
            if (child.IsW("pPr")) {
                continue;
            }
            inlineContent.Add(child);
        }
        return inlineContent;
    }

    private static void AppendBookmarkMarker(WordParagraph wordParagraph, RtfBookmarkMarker marker, Dictionary<string, string> openBookmarks) {
        if (marker.Kind == RtfBookmarkMarkerKind.Start) {
            string id = wordParagraph._document.BookmarkId.ToString();
            openBookmarks[marker.Name] = id;
            OfficeXmlNode paragraph = wordParagraph.Node();
            OfficeXmlNode bookmarkStart = paragraph.CreateElement(WordXml.W, "bookmarkStart");
            bookmarkStart.SetAttribute(WordXml.W, "name", marker.Name);
            bookmarkStart.SetAttribute(WordXml.W, "id", id);
            paragraph.Append(bookmarkStart);
            return;
        }

        if (openBookmarks.TryGetValue(marker.Name, out string? bookmarkId)) {
            OfficeXmlNode paragraph = wordParagraph.Node();
            OfficeXmlNode bookmarkEnd = paragraph.CreateElement(WordXml.W, "bookmarkEnd");
            bookmarkEnd.SetAttribute(WordXml.W, "id", bookmarkId);
            paragraph.Append(bookmarkEnd);
            openBookmarks.Remove(marker.Name);
        }
    }

    /// <summary>The run's own <c>w:rPr</c>, or <see cref="OfficeXmlNode.None"/>.</summary>
    private static OfficeXmlNode DirectRunProperties(WordParagraph run) => run.RunNode().ModelChild("rPr");

    /// <summary>The run's <c>w:rPr</c>, added first among its children when missing.</summary>
    private static OfficeXmlNode GetOrAddRunProperties(WordParagraph run) =>
        run.RunNode().GetOrAddChild(WordXml.W, "rPr", WRView.Order);

    /// <summary>The paragraph's <c>w:pPr</c>, added first among its children when missing.</summary>
    private static OfficeXmlNode GetOrAddParagraphProperties(WordParagraph paragraph) =>
        paragraph.Node().GetOrAddChild(WordXml.W, "pPr", WPView.Order);

    /// <summary>Appends a new child holding <paramref name="value"/> in <c>w:val</c>.</summary>
    private static void AppendValueChild(OfficeXmlNode parent, string localName, string value) {
        OfficeXmlNode child = parent.CreateElement(WordXml.W, localName);
        child.SetAttribute(WordXml.W, "val", value);
        parent.Append(child);
    }

    private sealed class ComplexFieldCapture {
        internal ComplexFieldCapture(bool restrictToEquationFields = false) {
            RestrictToEquationFields = restrictToEquationFields;
        }

        public bool RestrictToEquationFields { get; set; }

        public System.Text.StringBuilder Instruction { get; } = new System.Text.StringBuilder();

        public RtfParagraph Result { get; } = new RtfParagraph();

        public bool CapturingResult { get; set; }

        public RtfRun? PreviousRun { get; set; }
    }
}
