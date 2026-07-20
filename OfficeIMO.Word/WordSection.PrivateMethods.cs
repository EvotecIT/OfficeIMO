using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Diagnostics;
using System.Threading;
using Ovml = DocumentFormat.OpenXml.Vml.Office;

namespace OfficeIMO.Word {
    /// <summary>
    /// Section in WordDocument
    /// </summary>
    public partial class WordSection {
        /// <summary>
        /// Converts tables to WordTables
        /// </summary>
        /// <param name="document"></param>
        /// <param name="tables"></param>
        /// <returns></returns>
        internal static List<WordTable> ConvertTableToWordTable(WordDocument document, IEnumerable<Table> tables) {
            var list = new List<WordTable>();
            foreach (Table table in tables) {
                list.Add(new WordTable(document, table));
            }
            return list;
        }

        /// <summary>
        /// Converts SdtBlock to WordWatermark if it's a watermark
        /// Hopefully detection is good enough, but it's possible that it will catch other things as well
        /// </summary>
        /// <param name="document"></param>
        /// <param name="sdtBlock"></param>
        /// <returns></returns>
        internal static List<WordWatermark> ConvertStdBlockToWatermark(WordDocument document, IEnumerable<SdtBlock> sdtBlock) {
            var list = new List<WordWatermark>();
            foreach (SdtBlock block in sdtBlock) {
                var watermark = ConvertStdBlockToWatermark(document, block);
                if (watermark != null) {
                    list.Add(watermark);
                }
            }
            return list;
        }

        /// <summary>
        /// Converts SdtBlock to WordWatermark if it's a watermark
        /// </summary>
        /// <param name="document"></param>
        /// <param name="sdtBlock"></param>
        /// <returns></returns>
        internal static WordWatermark? ConvertStdBlockToWatermark(WordDocument document, SdtBlock? sdtBlock) {
            if (sdtBlock == null) {
                return null;
            }
            var sdtContent = sdtBlock.SdtContentBlock;
            if (sdtContent == null) {
                return null;
            }
            var paragraphs = sdtContent.ChildElements.OfType<Paragraph>().FirstOrDefault();
            if (paragraphs == null) {
                return null;
            }
            var run = paragraphs.ChildElements.OfType<Run>().FirstOrDefault();
            if (run == null) {
                return null;
            }
            var picture = run.ChildElements.OfType<Picture>().FirstOrDefault();
            if (picture == null) {
                return null;
            }
            return new WordWatermark(document, sdtBlock);
        }

        /// <summary>
        /// Converts StdBlock to WordCoverPage
        /// </summary>
        /// <param name="document"></param>
        /// <param name="sdtBlocks"></param>
        /// <returns></returns>
        internal static WordCoverPage? ConvertStdBlockToCoverPage(WordDocument document, IEnumerable<SdtBlock?> sdtBlocks) {
            foreach (var sdtBlock in sdtBlocks) {
                if (sdtBlock == null) {
                    continue;
                }
                var sdtProperties = sdtBlock.ChildElements.OfType<SdtProperties>().FirstOrDefault();
                var docPartObject = sdtProperties?.ChildElements.OfType<SdtContentDocPartObject>().FirstOrDefault();
                var docPartGallery = docPartObject?.ChildElements.OfType<DocPartGallery>().FirstOrDefault();

                if (docPartGallery != null && docPartGallery.Val == "Cover Pages") {
                    return new WordCoverPage(document, sdtBlock!);
                }
            }

            return null;
        }

        /// <summary>
        /// Converts StdBlock to WordTableOfContent
        /// </summary>
        /// <param name="document"></param>
        /// <param name="sdtBlocks"></param>
        /// <returns></returns>
        internal static WordTableOfContent? ConvertStdBlockToTableOfContent(WordDocument document, IEnumerable<SdtBlock?> sdtBlocks) {
            foreach (var sdtBlock in sdtBlocks) {
                if (sdtBlock == null) {
                    continue;
                }
                var sdtProperties = sdtBlock.ChildElements.OfType<SdtProperties>().FirstOrDefault();
                var docPartObject = sdtProperties?.ChildElements.OfType<SdtContentDocPartObject>().FirstOrDefault();
                var docPartGallery = docPartObject?.ChildElements.OfType<DocPartGallery>().FirstOrDefault();

                if (docPartGallery != null && docPartGallery.Val == "Table of Contents") {
                    return new WordTableOfContent(document, sdtBlock!, queueUpdateOnOpen: false);
                }
            }
            return ConvertRawSimpleFieldToTableOfContent(document) ?? ConvertRawComplexFieldToTableOfContent(document);
        }

        private static WordTableOfContent? ConvertRawSimpleFieldToTableOfContent(WordDocument document) {
            Body? body = document._document.Body;
            if (body == null) {
                return null;
            }

            foreach (Paragraph paragraph in body.Elements<Paragraph>().ToList()) {
                int? fieldChildIndex = FindSimpleTocOrIndexFieldChildIndex(paragraph);
                if (fieldChildIndex == null) {
                    continue;
                }

                SdtBlock sdtBlock = CreateImportedTableOfContentBlock();
                SdtContentBlock content = sdtBlock.SdtContentBlock!;
                int insertIndex = body.ChildElements.ToList().IndexOf(paragraph);

                Paragraph? fieldStartPrefix = SplitFieldStartPrefix(paragraph, fieldChildIndex.Value);
                if (fieldStartPrefix != null) {
                    fieldChildIndex = FindSimpleTocOrIndexFieldChildIndex(paragraph);
                    if (fieldChildIndex == null) {
                        continue;
                    }
                }

                Paragraph? fieldPrefix = SplitSimpleFieldPrefix(paragraph, fieldChildIndex.Value);
                if (fieldPrefix != null) {
                    content.Append(fieldPrefix);
                } else {
                    paragraph.Remove();
                    content.Append(paragraph);
                }

                document.AssignNewSdtIds(sdtBlock);
                if (insertIndex >= 0 && insertIndex <= body.ChildElements.Count) {
                    if (fieldStartPrefix != null) {
                        body.InsertAt(fieldStartPrefix, insertIndex);
                        insertIndex++;
                    }

                    body.InsertAt(sdtBlock, insertIndex);
                } else {
                    if (fieldStartPrefix != null) {
                        body.Append(fieldStartPrefix);
                    }

                    body.Append(sdtBlock);
                }

                return new WordTableOfContent(document, sdtBlock, queueUpdateOnOpen: false);
            }

            return null;
        }

        private static WordTableOfContent? ConvertRawComplexFieldToTableOfContent(WordDocument document) {
            Body? body = document._document.Body;
            if (body == null) {
                return null;
            }

            List<Paragraph> paragraphs = body.Elements<Paragraph>().ToList();
            for (int index = 0; index < paragraphs.Count; index++) {
                Paragraph paragraph = paragraphs[index];
                int? fieldStartChildIndex = FindTocOrIndexComplexFieldStart(paragraph);
                if (fieldStartChildIndex == null) {
                    continue;
                }

                ComplexFieldEnd? fieldEnd = FindComplexFieldEnd(paragraphs, index, fieldStartChildIndex.Value);
                if (fieldEnd == null) {
                    continue;
                }

                SdtBlock sdtBlock = CreateImportedTableOfContentBlock();
                SdtContentBlock content = sdtBlock.SdtContentBlock!;
                int insertIndex = body.ChildElements.ToList().IndexOf(paragraph);
                Paragraph? fieldStartPrefix = SplitFieldStartPrefix(paragraph, fieldStartChildIndex.Value);
                if (fieldStartPrefix != null) {
                    ComplexFieldEnd? adjustedFieldEnd = FindComplexFieldEnd(paragraphs, index, 0);
                    if (adjustedFieldEnd != null) {
                        fieldEnd = adjustedFieldEnd;
                    }
                }

                foreach (Paragraph tocParagraph in paragraphs.Skip(index).Take(fieldEnd.ParagraphIndex - index + 1)) {
                    if (ReferenceEquals(tocParagraph, paragraphs[fieldEnd.ParagraphIndex])) {
                        Paragraph? fieldEndPrefix = SplitFieldEndPrefix(tocParagraph, fieldEnd.ChildIndex);
                        if (fieldEndPrefix != null) {
                            content.Append(fieldEndPrefix);
                            continue;
                        }
                    }

                    tocParagraph.Remove();
                    content.Append(tocParagraph);
                }

                document.AssignNewSdtIds(sdtBlock);
                if (insertIndex >= 0 && insertIndex <= body.ChildElements.Count) {
                    if (fieldStartPrefix != null) {
                        body.InsertAt(fieldStartPrefix, insertIndex);
                        insertIndex++;
                    }

                    body.InsertAt(sdtBlock, insertIndex);
                } else {
                    if (fieldStartPrefix != null) {
                        body.Append(fieldStartPrefix);
                    }

                    body.Append(sdtBlock);
                }

                return new WordTableOfContent(document, sdtBlock, queueUpdateOnOpen: false);
            }

            return null;
        }

        private static int? FindTocOrIndexComplexFieldStart(Paragraph paragraph) {
            List<OpenXmlElement> children = paragraph.ChildElements.ToList();
            var fieldStarts = new Stack<(int StartIndex, System.Text.StringBuilder Instruction)>();
            for (int childIndex = 0; childIndex < children.Count; childIndex++) {
                OpenXmlElement child = children[childIndex];
                foreach (OpenXmlElement descendant in child.Descendants<OpenXmlElement>()) {
                    if (descendant is FieldChar fieldChar && IsFieldBegin(fieldChar)) {
                        fieldStarts.Push((childIndex, new System.Text.StringBuilder()));
                    } else if (descendant is FieldChar endFieldChar && IsFieldEnd(endFieldChar) && fieldStarts.Count > 0) {
                        fieldStarts.Pop();
                    } else if (descendant is FieldCode fieldCode && fieldStarts.Count > 0) {
                        var current = fieldStarts.Peek();
                        current.Instruction.Append(fieldCode.Text);
                        if (!string.IsNullOrWhiteSpace(current.Instruction.ToString()) &&
                            IsTocOrIndexInstruction(current.Instruction.ToString())) {
                            return current.StartIndex;
                        }
                    }
                }
            }

            return null;
        }

        private static int? FindSimpleTocOrIndexFieldChildIndex(Paragraph paragraph) {
            List<OpenXmlElement> children = paragraph.ChildElements.ToList();
            for (int childIndex = 0; childIndex < children.Count; childIndex++) {
                IEnumerable<SimpleField> simpleFields = children[childIndex] is SimpleField directSimpleField
                    ? new[] { directSimpleField }.Concat(children[childIndex].Descendants<SimpleField>())
                    : children[childIndex].Descendants<SimpleField>();

                foreach (SimpleField simpleField in simpleFields) {
                    string instruction = simpleField.Instruction?.Value ?? simpleField.Instruction?.ToString() ?? string.Empty;
                    if (!string.IsNullOrWhiteSpace(instruction) && IsTocOrIndexInstruction(instruction)) {
                        return childIndex;
                    }
                }
            }

            return null;
        }

        private static bool IsTocOrIndexInstruction(string instruction) {
            WordFieldInventory.ParsedFieldInstruction parsed = WordFieldInventory.ParseInstruction(instruction);
            return parsed.FieldType == WordFieldType.TOC ||
                   parsed.FieldType == WordFieldType.Index;
        }

        private static ComplexFieldEnd? FindComplexFieldEnd(IReadOnlyList<Paragraph> paragraphs, int startIndex, int startChildIndex) {
            int depth = 0;
            bool started = false;

            for (int index = startIndex; index < paragraphs.Count; index++) {
                List<OpenXmlElement> children = paragraphs[index].ChildElements.ToList();
                int firstChildIndex = index == startIndex ? startChildIndex : 0;
                for (int childIndex = firstChildIndex; childIndex < children.Count; childIndex++) {
                    foreach (FieldChar fieldChar in children[childIndex].Descendants<FieldChar>()) {
                        if (IsFieldBegin(fieldChar)) {
                            depth++;
                            started = true;
                        } else if (IsFieldEnd(fieldChar) && started) {
                            depth--;
                            if (depth == 0) {
                                return new ComplexFieldEnd(index, childIndex);
                            }
                        }
                    }
                }
            }

            return null;
        }

        private static bool IsFieldBegin(FieldChar fieldChar) =>
            fieldChar.FieldCharType?.Value == FieldCharValues.Begin;

        private static bool IsFieldEnd(FieldChar fieldChar) =>
            fieldChar.FieldCharType?.Value == FieldCharValues.End;

        private static Paragraph? SplitFieldEndPrefix(Paragraph paragraph, int endChildIndex) {
            List<OpenXmlElement> children = paragraph.ChildElements.ToList();
            if (endChildIndex < 0 || !HasMeaningfulContentAfterFieldEnd(children, endChildIndex)) {
                return null;
            }

            var prefix = new Paragraph();
            ParagraphProperties? properties = paragraph.GetFirstChild<ParagraphProperties>();
            if (properties != null) {
                prefix.Append((ParagraphProperties)properties.CloneNode(true));
            }

            for (int index = 0; index <= endChildIndex; index++) {
                if (children[index] is ParagraphProperties) {
                    continue;
                }

                prefix.Append(index == endChildIndex
                    ? CloneThroughFieldEnd(children[index])
                    : children[index].CloneNode(true));
            }

            for (int index = endChildIndex; index >= 0; index--) {
                if (children[index] is ParagraphProperties) {
                    continue;
                }

                if (index == endChildIndex && RemoveThroughFieldEnd(children[index])) {
                    continue;
                }

                children[index].Remove();
            }

            return prefix;
        }

        private static bool HasMeaningfulContentAfterFieldEnd(IReadOnlyList<OpenXmlElement> children, int endChildIndex) {
            return children.Skip(endChildIndex + 1).Any(HasMeaningfulContent) ||
                   HasMeaningfulContentAfterFieldEnd(children[endChildIndex]);
        }

        private static bool HasMeaningfulContentAfterFieldEnd(OpenXmlElement element) {
            FieldChar? end = element.Descendants<FieldChar>().FirstOrDefault(IsFieldEnd);
            if (end == null) {
                return false;
            }

            bool afterEnd = false;
            foreach (OpenXmlElement descendant in element.Descendants<OpenXmlElement>()) {
                if (ReferenceEquals(descendant, end)) {
                    afterEnd = true;
                    continue;
                }

                if (!afterEnd) {
                    continue;
                }

                if (descendant is Text text && !string.IsNullOrWhiteSpace(text.Text)) {
                    return true;
                }

                if (descendant is SimpleField ||
                    descendant is FieldCode ||
                    descendant is DocumentFormat.OpenXml.Wordprocessing.Drawing ||
                    descendant is Picture) {
                    return true;
                }
            }

            return false;
        }

        private static OpenXmlElement CloneThroughFieldEnd(OpenXmlElement element) {
            if (element is not Run run) {
                return element.CloneNode(true);
            }

            var clone = new Run();
            foreach (OpenXmlElement child in run.ChildElements) {
                clone.Append(child.CloneNode(true));
                if (child is FieldChar fieldChar && IsFieldEnd(fieldChar)) {
                    break;
                }
            }

            return clone;
        }

        private static bool RemoveThroughFieldEnd(OpenXmlElement element) {
            if (element is not Run run) {
                return false;
            }

            bool removedEnd = false;
            foreach (OpenXmlElement child in run.ChildElements.ToList()) {
                child.Remove();
                if (child is FieldChar fieldChar && IsFieldEnd(fieldChar)) {
                    removedEnd = true;
                    break;
                }
            }

            if (!removedEnd || !HasMeaningfulContent(run)) {
                return false;
            }

            return true;
        }

        private static Paragraph? SplitFieldStartPrefix(Paragraph paragraph, int startChildIndex) {
            List<OpenXmlElement> children = paragraph.ChildElements.ToList();
            Run? sameRunPrefix = startChildIndex >= 0 && startChildIndex < children.Count
                ? CloneBeforeFieldBegin(children[startChildIndex])
                : null;
            bool hasPreviousContent = startChildIndex > 0 && children.Take(startChildIndex).Any(HasMeaningfulContent);
            if (!hasPreviousContent && sameRunPrefix == null) {
                return null;
            }

            var prefix = new Paragraph();
            ParagraphProperties? properties = paragraph.GetFirstChild<ParagraphProperties>();
            if (properties != null) {
                prefix.Append((ParagraphProperties)properties.CloneNode(true));
            }

            for (int index = 0; index < startChildIndex; index++) {
                if (children[index] is ParagraphProperties) {
                    continue;
                }

                prefix.Append(children[index].CloneNode(true));
            }

            if (sameRunPrefix != null) {
                prefix.Append(sameRunPrefix);
            }

            for (int index = startChildIndex - 1; index >= 0; index--) {
                if (children[index] is ParagraphProperties) {
                    continue;
                }

                children[index].Remove();
            }

            if (startChildIndex >= 0 && startChildIndex < children.Count) {
                RemoveBeforeFieldBegin(children[startChildIndex]);
            }

            return prefix;
        }

        private static Run? CloneBeforeFieldBegin(OpenXmlElement element) {
            if (element is not Run run) {
                return null;
            }

            var clone = new Run();
            foreach (OpenXmlElement child in run.ChildElements) {
                if (child is FieldChar fieldChar && IsFieldBegin(fieldChar)) {
                    break;
                }

                clone.Append(child.CloneNode(true));
            }

            return HasMeaningfulContent(clone) ? clone : null;
        }

        private static void RemoveBeforeFieldBegin(OpenXmlElement element) {
            if (element is not Run run) {
                return;
            }

            foreach (OpenXmlElement child in run.ChildElements.ToList()) {
                if (child is FieldChar fieldChar && IsFieldBegin(fieldChar)) {
                    break;
                }

                child.Remove();
            }
        }

        private static Paragraph? SplitSimpleFieldPrefix(Paragraph paragraph, int fieldChildIndex) {
            List<OpenXmlElement> children = paragraph.ChildElements.ToList();
            if (fieldChildIndex < 0 || !children.Skip(fieldChildIndex + 1).Any(HasMeaningfulContent)) {
                return null;
            }

            var prefix = new Paragraph();
            ParagraphProperties? properties = paragraph.GetFirstChild<ParagraphProperties>();
            if (properties != null) {
                prefix.Append((ParagraphProperties)properties.CloneNode(true));
            }

            for (int index = 0; index <= fieldChildIndex; index++) {
                if (children[index] is ParagraphProperties) {
                    continue;
                }

                prefix.Append(children[index].CloneNode(true));
            }

            for (int index = fieldChildIndex; index >= 0; index--) {
                if (children[index] is ParagraphProperties) {
                    continue;
                }

                children[index].Remove();
            }

            return prefix;
        }

        private sealed class ComplexFieldEnd {
            internal ComplexFieldEnd(int paragraphIndex, int childIndex) {
                ParagraphIndex = paragraphIndex;
                ChildIndex = childIndex;
            }

            internal int ParagraphIndex { get; }

            internal int ChildIndex { get; }
        }

        private static bool HasMeaningfulContent(OpenXmlElement element) {
            return element.Descendants<Text>().Any(text => !string.IsNullOrWhiteSpace(text.Text)) ||
                   element.Descendants<SimpleField>().Any() ||
                   element.Descendants<FieldCode>().Any() ||
                   element.Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>().Any() ||
                   element.Descendants<Picture>().Any();
        }

        private static SdtBlock CreateImportedTableOfContentBlock() {
            return new SdtBlock(
                new SdtProperties(
                    new SdtId(),
                    new SdtContentDocPartObject(
                        new DocPartGallery { Val = "Table of Contents" },
                        new DocPartUnique())),
                new SdtContentBlock());
        }

        /// <summary>
        /// Converts StdBlock to WordElement
        /// </summary>
        /// <param name="document"></param>
        /// <param name="sdtBlock"></param>
        /// <returns></returns>
        internal static WordElement ConvertStdBlockToWordElements(WordDocument document, SdtBlock? sdtBlock) {
            if (sdtBlock == null) {
                return new WordStructuredDocumentTag(document, new SdtBlock());
            }

            var sdtProperties = sdtBlock.ChildElements.OfType<SdtProperties>().FirstOrDefault();
            var docPartObject = sdtProperties?.ChildElements.OfType<SdtContentDocPartObject>().FirstOrDefault();
            var docPartGallery = docPartObject?.ChildElements.OfType<DocPartGallery>().FirstOrDefault();

            if (docPartGallery != null && docPartGallery.Val == "Cover Pages") {
                return new WordCoverPage(document, sdtBlock!);
            } else if (docPartGallery != null && docPartGallery.Val == "Table of Contents") {
                return new WordTableOfContent(document, sdtBlock!);
            }

            var watermark = ConvertStdBlockToWatermark(document, sdtBlock);
            if (watermark != null) {
                return watermark;
            }

            return new WordStructuredDocumentTag(document, sdtBlock!);
        }

        /// <summary>
        /// Converts StdBlock to WordElement
        /// </summary>
        /// <param name="document"></param>
        /// <param name="sdtBlocks"></param>
        /// <returns></returns>
        internal static List<WordElement> ConvertStdBlockToWordElements(WordDocument document, IEnumerable<SdtBlock?> sdtBlocks) {
            var list = new List<WordElement>();
            foreach (var sdtBlock in sdtBlocks) {
                var element = ConvertStdBlockToWordElements(document, sdtBlock);
                if (element != null) {
                    list.Add(element);
                }
            }
            return list;
        }


        /// <summary>
        /// Converts paragraphs to WordParagraphs
        /// </summary>
        /// <param name="document"></param>
        /// <param name="paragraphs"></param>
        /// <returns></returns>
        internal static List<WordParagraph> ConvertParagraphsToWordParagraphs(WordDocument document, IEnumerable<Paragraph> paragraphs) {
            var list = new List<WordParagraph>();

            foreach (Paragraph paragraph in paragraphs) {
                list.AddRange(ConvertParagraphToWordParagraphs(document, paragraph));
            }

            return list;
        }

        /// <summary>
        /// Converts paragraph to WordParagraphs
        /// </summary>
        /// <param name="document"></param>
        /// <param name="paragraph"></param>
        /// <param name="splitPaginationMarkers"></param>
        /// <param name="cancellationToken"></param>
        /// <returns></returns>
        internal static List<WordParagraph> ConvertParagraphToWordParagraphs(
            WordDocument document,
            Paragraph paragraph,
            bool splitPaginationMarkers = false,
            CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            var list = new List<WordParagraph>();
            var childElements = paragraph.ChildElements;
            if (childElements.Count == 1 && childElements[0] is ParagraphProperties) {
                // basically empty, we still want to track it, but that's about it
                list.Add(new WordParagraph(document, paragraph));
            } else if (childElements.Any()) {
                List<Run> runList = new List<Run>();
                bool foundField = false;

                void AddRun(Run run, Hyperlink? hyperlink = null) {
                    cancellationToken.ThrowIfCancellationRequested();
                    WordParagraph wordParagraph;
                    IReadOnlyList<Run> logicalRuns = splitPaginationMarkers ? SplitRunByPaginationMarkers(run) : new[] { run };
                    if (logicalRuns.Count > 1) {
                        foreach (Run logicalRun in logicalRuns) {
                            AddRun(logicalRun, hyperlink);
                        }

                        return;
                    }

                    var fieldChar = run.ChildElements.OfType<FieldChar>().FirstOrDefault();
                    if (foundField == true) {
                        if (fieldChar?.FieldCharType?.Value == FieldCharValues.End) {
                            foundField = false;
                            runList.Add(run);
                            if (!ProcessComplexFieldWithHardBreaks(runList, hyperlink)) {
                                wordParagraph = new WordParagraph(document, paragraph, runList);
                                wordParagraph._hyperlink = hyperlink;
                                list.Add(wordParagraph);
                            }

                            runList = new List<Run>();
                        } else {
                            runList.Add(run);
                        }
                    } else {
                        if (fieldChar?.FieldCharType?.Value == FieldCharValues.Begin) {
                            runList.Add(run);
                            foundField = true;
                        } else {
                            wordParagraph = new WordParagraph(document, paragraph, run);
                            wordParagraph._hyperlink = hyperlink;
                            list.Add(wordParagraph);
                        }
                    }
                }

                void ProcessElement(OpenXmlElement element, Hyperlink? hyperlinkContext = null) {
                    cancellationToken.ThrowIfCancellationRequested();
                    WordParagraph wordParagraph;
                    if (element is Run run) {
                        AddRun(run, hyperlinkContext);
                    } else if (element is InsertedRun || element is MoveToRun) {
                        foreach (OpenXmlElement child in element.ChildElements) {
                            ProcessElement(child, hyperlinkContext);
                        }
                    } else if (element is DeletedRun || element is MoveFromRun) {
                        // Image export follows Word's final view by default: inserted/moved-to text is visible, deleted/moved-from text is hidden.
                    } else if (element is Hyperlink hyperlink) {
                        if (splitPaginationMarkers && ContainsPaginationMarker(hyperlink)) {
                            foreach (OpenXmlElement child in hyperlink.ChildElements) {
                                ProcessElement(child, hyperlink);
                            }
                        } else {
                            wordParagraph = new WordParagraph(document, paragraph, hyperlink);
                            list.Add(wordParagraph);
                        }
                    } else if (element is SimpleField simpleField) {
                        if (!splitPaginationMarkers || !ProcessSimpleFieldWithHardBreaks(simpleField, hyperlinkContext)) {
                            wordParagraph = new WordParagraph(document, paragraph, simpleField);
                            list.Add(wordParagraph);
                        }
                    } else if (element is BookmarkStart bookmarkStart) {
                        wordParagraph = new WordParagraph(document, paragraph, bookmarkStart);
                        list.Add(wordParagraph);
                    } else if (element is BookmarkEnd) {
                        // not needed, we will search for matching bookmark end in the bookmark (i guess)
                    } else if (element is CommentRangeStart || element is CommentRangeEnd) {
                        // Comment range anchors are structural; the visible final-view marker is carried by CommentReference.
                    } else if (element is DocumentFormat.OpenXml.Math.OfficeMath officeMath) {
                        wordParagraph = new WordParagraph(document, paragraph, officeMath);
                        list.Add(wordParagraph);
                    } else if (element is DocumentFormat.OpenXml.Math.Paragraph mathParagraph) {
                        wordParagraph = new WordParagraph(document, paragraph, mathParagraph);
                        list.Add(wordParagraph);
                    } else if (element is SdtRun sdtRun) {
                        if (splitPaginationMarkers && ProcessInlineContentControl(sdtRun, hyperlinkContext)) {
                            return;
                        }

                        list.Add(new WordParagraph(document, paragraph, sdtRun));
                    } else if (element is ProofError) {

                    } else if (element is ParagraphProperties) {

                    } else {
                        Debug.WriteLine("Please implement me! " + element.GetType().Name);
                    }
                }

                bool ProcessInlineContentControl(SdtRun sdtRun, Hyperlink? hyperlinkContext) {
                    if (!ContainsPaginationMarker(sdtRun)) {
                        return false;
                    }

                    SdtContentRun? contentRun = sdtRun.SdtContentRun;
                    if (contentRun == null || contentRun.ChildElements.Count == 0) {
                        return false;
                    }

                    foreach (OpenXmlElement child in contentRun.ChildElements) {
                        cancellationToken.ThrowIfCancellationRequested();
                        ProcessElement(child, hyperlinkContext);
                    }

                    return true;
                }

                bool ProcessSimpleFieldWithHardBreaks(SimpleField simpleField, Hyperlink? hyperlinkContext) {
                    if (!ContainsPaginationMarker(simpleField)) {
                        return false;
                    }

                    foreach (OpenXmlElement child in simpleField.ChildElements) {
                        cancellationToken.ThrowIfCancellationRequested();
                        ProcessElement(child, hyperlinkContext);
                    }

                    return true;
                }

                bool ProcessComplexFieldWithHardBreaks(IReadOnlyList<Run> fieldRuns, Hyperlink? hyperlinkContext) {
                    List<Run> resultRuns = ExtractComplexFieldResultRuns(fieldRuns);
                    if (!resultRuns.Any(ContainsPaginationMarker)) {
                        return false;
                    }

                    foreach (Run resultRun in resultRuns) {
                        cancellationToken.ThrowIfCancellationRequested();
                        AddRun(resultRun, hyperlinkContext);
                    }

                    return true;
                }

                foreach (var element in paragraph.ChildElements) {
                    cancellationToken.ThrowIfCancellationRequested();
                    ProcessElement(element);
                }
            } else {
                // add empty word paragraph
                list.Add(new WordParagraph(document, paragraph));
            }
            return list;
        }

        private static List<Run> ExtractComplexFieldResultRuns(IReadOnlyList<Run> fieldRuns) {
            var resultRuns = new List<Run>();
            bool sawSeparator = false;
            int fieldDepth = 0;

            foreach (Run run in fieldRuns) {
                RunProperties? runProperties = run.GetFirstChild<RunProperties>();
                var currentChildren = new List<OpenXmlElement>();

                void FlushCurrentChildren() {
                    if (currentChildren.Count == 0) {
                        return;
                    }

                    resultRuns.Add(CreateLogicalRun(runProperties, currentChildren));
                    currentChildren.Clear();
                }

                foreach (OpenXmlElement child in run.ChildElements) {
                    if (child is RunProperties) {
                        continue;
                    }

                    if (child is FieldChar fieldChar) {
                        FieldCharValues? fieldCharType = fieldChar.FieldCharType?.Value;
                        if (fieldCharType == FieldCharValues.Begin) {
                            fieldDepth++;
                            continue;
                        }

                        if (fieldCharType == FieldCharValues.Separate && fieldDepth > 0) {
                            sawSeparator = true;
                            continue;
                        }

                        if (fieldCharType == FieldCharValues.End && fieldDepth > 0) {
                            FlushCurrentChildren();
                            fieldDepth--;
                            if (fieldDepth == 0) {
                                return resultRuns;
                            }

                            continue;
                        }
                    }

                    if (sawSeparator && fieldDepth > 0) {
                        currentChildren.Add((OpenXmlElement)child.CloneNode(true));
                    }
                }

                FlushCurrentChildren();
            }

            return resultRuns;
        }

        private static IReadOnlyList<Run> SplitRunByPaginationMarkers(Run run) {
            if (!ContainsPaginationMarker(run)) {
                return new[] { run };
            }

            RunProperties? runProperties = run.GetFirstChild<RunProperties>();
            var result = new List<Run>();
            var currentChildren = new List<OpenXmlElement>();

            void FlushCurrentChildren() {
                if (currentChildren.Count == 0) {
                    return;
                }

                result.Add(CreateLogicalRun(runProperties, currentChildren));
                currentChildren.Clear();
            }

            foreach (OpenXmlElement child in run.ChildElements) {
                if (child is RunProperties) {
                    continue;
                }

                if (child is Break breakNode && IsHardPaginationBreak(breakNode)) {
                    FlushCurrentChildren();
                    result.Add(CreateLogicalRun(runProperties, new OpenXmlElement[] { (OpenXmlElement)breakNode.CloneNode(true) }));
                    continue;
                }

                if (child is LastRenderedPageBreak) {
                    FlushCurrentChildren();
                    result.Add(CreateLogicalRun(runProperties, new OpenXmlElement[] { new Break { Type = BreakValues.Page } }));
                    continue;
                }

                currentChildren.Add((OpenXmlElement)child.CloneNode(true));
            }

            FlushCurrentChildren();
            return result.Count == 0 ? new[] { run } : result;
        }

        private static Run CreateLogicalRun(RunProperties? runProperties, IEnumerable<OpenXmlElement> children) {
            var logicalRun = new Run();
            if (runProperties != null) {
                logicalRun.Append((RunProperties)runProperties.CloneNode(true));
            }

            foreach (OpenXmlElement child in children) {
                logicalRun.Append(child);
            }

            return logicalRun;
        }

        private static bool IsHardPaginationBreak(Break breakNode) =>
            breakNode.Type?.Value == BreakValues.Page ||
            breakNode.Type?.Value == BreakValues.Column;

        private static bool ContainsPaginationMarker(OpenXmlElement element) =>
            element.Descendants<LastRenderedPageBreak>().Any() ||
            element.Descendants<Break>().Any(IsHardPaginationBreak);

        private int GetSectionOrdinal() {
            int sectionIndex = _document.Sections.IndexOf(this);
            if (sectionIndex < 0) {
                throw new InvalidOperationException("The section is not attached to the document.");
            }

            return sectionIndex;
        }

        private int GetSectionCount() {
            return Math.Max(_document.Sections.Count, 1);
        }

        private static bool IsSectionBoundaryParagraph(Paragraph paragraph) {
            return paragraph.ParagraphProperties?.SectionProperties != null;
        }

        private static bool IsPureSectionBreakParagraph(Paragraph paragraph) {
            if (!IsSectionBoundaryParagraph(paragraph)) {
                return false;
            }

            if (paragraph.ChildElements.Any(element => element is not ParagraphProperties)) {
                return false;
            }

            return paragraph.ParagraphProperties?.ChildElements.All(element => element is SectionProperties) != false;
        }

        /// <summary>
        /// Get all paragraphs in given section
        /// </summary>
        /// <returns></returns>
        private List<WordParagraph> GetParagraphsList() {
            int targetSection = GetSectionOrdinal();
            var paragraphsBySection = Enumerable.Range(0, GetSectionCount())
                .Select(_ => new List<Paragraph>())
                .ToList();

            var body = _wordprocessingDocument?.MainDocumentPart?.Document?.Body;
            if (body == null) {
                return new List<WordParagraph>();
            }

            int currentSection = 0;
            foreach (var element in body.ChildElements) {
                if (element is not Paragraph paragraph) {
                    continue;
                }

                if (!IsPureSectionBreakParagraph(paragraph)) {
                    paragraphsBySection[currentSection].Add(paragraph);
                }

                if (IsSectionBoundaryParagraph(paragraph) && currentSection < paragraphsBySection.Count - 1) {
                    currentSection++;
                }
            }

            return ConvertParagraphsToWordParagraphs(_document, paragraphsBySection[targetSection]);
        }

        /// <summary>
        /// This method gets all lists for given document (for all sections)
        /// </summary>
        /// <param name="document"></param>
        /// <returns></returns>
        internal static List<WordList> GetAllDocumentsLists(WordDocument document) {
            var numbering = document._wordprocessingDocument.MainDocumentPart?.NumberingDefinitionsPart?.Numbering;
            if (numbering == null) {
                return new List<WordList>(0);
            }

            return numbering.ChildElements.OfType<NumberingInstance>()
                .Select(element => new WordList(document, element.NumberID!.Value))
                .ToList();
        }

        /// <summary>
        /// This method gets lists for given section. It's possible that given list spreads thru multiple sections.
        /// Therefore sum of all sections lists doesn't necessary match all lists count for a document.
        /// </summary>
        /// <returns></returns>
        private List<WordList> GetLists() {
            List<WordList> allLists = GetAllDocumentsLists(_document);

            List<WordList> lists = new List<WordList>();
            var usedNumbers = this.ParagraphListItemsNumbers;
            foreach (var list in allLists) {
                if (usedNumbers.Contains(list._numberId)) {
                    lists.Add(list);
                }
            }
            return lists;
        }

        /// <summary>
        /// Gets list of tables in given section
        /// </summary>
        /// <returns></returns>
        private List<WordTable> GetTablesList() {
            int targetSection = GetSectionOrdinal();
            var tablesBySection = Enumerable.Range(0, GetSectionCount())
                .Select(_ => new List<WordTable>())
                .ToList();

            var body = _wordprocessingDocument?.MainDocumentPart?.Document?.Body;
            if (body == null) {
                return new List<WordTable>();
            }

            int currentSection = 0;
            foreach (var element in body.ChildElements) {
                if (element is Paragraph paragraph) {
                    if (IsSectionBoundaryParagraph(paragraph) && currentSection < tablesBySection.Count - 1) {
                        currentSection++;
                    }
                } else if (element is Table table) {
                    tablesBySection[currentSection].Add(new WordTable(_document, table));
                }
            }

            return tablesBySection[targetSection];
        }

        /// <summary>
        /// Gets list of embedded documents in given section
        /// </summary>
        /// <returns></returns>
        private List<WordEmbeddedDocument> GetEmbeddedDocumentsList() {
            int targetSection = GetSectionOrdinal();
            var embeddedDocumentsBySection = Enumerable.Range(0, GetSectionCount())
                .Select(_ => new List<WordEmbeddedDocument>())
                .ToList();

            var body = _wordprocessingDocument?.MainDocumentPart?.Document?.Body;
            if (body == null) {
                return new List<WordEmbeddedDocument>();
            }

            int currentSection = 0;
            foreach (var element in body.ChildElements) {
                if (element is Paragraph paragraph) {
                    if (IsSectionBoundaryParagraph(paragraph) && currentSection < embeddedDocumentsBySection.Count - 1) {
                        currentSection++;
                    }
                } else if (element is AltChunk altChunk) {
                    embeddedDocumentsBySection[currentSection].Add(new WordEmbeddedDocument(_document, altChunk));
                }
            }

            return embeddedDocumentsBySection[targetSection];
        }

        private List<WordEmbeddedObject> GetEmbeddedObjectsList() {
            int targetSection = GetSectionOrdinal();
            var embeddedObjectsBySection = Enumerable.Range(0, GetSectionCount())
                .Select(_ => new List<WordEmbeddedObject>())
                .ToList();

            var body = _wordprocessingDocument?.MainDocumentPart?.Document?.Body;
            if (body == null) {
                return new List<WordEmbeddedObject>();
            }

            int currentSection = 0;
            foreach (var element in body.ChildElements) {
                if (element is not Paragraph paragraph) {
                    continue;
                }

                foreach (var run in paragraph.ChildElements.OfType<Run>()) {
                    if (run.Descendants<Ovml.OleObject>().Any()) {
                        embeddedObjectsBySection[currentSection].Add(new WordEmbeddedObject(_document, run));
                    }
                }

                if (IsSectionBoundaryParagraph(paragraph) && currentSection < embeddedObjectsBySection.Count - 1) {
                    currentSection++;
                }
            }

            return embeddedObjectsBySection[targetSection];
        }

        /// <summary>
        /// Gets list of word elements in given section
        /// </summary>
        /// <returns></returns>
        private List<WordElement> GetWordElements() {
            int targetSection = GetSectionOrdinal();
            var elementsBySection = Enumerable.Range(0, GetSectionCount())
                .Select(_ => new List<WordElement>())
                .ToList();

            var body = _wordprocessingDocument?.MainDocumentPart?.Document?.Body;
            if (body == null) {
                return new List<WordElement>();
            }

            int currentSection = 0;
            foreach (var element in body.ChildElements) {
                if (element is Paragraph paragraph) {
                    if (!IsPureSectionBreakParagraph(paragraph)) {
                        elementsBySection[currentSection].AddRange(ConvertParagraphToWordParagraphs(_document, paragraph));
                    }

                    if (IsSectionBoundaryParagraph(paragraph) && currentSection < elementsBySection.Count - 1) {
                        currentSection++;
                    }
                } else if (element is AltChunk altChunk) {
                    elementsBySection[currentSection].Add(new WordEmbeddedDocument(_document, altChunk));
                } else if (element is SdtBlock sdtBlock) {
                    elementsBySection[currentSection].Add(ConvertStdBlockToWordElements(_document, sdtBlock));
                } else if (element is Table table) {
                    elementsBySection[currentSection].Add(new WordTable(_document, table));
                }
            }

            return elementsBySection[targetSection];
        }

        /// <summary>
        /// Gets list of word elements by type in given section
        /// </summary>
        /// <returns></returns>
        private List<WordElement> GetWordElementsByType() {
            var listElements = GetWordElements();
            var additionalElements = new List<WordElement>();

            foreach (var element in listElements) {
                if (element is WordParagraph wordParagraph) {
                    if (wordParagraph.IsBookmark) {
                        additionalElements.Add(wordParagraph.Bookmark!);
                    } else if (wordParagraph.IsBreak) {
                        additionalElements.Add(wordParagraph.Break!);
                    } else if (wordParagraph.IsChart) {
                        additionalElements.Add(wordParagraph.Chart!);
                    } else if (wordParagraph.IsEndNote) {
                        additionalElements.Add(wordParagraph.EndNote!);
                    } else if (wordParagraph.IsEquation) {
                        additionalElements.Add(wordParagraph.Equation!);
                    } else if (wordParagraph.IsField) {
                        additionalElements.Add(wordParagraph.Field!);
                    } else if (wordParagraph.IsFootNote) {
                        additionalElements.Add(wordParagraph.FootNote!);
                    } else if (wordParagraph.IsImage) {
                        additionalElements.Add(wordParagraph.Image!);
                    } else if (wordParagraph.IsListItem) {
                        additionalElements.Add(wordParagraph);
                    } else if (wordParagraph.IsPageBreak) {
                        additionalElements.Add(wordParagraph.PageBreak!);
                    } else if (wordParagraph.IsStructuredDocumentTag) {
                        additionalElements.Add(wordParagraph.StructuredDocumentTag!);
                    } else if (wordParagraph.IsTab) {
                        additionalElements.Add(wordParagraph.Tab!);
                    } else if (wordParagraph.IsTextBox) {
                        additionalElements.Add(wordParagraph.TextBox!);
                    } else if (wordParagraph.IsHyperLink) {
                        additionalElements.Add(wordParagraph.Hyperlink!);
                    } else {
                        additionalElements.Add(wordParagraph);
                    }
                } else {
                    additionalElements.Add(element);
                }
            }
            return additionalElements;
        }

        /// <summary>
        /// Gets list of watermarks in given section
        /// </summary>
        /// <returns></returns>
        private List<SdtBlock> GetSdtBlockList() {
            int targetSection = GetSectionOrdinal();
            var sdtBlocksBySection = Enumerable.Range(0, GetSectionCount())
                .Select(_ => new List<SdtBlock>())
                .ToList();

            var body = _wordprocessingDocument?.MainDocumentPart?.Document?.Body;
            if (body == null) {
                return new List<SdtBlock>();
            }

            int currentSection = 0;
            foreach (var element in body.ChildElements) {
                if (element is Paragraph paragraph) {
                    if (IsSectionBoundaryParagraph(paragraph) && currentSection < sdtBlocksBySection.Count - 1) {
                        currentSection++;
                    }
                } else if (element is SdtBlock sdtBlock) {
                    sdtBlocksBySection[currentSection].Add(sdtBlock);
                }
            }

            return sdtBlocksBySection[targetSection];
        }

        /// <summary>
        /// This method moves headers and footers and title page to section before it.
        /// It also copies all other parts of sections (PageSize,PageMargin and others) to section before it.
        /// This is because headers/footers when applied to section apply to the rest of the document
        /// unless there are headers/footers on next section.
        /// On the other hand page size doesn't apply to other sections
        /// and word uses default values. 
        /// </summary>
        /// <param name="sectionProperties"></param>
        /// <param name="newSectionProperties"></param>
        private static void CopySectionProperties(SectionProperties sectionProperties, SectionProperties newSectionProperties) {
            bool canMoveInheritedSectionParts = newSectionProperties.ChildElements.Count == 0;
            if (canMoveInheritedSectionParts || HasOnlySectionType(newSectionProperties)) {
                var listSectionEntries = sectionProperties.ChildElements.ToList();
                foreach (var element in listSectionEntries) {
                    if (element is HeaderReference) {
                        if (canMoveInheritedSectionParts) {
                            newSectionProperties.Append(element.CloneNode(true));
                            sectionProperties.RemoveChild(element);
                        }
                    } else if (element is FooterReference) {
                        if (canMoveInheritedSectionParts) {
                            newSectionProperties.Append(element.CloneNode(true));
                            sectionProperties.RemoveChild(element);
                        }
                    } else if (element is PageSize) {
                        AppendMissingSectionElement<PageSize>(newSectionProperties, element);
                    } else if (element is PageMargin) {
                        AppendMissingSectionElement<PageMargin>(newSectionProperties, element);
                        //sectionProperties.RemoveChild(element);
                        //} else if (element is Columns) {
                        //    newSectionProperties.Append(element.CloneNode(true));
                        //} else if (element is DocGrid) {
                        //    newSectionProperties.Append(element.CloneNode(true));
                        //} else if (element is SectionType) {
                        //    newSectionProperties.Append(element.CloneNode(true));
                    } else if (element is FootnoteProperties footnoteProps) {
                        if (canMoveInheritedSectionParts) {
                            var cloned = (FootnoteProperties)footnoteProps.CloneNode(true);
                            cloned.RemoveAllChildren<NumberingRestart>();
                            newSectionProperties.Append(cloned);
                            footnoteProps.RemoveAllChildren<NumberingRestart>();
                        }
                    } else if (element is EndnoteProperties endnoteProps) {
                        if (canMoveInheritedSectionParts) {
                            var cloned = (EndnoteProperties)endnoteProps.CloneNode(true);
                            cloned.RemoveAllChildren<NumberingRestart>();
                            newSectionProperties.Append(cloned);
                            endnoteProps.RemoveAllChildren<NumberingRestart>();
                        }
                    } else if (element is TitlePage) {
                        if (canMoveInheritedSectionParts) {
                            newSectionProperties.Append(element.CloneNode(true));
                            sectionProperties.RemoveChild(element);
                        }
                    } else {
                        if (canMoveInheritedSectionParts) {
                            newSectionProperties.Append(element.CloneNode(true));
                        }
                        //throw new NotImplementedException("This isn't implemented yet?");
                    }
                }
                //sectionProperties.RemoveAllChildren();
                //newSectionProperties.Append(listSectionEntries);
            }
        }

        private static bool HasOnlySectionType(SectionProperties sectionProperties) =>
            sectionProperties.ChildElements.Count > 0 &&
            sectionProperties.ChildElements.All(element => element is SectionType);

        private static void AppendMissingSectionElement<TElement>(SectionProperties sectionProperties, OpenXmlElement source)
            where TElement : OpenXmlElement {
            if (!sectionProperties.Elements<TElement>().Any()) {
                sectionProperties.Append(source.CloneNode(true));
            }
        }

    }
}
