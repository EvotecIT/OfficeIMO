using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text;
using V = DocumentFormat.OpenXml.Vml;

namespace OfficeIMO.Word {
    public static partial class WordDocumentComparer {
        private static void AnalyzeParagraphs(WordDocument source, WordDocument target, WordComparisonResult result, WordComparisonOptions options) {
            List<ParagraphSnapshot> sourceParagraphs = GetLogicalBodyParagraphs(source, options);
            List<ParagraphSnapshot> targetParagraphs = GetLogicalBodyParagraphs(target, options);
            IReadOnlyList<MatchedIndexPair> matchedParagraphs = FindMatchingIndexes(
                sourceParagraphs,
                targetParagraphs,
                ParagraphSnapshotEqualityComparer.Instance);

            int sourceStart = 0;
            int targetStart = 0;

            foreach (MatchedIndexPair match in matchedParagraphs) {
                AddParagraphRangeFindings(sourceParagraphs, targetParagraphs, sourceStart, match.SourceIndex, targetStart, match.TargetIndex, result, options);
                AnalyzeParagraphStyle(sourceParagraphs[match.SourceIndex], targetParagraphs[match.TargetIndex], match.SourceIndex, match.TargetIndex, result, options);
                AnalyzeParagraphEffectiveFormatting(sourceParagraphs[match.SourceIndex], targetParagraphs[match.TargetIndex], match.SourceIndex, match.TargetIndex, result, options);
                AnalyzeParagraphRuns(sourceParagraphs[match.SourceIndex], targetParagraphs[match.TargetIndex], match.SourceIndex, match.TargetIndex, result, options);
                sourceStart = match.SourceIndex + 1;
                targetStart = match.TargetIndex + 1;
            }

            AddParagraphRangeFindings(sourceParagraphs, targetParagraphs, sourceStart, sourceParagraphs.Count, targetStart, targetParagraphs.Count, result, options);
        }

        private static void AddParagraphRangeFindings(
            IReadOnlyList<ParagraphSnapshot> sourceParagraphs,
            IReadOnlyList<ParagraphSnapshot> targetParagraphs,
            int sourceStart,
            int sourceEnd,
            int targetStart,
            int targetEnd,
            WordComparisonResult result,
            WordComparisonOptions options) {
            int sourceIndex = sourceStart;
            int targetIndex = targetStart;

            while (sourceIndex < sourceEnd && targetIndex < targetEnd) {
                int betterTargetIndex = FindBetterTargetAlignmentIndex(sourceParagraphs[sourceIndex], targetParagraphs, targetIndex, targetEnd);
                if (betterTargetIndex > targetIndex) {
                    while (targetIndex < betterTargetIndex) {
                        AddInsertedParagraphFinding(targetParagraphs, targetIndex, result);
                        targetIndex++;
                    }

                    continue;
                }

                int betterSourceIndex = FindBetterSourceAlignmentIndex(sourceParagraphs, sourceIndex, sourceEnd, targetParagraphs[targetIndex]);
                if (betterSourceIndex > sourceIndex) {
                    while (sourceIndex < betterSourceIndex) {
                        AddDeletedParagraphFinding(sourceParagraphs, sourceIndex, result);
                        sourceIndex++;
                    }

                    continue;
                }

                string sourceText = sourceParagraphs[sourceIndex].Text;
                string targetText = targetParagraphs[targetIndex].Text;

                if (!string.Equals(sourceParagraphs[sourceIndex].PartKind, targetParagraphs[targetIndex].PartKind, StringComparison.Ordinal)) {
                    if (AreEquivalentRenumberedNoteParagraphs(sourceParagraphs[sourceIndex], targetParagraphs[targetIndex])) {
                        AnalyzeParagraphStyle(sourceParagraphs[sourceIndex], targetParagraphs[targetIndex], sourceIndex, targetIndex, result, options);
                        AnalyzeParagraphEffectiveFormatting(sourceParagraphs[sourceIndex], targetParagraphs[targetIndex], sourceIndex, targetIndex, result, options);
                        AnalyzeParagraphRuns(sourceParagraphs[sourceIndex], targetParagraphs[targetIndex], sourceIndex, targetIndex, result, options);
                        sourceIndex++;
                        targetIndex++;
                        continue;
                    }

                    result.Add(new WordComparisonFinding(
                        WordComparisonScope.Paragraph,
                        WordComparisonChangeKind.Deleted,
                        ParagraphLocation(sourceIndex),
                        sourceIndex,
                        null,
                        sourceText,
                        null,
                        "Paragraph deleted."),
                        sourceParagraphs[sourceIndex].DocumentOrder);
                    result.Add(new WordComparisonFinding(
                        WordComparisonScope.Paragraph,
                        WordComparisonChangeKind.Inserted,
                        ParagraphLocation(targetIndex),
                        null,
                        targetIndex,
                        null,
                        targetText,
                        "Paragraph inserted."),
                        targetParagraphs[targetIndex].DocumentOrder);
                    sourceIndex++;
                    targetIndex++;
                    continue;
                }

                if (!string.Equals(sourceParagraphs[sourceIndex].MatchText, targetParagraphs[targetIndex].MatchText, StringComparison.Ordinal)) {
                    result.Add(new WordComparisonFinding(
                        WordComparisonScope.Paragraph,
                        WordComparisonChangeKind.Modified,
                        ParagraphLocation(targetIndex),
                        sourceIndex,
                        targetIndex,
                        sourceText,
                        targetText,
                        "Paragraph text changed."),
                        targetParagraphs[targetIndex].DocumentOrder);
                }

                AnalyzeParagraphStyle(sourceParagraphs[sourceIndex], targetParagraphs[targetIndex], sourceIndex, targetIndex, result, options);
                AnalyzeParagraphEffectiveFormatting(sourceParagraphs[sourceIndex], targetParagraphs[targetIndex], sourceIndex, targetIndex, result, options);
                AnalyzeParagraphRuns(sourceParagraphs[sourceIndex], targetParagraphs[targetIndex], sourceIndex, targetIndex, result, options);

                sourceIndex++;
                targetIndex++;
            }

            while (targetIndex < targetEnd) {
                AddInsertedParagraphFinding(targetParagraphs, targetIndex, result);
                targetIndex++;
            }

            while (sourceIndex < sourceEnd) {
                AddDeletedParagraphFinding(sourceParagraphs, sourceIndex, result);
                sourceIndex++;
            }
        }

        private static bool AreEquivalentRenumberedNoteParagraphs(ParagraphSnapshot sourceParagraph, ParagraphSnapshot targetParagraph) {
            if (!IsNotePartKind(sourceParagraph.PartKind) ||
                !IsNotePartKind(targetParagraph.PartKind) ||
                !string.Equals(GetNotePartKindPrefix(sourceParagraph.PartKind), GetNotePartKindPrefix(targetParagraph.PartKind), StringComparison.Ordinal)) {
                return false;
            }

            return string.Equals(sourceParagraph.MatchText, targetParagraph.MatchText, StringComparison.Ordinal) &&
                string.Equals(sourceParagraph.ComparisonText, targetParagraph.ComparisonText, StringComparison.Ordinal);
        }

        private static bool IsNotePartKind(string partKind) =>
            partKind.StartsWith(FootnotePartKeyPrefix, StringComparison.Ordinal) ||
            partKind.StartsWith(EndnotePartKeyPrefix, StringComparison.Ordinal);

        private static string GetNotePartKindPrefix(string partKind) {
            if (partKind.StartsWith(FootnotePartKeyPrefix, StringComparison.Ordinal)) {
                return FootnotePartKeyPrefix;
            }

            return partKind.StartsWith(EndnotePartKeyPrefix, StringComparison.Ordinal)
                ? EndnotePartKeyPrefix
                : string.Empty;
        }

        private static void AnalyzeParagraphStyle(
            ParagraphSnapshot sourceParagraph,
            ParagraphSnapshot targetParagraph,
            int sourceParagraphIndex,
            int targetParagraphIndex,
            WordComparisonResult result,
            WordComparisonOptions options) {
            if (!options.CompareParagraphStyleIds) {
                return;
            }

            if (string.Equals(sourceParagraph.StyleId, targetParagraph.StyleId, StringComparison.Ordinal)) {
                return;
            }

            result.Add(new WordComparisonFinding(
                WordComparisonScope.Paragraph,
                WordComparisonChangeKind.Modified,
                ParagraphLocation(targetParagraphIndex),
                sourceParagraphIndex,
                targetParagraphIndex,
                sourceParagraph.StyleId,
                targetParagraph.StyleId,
                "Paragraph style id changed."),
                targetParagraph.DocumentOrder);
        }

        private static void AnalyzeParagraphEffectiveFormatting(
            ParagraphSnapshot sourceParagraph,
            ParagraphSnapshot targetParagraph,
            int sourceParagraphIndex,
            int targetParagraphIndex,
            WordComparisonResult result,
            WordComparisonOptions options) {
            if (!options.CompareEffectiveFormatting) {
                return;
            }

            if (string.Equals(sourceParagraph.FormatSignature, targetParagraph.FormatSignature, StringComparison.Ordinal)) {
                return;
            }

            result.Add(new WordComparisonFinding(
                WordComparisonScope.Paragraph,
                WordComparisonChangeKind.Modified,
                ParagraphLocation(targetParagraphIndex),
                sourceParagraphIndex,
                targetParagraphIndex,
                sourceParagraph.StyleId,
                targetParagraph.StyleId,
                "Paragraph effective formatting changed."),
                targetParagraph.DocumentOrder);
        }

        private static int FindBetterTargetAlignmentIndex(ParagraphSnapshot sourceParagraph, IReadOnlyList<ParagraphSnapshot> targetParagraphs, int targetStart, int targetEnd) {
            int bestIndex = targetStart;
            double currentVisibleSimilarity = GetParagraphVisibleTextSimilarity(sourceParagraph, targetParagraphs[targetStart]);
            double bestVisibleSimilarity = currentVisibleSimilarity;
            double bestSimilarity = GetParagraphSimilarity(sourceParagraph, targetParagraphs[targetStart]);

            foreach (int index in SelectBoundedAlignmentCandidates(
                targetStart + 1,
                targetEnd,
                index => GetParagraphAlignmentPrefilterSimilarity(sourceParagraph, targetParagraphs[index]))) {
                double visibleSimilarity = GetParagraphVisibleTextSimilarity(sourceParagraph, targetParagraphs[index]);
                double similarity = GetParagraphSimilarity(sourceParagraph, targetParagraphs[index]);
                if (visibleSimilarity < bestVisibleSimilarity ||
                    (visibleSimilarity == bestVisibleSimilarity && similarity <= bestSimilarity)) {
                    continue;
                }

                bestVisibleSimilarity = visibleSimilarity;
                bestSimilarity = similarity;
                bestIndex = index;
                if (similarity >= 1) {
                    break;
                }
            }

            return bestVisibleSimilarity > currentVisibleSimilarity || bestSimilarity > GetParagraphSimilarity(sourceParagraph, targetParagraphs[targetStart])
                ? bestIndex
                : targetStart;
        }

        private static int FindBetterSourceAlignmentIndex(IReadOnlyList<ParagraphSnapshot> sourceParagraphs, int sourceStart, int sourceEnd, ParagraphSnapshot targetParagraph) {
            int bestIndex = sourceStart;
            double currentVisibleSimilarity = GetParagraphVisibleTextSimilarity(sourceParagraphs[sourceStart], targetParagraph);
            double bestVisibleSimilarity = currentVisibleSimilarity;
            double bestSimilarity = GetParagraphSimilarity(sourceParagraphs[sourceStart], targetParagraph);

            foreach (int index in SelectBoundedAlignmentCandidates(
                sourceStart + 1,
                sourceEnd,
                index => GetParagraphAlignmentPrefilterSimilarity(sourceParagraphs[index], targetParagraph))) {
                double visibleSimilarity = GetParagraphVisibleTextSimilarity(sourceParagraphs[index], targetParagraph);
                double similarity = GetParagraphSimilarity(sourceParagraphs[index], targetParagraph);
                if (visibleSimilarity < bestVisibleSimilarity ||
                    (visibleSimilarity == bestVisibleSimilarity && similarity <= bestSimilarity)) {
                    continue;
                }

                bestVisibleSimilarity = visibleSimilarity;
                bestSimilarity = similarity;
                bestIndex = index;
                if (similarity >= 1) {
                    break;
                }
            }

            return bestVisibleSimilarity > currentVisibleSimilarity || bestSimilarity > GetParagraphSimilarity(sourceParagraphs[sourceStart], targetParagraph)
                ? bestIndex
                : sourceStart;
        }

        private static double GetParagraphSimilarity(ParagraphSnapshot sourceParagraph, ParagraphSnapshot targetParagraph) {
            if (!string.Equals(sourceParagraph.PartKind, targetParagraph.PartKind, StringComparison.Ordinal)) {
                return 0;
            }

            return Math.Max(
                GetTextSimilarity(sourceParagraph.MatchText, targetParagraph.MatchText),
                GetTextSimilarity(sourceParagraph.ComparisonText, targetParagraph.ComparisonText));
        }

        private static double GetParagraphAlignmentPrefilterSimilarity(ParagraphSnapshot sourceParagraph, ParagraphSnapshot targetParagraph) {
            if (!string.Equals(sourceParagraph.PartKind, targetParagraph.PartKind, StringComparison.Ordinal)) {
                return 0;
            }

            return Math.Max(
                GetBoundedTextSimilarity(sourceParagraph.MatchText, targetParagraph.MatchText),
                GetBoundedTextSimilarity(sourceParagraph.ComparisonText, targetParagraph.ComparisonText));
        }

        private static IReadOnlyList<int> SelectBoundedAlignmentCandidates(int start, int end, Func<int, double> getPrefilterSimilarity) {
            var candidates = new SortedSet<AlignmentCandidate>(AlignmentCandidateComparer.Instance);
            for (int index = start; index < end; index++) {
                var candidate = new AlignmentCandidate(index, getPrefilterSimilarity(index));
                candidates.Add(candidate);
                if (candidates.Count > MaxComparisonAlignmentWindow) {
                    candidates.Remove(candidates.Min);
                }
            }

            return candidates
                .OrderByDescending(candidate => candidate.Similarity)
                .ThenBy(candidate => candidate.Index)
                .Select(candidate => candidate.Index)
                .ToArray();
        }

        private readonly struct AlignmentCandidate {
            internal AlignmentCandidate(int index, double similarity) {
                Index = index;
                Similarity = similarity;
            }

            internal int Index { get; }
            internal double Similarity { get; }
        }

        private sealed class AlignmentCandidateComparer : IComparer<AlignmentCandidate> {
            internal static readonly AlignmentCandidateComparer Instance = new AlignmentCandidateComparer();

            public int Compare(AlignmentCandidate x, AlignmentCandidate y) {
                int similarity = x.Similarity.CompareTo(y.Similarity);
                return similarity != 0 ? similarity : y.Index.CompareTo(x.Index);
            }
        }

        private static double GetParagraphVisibleTextSimilarity(ParagraphSnapshot sourceParagraph, ParagraphSnapshot targetParagraph) {
            if (!string.Equals(sourceParagraph.PartKind, targetParagraph.PartKind, StringComparison.Ordinal)) {
                return 0;
            }

            if (sourceParagraph.ComparisonText.Length > 0 &&
                targetParagraph.ComparisonText.Length > 0 &&
                (sourceParagraph.ComparisonText.IndexOf(targetParagraph.ComparisonText, StringComparison.Ordinal) >= 0 ||
                 targetParagraph.ComparisonText.IndexOf(sourceParagraph.ComparisonText, StringComparison.Ordinal) >= 0)) {
                return GetContainmentAwareTextSimilarity(sourceParagraph.ComparisonText, targetParagraph.ComparisonText);
            }

            return GetTextSimilarity(sourceParagraph.ComparisonText, targetParagraph.ComparisonText);
        }

        private static void AddInsertedParagraphFinding(IReadOnlyList<ParagraphSnapshot> targetParagraphs, int targetIndex, WordComparisonResult result) {
            result.Add(new WordComparisonFinding(
                WordComparisonScope.Paragraph,
                WordComparisonChangeKind.Inserted,
                ParagraphLocation(targetIndex),
                null,
                targetIndex,
                null,
                targetParagraphs[targetIndex].Text,
                "Paragraph inserted."),
                targetParagraphs[targetIndex].DocumentOrder);
        }

        private static void AddDeletedParagraphFinding(IReadOnlyList<ParagraphSnapshot> sourceParagraphs, int sourceIndex, WordComparisonResult result) {
            result.Add(new WordComparisonFinding(
                WordComparisonScope.Paragraph,
                WordComparisonChangeKind.Deleted,
                ParagraphLocation(sourceIndex),
                sourceIndex,
                null,
                sourceParagraphs[sourceIndex].Text,
                null,
                "Paragraph deleted."),
                sourceParagraphs[sourceIndex].DocumentOrder);
        }

        private static List<ParagraphSnapshot> GetLogicalBodyParagraphs(WordDocument document, WordComparisonOptions options) {
            var snapshots = new List<ParagraphSnapshot>();
            MainDocumentPart? mainPart = document._wordprocessingDocument.MainDocumentPart;
            AddParagraphSnapshots(snapshots, mainPart, mainPart?.Document?.Body, BodyPartKey, BodyPartOrderBase, options);

            if (mainPart != null) {
                int headerIndex = 0;
                foreach (KeyValuePair<HeaderPart, string> headerPartKey in CreateOrderedHeaderPartKeys(mainPart)) {
                    AddParagraphSnapshots(snapshots, headerPartKey.Key, headerPartKey.Key.Header, headerPartKey.Value, HeaderPartOrderBase + (headerIndex * RelatedPartOrderStride), options);
                    headerIndex++;
                }

                int footerIndex = 0;
                foreach (KeyValuePair<FooterPart, string> footerPartKey in CreateOrderedFooterPartKeys(mainPart)) {
                    AddParagraphSnapshots(snapshots, footerPartKey.Key, footerPartKey.Key.Footer, footerPartKey.Value, FooterPartOrderBase + (footerIndex * RelatedPartOrderStride), options);
                    footerIndex++;
                }

                List<Footnote> footnotes = GetReferencedFootnotes(mainPart);
                for (int footnoteIndex = 0; footnoteIndex < footnotes.Count; footnoteIndex++) {
                    string noteId = footnotes[footnoteIndex].Id?.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) ??
                        footnoteIndex.ToString(System.Globalization.CultureInfo.InvariantCulture);
                    AddParagraphSnapshots(snapshots, mainPart.FootnotesPart, footnotes[footnoteIndex], FootnotePartKeyPrefix + noteId, FootnotePartOrderBase + (footnoteIndex * RelatedPartOrderStride), options);
                }

                List<Endnote> endnotes = GetReferencedEndnotes(mainPart);
                for (int endnoteIndex = 0; endnoteIndex < endnotes.Count; endnoteIndex++) {
                    string noteId = endnotes[endnoteIndex].Id?.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) ??
                        endnoteIndex.ToString(System.Globalization.CultureInfo.InvariantCulture);
                    AddParagraphSnapshots(snapshots, mainPart.EndnotesPart, endnotes[endnoteIndex], EndnotePartKeyPrefix + noteId, EndnotePartOrderBase + (endnoteIndex * RelatedPartOrderStride), options);
                }
            }

            return snapshots;
        }

        private static void AddParagraphSnapshots(List<ParagraphSnapshot> snapshots, OpenXmlPart? part, OpenXmlElement? container, string partKind, int orderBase, WordComparisonOptions options) {
            foreach (OrderedElement ordered in EnumerateDescendantsWithOrder(container, orderBase)) {
                if (ordered.Element is not Paragraph paragraph) {
                    continue;
                }

                if (paragraph.Ancestors<TableCell>().Any()) {
                    continue;
                }

                ParagraphTextSnapshot text = GetParagraphTextSnapshot(paragraph, part, options);
                if (text.Text.Length == 0 && HasImageContent(paragraph)) {
                    continue;
                }

                snapshots.Add(new ParagraphSnapshot(
                    partKind,
                    part,
                    paragraph,
                    text.Text,
                    text.ComparisonText,
                    text.MatchText,
                    GetParagraphStyleId(paragraph),
                    GetParagraphFormatSignature(paragraph, part, options),
                    ordered.DocumentOrder));
            }
        }

        private static string GetParagraphStyleId(Paragraph paragraph) {
            return paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value ?? string.Empty;
        }

        private static string GetParagraphText(Paragraph paragraph) {
            return GetParagraphTextSnapshot(paragraph, null).Text;
        }

        private static string GetParagraphMatchText(Paragraph paragraph) {
            return GetParagraphTextSnapshot(paragraph, null).MatchText;
        }

        private static string GetParagraphMatchText(Paragraph paragraph, OpenXmlPart? part) {
            return GetParagraphTextSnapshot(paragraph, part).MatchText;
        }

        private static string GetParagraphMatchText(Paragraph paragraph, OpenXmlPart? part, WordComparisonOptions options) {
            return GetParagraphTextSnapshot(paragraph, part, options).MatchText;
        }

        private static ParagraphTextSnapshot GetParagraphTextSnapshot(Paragraph paragraph, OpenXmlPart? part) {
            return GetParagraphTextSnapshot(paragraph, part, WordComparisonOptions.Default);
        }

        private static ParagraphTextSnapshot GetParagraphTextSnapshot(Paragraph paragraph, OpenXmlPart? part, WordComparisonOptions options) {
            var textBuilder = new StringBuilder();
            var matchBuilder = new StringBuilder();
            var pendingTextBuilder = new StringBuilder();
            AppendNumberingMatchToken(matchBuilder, paragraph, part, options);
            foreach (OpenXmlElement element in EnumerateComparableDescendants(paragraph)) {
                if (element.Ancestors<Paragraph>().FirstOrDefault() != paragraph) {
                    continue;
                }

                switch (element) {
                    case Text text:
                        textBuilder.Append(text.Text);
                        pendingTextBuilder.Append(text.Text);
                        break;
                    case DeletedText deletedText:
                        FlushPendingTextToken(matchBuilder, pendingTextBuilder, options);
                        string deletedValue = deletedText.Text ?? string.Empty;
                        textBuilder.Append("[Deleted:");
                        textBuilder.Append(deletedValue);
                        textBuilder.Append(']');
                        AppendMatchToken(matchBuilder, "deletedText", NormalizeComparisonText(deletedValue, options));
                        break;
                    case Hyperlink hyperlink:
                        FlushPendingTextToken(matchBuilder, pendingTextBuilder, options);
                        AppendMatchToken(matchBuilder, "hyperlink", NormalizeComparisonText(GetHyperlinkSignature(part, hyperlink), options));
                        break;
                    case SimpleField simpleField:
                        FlushPendingTextToken(matchBuilder, pendingTextBuilder, options);
                        AppendMatchToken(matchBuilder, "simpleField", NormalizeComparisonText(simpleField.Instruction?.Value ?? string.Empty, options));
                        break;
                    case FieldCode fieldCode:
                        FlushPendingTextToken(matchBuilder, pendingTextBuilder, options);
                        AppendMatchToken(matchBuilder, "fieldCode", NormalizeComparisonText(fieldCode.Text ?? string.Empty, options));
                        break;
                    case TabChar:
                        FlushPendingTextToken(matchBuilder, pendingTextBuilder, options);
                        textBuilder.Append('\t');
                        AppendMatchToken(matchBuilder, "tab", string.Empty);
                        break;
                    case Break breakNode:
                        FlushPendingTextToken(matchBuilder, pendingTextBuilder, options);
                        if (breakNode.Type == null || breakNode.Type.Value == BreakValues.TextWrapping) {
                            textBuilder.Append('\n');
                            AppendMatchToken(matchBuilder, "break", "textWrapping");
                        } else if (breakNode.Type.Value == BreakValues.Page) {
                            textBuilder.Append("[PageBreak]");
                            AppendMatchToken(matchBuilder, "break", "page");
                        } else if (breakNode.Type.Value == BreakValues.Column) {
                            textBuilder.Append("[ColumnBreak]");
                            AppendMatchToken(matchBuilder, "break", "column");
                        } else {
                            string breakType = breakNode.Type.Value.ToString();
                            textBuilder.Append("[Break:");
                            textBuilder.Append(breakType);
                            textBuilder.Append(']');
                            AppendMatchToken(matchBuilder, "break", breakType);
                        }

                        break;
                    case FootnoteReference footnoteReference:
                        FlushPendingTextToken(matchBuilder, pendingTextBuilder, options);
                        string footnoteReferenceId = footnoteReference.Id?.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty;
                        textBuilder.Append("[FootnoteReference:");
                        textBuilder.Append(footnoteReferenceId);
                        textBuilder.Append(']');
                        AppendMatchToken(matchBuilder, "footnoteReference", NormalizeComparisonText(GetFootnoteReferenceSignature(part, footnoteReference.Id?.Value, options), options));
                        break;
                    case EndnoteReference endnoteReference:
                        FlushPendingTextToken(matchBuilder, pendingTextBuilder, options);
                        string endnoteReferenceId = endnoteReference.Id?.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty;
                        textBuilder.Append("[EndnoteReference:");
                        textBuilder.Append(endnoteReferenceId);
                        textBuilder.Append(']');
                        AppendMatchToken(matchBuilder, "endnoteReference", NormalizeComparisonText(GetEndnoteReferenceSignature(part, endnoteReference.Id?.Value, options), options));
                        break;
                    case SymbolChar symbol:
                        FlushPendingTextToken(matchBuilder, pendingTextBuilder, options);
                        string symbolFont = symbol.Font?.Value ?? string.Empty;
                        string symbolChar = symbol.Char?.Value ?? string.Empty;
                        textBuilder.Append("[Symbol:");
                        textBuilder.Append(symbolFont);
                        textBuilder.Append(':');
                        textBuilder.Append(symbolChar);
                        textBuilder.Append(']');
                        AppendMatchToken(matchBuilder, "symbol", symbolFont + ":" + symbolChar);
                        break;
                    case NoBreakHyphen:
                        FlushPendingTextToken(matchBuilder, pendingTextBuilder, options);
                        textBuilder.Append('-');
                        AppendMatchToken(matchBuilder, "noBreakHyphen", string.Empty);
                        break;
                    case SoftHyphen:
                        FlushPendingTextToken(matchBuilder, pendingTextBuilder, options);
                        textBuilder.Append("[SoftHyphen]");
                        AppendMatchToken(matchBuilder, "softHyphen", string.Empty);
                        break;
                    case CarriageReturn:
                        FlushPendingTextToken(matchBuilder, pendingTextBuilder, options);
                        textBuilder.Append('\n');
                        AppendMatchToken(matchBuilder, "carriageReturn", string.Empty);
                        break;
                }
            }

            FlushPendingTextToken(matchBuilder, pendingTextBuilder, options);
            string paragraphText = textBuilder.ToString();
            return new ParagraphTextSnapshot(paragraphText, NormalizeComparisonText(paragraphText, options), matchBuilder.ToString());
        }

        private static void AppendNumberingMatchToken(StringBuilder builder, Paragraph paragraph, OpenXmlPart? part, WordComparisonOptions options) {
            NumberingProperties? numberingProperties = paragraph.ParagraphProperties?.NumberingProperties;
            if (numberingProperties == null) {
                return;
            }

            string level = numberingProperties.NumberingLevelReference?.Val?.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty;
            int? numberingId = numberingProperties.NumberingId?.Val?.Value;
            AppendMatchToken(builder, "numbering", "level=" + level + ";definition=" + NormalizeComparisonText(GetNumberingDefinitionSignature(part, numberingId, level), options));
        }

        private static string GetNumberingDefinitionSignature(OpenXmlPart? part, int? numberingId, string level) {
            if (numberingId == null) {
                return string.Empty;
            }

            MainDocumentPart? mainPart = GetMainDocumentPart(part);
            Numbering? numbering = mainPart?.NumberingDefinitionsPart?.Numbering;
            if (numbering == null) {
                return numberingId.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
            }

            NumberingInstance? instance = numbering.Elements<NumberingInstance>()
                .FirstOrDefault(item => item.NumberID?.Value == numberingId.Value);
            if (instance == null) {
                return numberingId.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
            }

            int? abstractNumberId = instance?.AbstractNumId?.Val?.Value;
            if (abstractNumberId == null) {
                return numberingId.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
            }

            AbstractNum? abstractNum = numbering.Elements<AbstractNum>()
                .FirstOrDefault(item => item.AbstractNumberId?.Value == abstractNumberId.Value);
            if (abstractNum == null) {
                return abstractNumberId.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
            }

            OpenXmlElement clone = abstractNum.CloneNode(true);
            if (clone is AbstractNum clonedAbstract) {
                clonedAbstract.AbstractNumberId = 0;
            }

            OpenXmlElement instanceClone = instance!.CloneNode(true);
            if (instanceClone is NumberingInstance clonedInstance) {
                clonedInstance.NumberID = 0;
                if (clonedInstance.AbstractNumId != null) {
                    clonedInstance.AbstractNumId.Val = 0;
                }
            }

            return clone.OuterXml + "|instance:" + instanceClone.OuterXml + "|level:" + level;
        }

        private static MainDocumentPart? GetMainDocumentPart(OpenXmlPart? part) {
            if (part is MainDocumentPart mainPart) {
                return mainPart;
            }

            return (part?.OpenXmlPackage as WordprocessingDocument)?.MainDocumentPart;
        }

        private static void FlushPendingTextToken(StringBuilder builder, StringBuilder pendingTextBuilder, WordComparisonOptions options) {
            if (pendingTextBuilder.Length == 0) {
                return;
            }

            AppendMatchToken(builder, "text", NormalizeComparisonText(pendingTextBuilder.ToString(), options));
            pendingTextBuilder.Clear();
        }

        private static void AppendMatchToken(StringBuilder builder, string kind, string value) {
            builder.Append(kind);
            builder.Append(':');
            builder.Append(value.Length.ToString(System.Globalization.CultureInfo.InvariantCulture));
            builder.Append(':');
            builder.Append(value);
            builder.Append(';');
        }

        private static string GetHyperlinkSignature(OpenXmlPart? part, Hyperlink hyperlink) {
            var builder = new StringBuilder();
            if (hyperlink.Id?.Value is string relationshipId) {
                builder.Append("id:");
                builder.Append(GetRelationshipTarget(part, relationshipId));
            }

            if (hyperlink.Anchor?.Value is string anchor) {
                builder.Append("#");
                builder.Append(anchor);
            }

            if (hyperlink.DocLocation?.Value is string docLocation) {
                builder.Append("@");
                builder.Append(docLocation);
            }

            if (hyperlink.History?.Value is bool history) {
                builder.Append("|history=");
                builder.Append(history ? "true" : "false");
            }

            return builder.ToString();
        }

        private static string GetRelationshipTarget(OpenXmlPart? part, string relationshipId) {
            if (part == null) {
                return relationshipId;
            }

            HyperlinkRelationship? hyperlinkRelationship = part.HyperlinkRelationships.FirstOrDefault(item => item.Id == relationshipId);
            if (hyperlinkRelationship != null) {
                return hyperlinkRelationship.Uri.ToString();
            }

            ExternalRelationship? externalRelationship = part.ExternalRelationships.FirstOrDefault(item => item.Id == relationshipId);
            if (externalRelationship != null) {
                return externalRelationship.Uri.ToString();
            }

            return relationshipId;
        }

        private static string GetFootnoteReferenceSignature(OpenXmlPart? part, long? noteId) {
            return GetFootnoteReferenceSignature(part, noteId, WordComparisonOptions.Default);
        }

        private static string GetFootnoteReferenceSignature(OpenXmlPart? part, long? noteId, WordComparisonOptions options) {
            if (part is not MainDocumentPart mainPart || noteId == null) {
                return noteId?.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty;
            }

            Footnote? footnote = mainPart.FootnotesPart?.Footnotes?.Elements<Footnote>()
                .FirstOrDefault(item => item.Id?.Value == noteId.Value && IsVisibleNote(item));
            return footnote == null ? string.Empty : GetNoteContentSignature(mainPart.FootnotesPart, footnote, options);
        }

        private static string GetEndnoteReferenceSignature(OpenXmlPart? part, long? noteId) {
            return GetEndnoteReferenceSignature(part, noteId, WordComparisonOptions.Default);
        }

        private static string GetEndnoteReferenceSignature(OpenXmlPart? part, long? noteId, WordComparisonOptions options) {
            if (part is not MainDocumentPart mainPart || noteId == null) {
                return noteId?.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty;
            }

            Endnote? endnote = mainPart.EndnotesPart?.Endnotes?.Elements<Endnote>()
                .FirstOrDefault(item => item.Id?.Value == noteId.Value && IsVisibleNote(item));
            return endnote == null ? string.Empty : GetNoteContentSignature(mainPart.EndnotesPart, endnote, options);
        }

        private static string GetNoteContentSignature(OpenXmlPart? notePart, OpenXmlElement note) {
            return GetNoteContentSignature(notePart, note, WordComparisonOptions.Default);
        }

        private static string GetNoteContentSignature(OpenXmlPart? notePart, OpenXmlElement note, WordComparisonOptions options) {
            return string.Join(
                TableRowSeparator,
                note.Elements<Paragraph>()
                    .Select(paragraph => GetParagraphMatchText(paragraph, notePart, options))
                    .ToArray());
        }

        private static bool HasImageContent(Paragraph paragraph) {
            return paragraph.Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>().Any() ||
                   paragraph.Descendants<V.ImageData>().Any();
        }

        private static double GetContainmentAwareTextSimilarity(string source, string target) {
            if (source.Length > 0 &&
                target.Length > 0 &&
                (source.IndexOf(target, StringComparison.Ordinal) >= 0 ||
                 target.IndexOf(source, StringComparison.Ordinal) >= 0)) {
                return 0.75 + (0.25 * Math.Min(source.Length, target.Length) / Math.Max(source.Length, target.Length));
            }

            return GetTextSimilarity(source, target);
        }

        private static double GetTextSimilarity(string source, string target) {
            if (string.Equals(source, target, StringComparison.Ordinal)) {
                return 1;
            }

            if (source.Length == 0 || target.Length == 0) {
                return 0;
            }

            if ((long)(source.Length + 1) * (target.Length + 1) > LcsCellLimit) {
                return GetBoundedTextSimilarity(source, target);
            }

            return (double)GetCommonSubsequenceLength(source, target) / Math.Max(source.Length, target.Length);
        }

        private static bool AreComparisonTextEqual(string? source, string? target, WordComparisonOptions options) =>
            string.Equals(NormalizeComparisonText(source ?? string.Empty, options), NormalizeComparisonText(target ?? string.Empty, options), StringComparison.Ordinal);

        private static string NormalizeComparisonText(string value, WordComparisonOptions options) {
            string normalized = value ?? string.Empty;
            if (options.IgnoreWhitespace) {
                normalized = string.Join(" ", normalized.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
            }

            if (options.IgnoreCase) {
                normalized = normalized.ToUpperInvariant();
            }

            return normalized;
        }

        private static double GetBoundedTextSimilarity(string source, string target) {
            int prefixLength = 0;
            int maxPrefixLength = Math.Min(source.Length, target.Length);
            while (prefixLength < maxPrefixLength && source[prefixLength] == target[prefixLength]) {
                prefixLength++;
            }

            int suffixLength = 0;
            int sourceSuffixIndex = source.Length - 1;
            int targetSuffixIndex = target.Length - 1;
            while (sourceSuffixIndex >= prefixLength &&
                   targetSuffixIndex >= prefixLength &&
                   source[sourceSuffixIndex] == target[targetSuffixIndex]) {
                suffixLength++;
                sourceSuffixIndex--;
                targetSuffixIndex--;
            }

            return (double)(prefixLength + suffixLength) / Math.Max(source.Length, target.Length);
        }

        private static int GetCommonSubsequenceLength(string source, string target) {
            int[,] lengths = new int[source.Length + 1, target.Length + 1];

            for (int sourceIndex = source.Length - 1; sourceIndex >= 0; sourceIndex--) {
                for (int targetIndex = target.Length - 1; targetIndex >= 0; targetIndex--) {
                    lengths[sourceIndex, targetIndex] = source[sourceIndex] == target[targetIndex]
                        ? lengths[sourceIndex + 1, targetIndex + 1] + 1
                        : Math.Max(lengths[sourceIndex + 1, targetIndex], lengths[sourceIndex, targetIndex + 1]);
                }
            }

            return lengths[0, 0];
        }

        private static string ParagraphLocation(int paragraphIndex) {
            return "paragraph[" + paragraphIndex.ToString(System.Globalization.CultureInfo.InvariantCulture) + "]";
        }

        private sealed class ParagraphSnapshot {
            internal ParagraphSnapshot(string partKind, OpenXmlPart? part, Paragraph paragraph, string text, string comparisonText, string matchText, string styleId, string formatSignature, int documentOrder) {
                PartKind = partKind;
                Part = part;
                Paragraph = paragraph;
                Text = text;
                ComparisonText = comparisonText;
                MatchText = matchText;
                StyleId = styleId;
                FormatSignature = formatSignature;
                DocumentOrder = documentOrder;
            }

            internal string PartKind { get; }

            internal OpenXmlPart? Part { get; }

            internal Paragraph Paragraph { get; }

            internal string Text { get; }

            internal string ComparisonText { get; }

            internal string MatchText { get; }

            internal string StyleId { get; }

            internal string FormatSignature { get; }

            internal int DocumentOrder { get; }
        }

        private readonly struct ParagraphTextSnapshot {
            internal ParagraphTextSnapshot(string text, string comparisonText, string matchText) {
                Text = text;
                ComparisonText = comparisonText;
                MatchText = matchText;
            }

            internal string Text { get; }

            internal string ComparisonText { get; }

            internal string MatchText { get; }
        }

        private sealed class ParagraphSnapshotEqualityComparer : IEqualityComparer<ParagraphSnapshot> {
            internal static readonly ParagraphSnapshotEqualityComparer Instance = new();

            public bool Equals(ParagraphSnapshot? x, ParagraphSnapshot? y) {
                if (ReferenceEquals(x, y)) {
                    return true;
                }

                if (x == null || y == null) {
                    return false;
                }

                return string.Equals(x.PartKind, y.PartKind, StringComparison.Ordinal) &&
                       string.Equals(x.MatchText, y.MatchText, StringComparison.Ordinal);
            }

            public int GetHashCode(ParagraphSnapshot obj) {
                unchecked {
                    return (StringComparer.Ordinal.GetHashCode(obj.PartKind) * 397) ^
                           StringComparer.Ordinal.GetHashCode(obj.MatchText);
                }
            }
        }
    }
}
