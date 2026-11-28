using System.Text;
using OfficeIMO.Word.OpenXml;

namespace OfficeIMO.Word {
    public static partial class WordDocumentComparer {
        private static void AnalyzeParagraphs(
            WordPackageView source,
            WordPackageView target,
            WordComparisonResult result,
            WordComparisonOptions options,
            ComparisonWorkBudget comparisonWorkBudget) {
            List<ParagraphSnapshot> sourceParagraphs = GetLogicalBodyParagraphs(source, options);
            List<ParagraphSnapshot> targetParagraphs = GetLogicalBodyParagraphs(target, options);
            IReadOnlyList<MatchedIndexPair> matchedParagraphs = FindMatchingIndexes(
                sourceParagraphs,
                targetParagraphs,
                ParagraphSnapshotEqualityComparer.Instance);

            int sourceStart = 0;
            int targetStart = 0;

            foreach (MatchedIndexPair match in matchedParagraphs) {
                AddParagraphRangeFindings(sourceParagraphs, targetParagraphs, sourceStart, match.SourceIndex, targetStart, match.TargetIndex, result, options, comparisonWorkBudget);
                AnalyzeParagraphStyle(sourceParagraphs[match.SourceIndex], targetParagraphs[match.TargetIndex], match.SourceIndex, match.TargetIndex, result, options);
                AnalyzeParagraphEffectiveFormatting(sourceParagraphs[match.SourceIndex], targetParagraphs[match.TargetIndex], match.SourceIndex, match.TargetIndex, result, options);
                AnalyzeParagraphRuns(sourceParagraphs[match.SourceIndex], targetParagraphs[match.TargetIndex], match.SourceIndex, match.TargetIndex, result, options);
                sourceStart = match.SourceIndex + 1;
                targetStart = match.TargetIndex + 1;
            }

            AddParagraphRangeFindings(sourceParagraphs, targetParagraphs, sourceStart, sourceParagraphs.Count, targetStart, targetParagraphs.Count, result, options, comparisonWorkBudget);
        }

        private static void AddParagraphRangeFindings(
            IReadOnlyList<ParagraphSnapshot> sourceParagraphs,
            IReadOnlyList<ParagraphSnapshot> targetParagraphs,
            int sourceStart,
            int sourceEnd,
            int targetStart,
            int targetEnd,
            WordComparisonResult result,
            WordComparisonOptions options,
            ComparisonWorkBudget comparisonWorkBudget) {
            int sourceIndex = sourceStart;
            int targetIndex = targetStart;

            while (sourceIndex < sourceEnd && targetIndex < targetEnd) {
                int betterTargetIndex = FindBetterTargetAlignmentIndex(sourceParagraphs[sourceIndex], targetParagraphs, targetIndex, targetEnd, comparisonWorkBudget);
                if (betterTargetIndex > targetIndex) {
                    while (targetIndex < betterTargetIndex) {
                        AddInsertedParagraphFinding(targetParagraphs, targetIndex, result);
                        targetIndex++;
                    }

                    continue;
                }

                int betterSourceIndex = FindBetterSourceAlignmentIndex(sourceParagraphs, sourceIndex, sourceEnd, targetParagraphs[targetIndex], comparisonWorkBudget);
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

        private static int FindBetterTargetAlignmentIndex(
            ParagraphSnapshot sourceParagraph,
            IReadOnlyList<ParagraphSnapshot> targetParagraphs,
            int targetStart,
            int targetEnd,
            ComparisonWorkBudget comparisonWorkBudget) {
            int bestIndex = targetStart;
            double currentVisibleSimilarity = GetParagraphVisibleTextSimilarity(sourceParagraph, targetParagraphs[targetStart], comparisonWorkBudget);
            double bestVisibleSimilarity = currentVisibleSimilarity;
            double currentSimilarity = GetParagraphSimilarity(sourceParagraph, targetParagraphs[targetStart], comparisonWorkBudget);
            double bestSimilarity = currentSimilarity;

            foreach (int index in SelectBoundedAlignmentCandidates(
                targetStart + 1,
                targetEnd,
                index => GetParagraphAlignmentPrefilterSimilarity(sourceParagraph, targetParagraphs[index], comparisonWorkBudget),
                comparisonWorkBudget)) {
                double visibleSimilarity = GetParagraphVisibleTextSimilarity(sourceParagraph, targetParagraphs[index], comparisonWorkBudget);
                double similarity = GetParagraphSimilarity(sourceParagraph, targetParagraphs[index], comparisonWorkBudget);
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

            return bestVisibleSimilarity > currentVisibleSimilarity || bestSimilarity > currentSimilarity
                ? bestIndex
                : targetStart;
        }

        private static int FindBetterSourceAlignmentIndex(
            IReadOnlyList<ParagraphSnapshot> sourceParagraphs,
            int sourceStart,
            int sourceEnd,
            ParagraphSnapshot targetParagraph,
            ComparisonWorkBudget comparisonWorkBudget) {
            int bestIndex = sourceStart;
            double currentVisibleSimilarity = GetParagraphVisibleTextSimilarity(sourceParagraphs[sourceStart], targetParagraph, comparisonWorkBudget);
            double bestVisibleSimilarity = currentVisibleSimilarity;
            double currentSimilarity = GetParagraphSimilarity(sourceParagraphs[sourceStart], targetParagraph, comparisonWorkBudget);
            double bestSimilarity = currentSimilarity;

            foreach (int index in SelectBoundedAlignmentCandidates(
                sourceStart + 1,
                sourceEnd,
                index => GetParagraphAlignmentPrefilterSimilarity(sourceParagraphs[index], targetParagraph, comparisonWorkBudget),
                comparisonWorkBudget)) {
                double visibleSimilarity = GetParagraphVisibleTextSimilarity(sourceParagraphs[index], targetParagraph, comparisonWorkBudget);
                double similarity = GetParagraphSimilarity(sourceParagraphs[index], targetParagraph, comparisonWorkBudget);
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

            return bestVisibleSimilarity > currentVisibleSimilarity || bestSimilarity > currentSimilarity
                ? bestIndex
                : sourceStart;
        }

        private static double GetParagraphSimilarity(
            ParagraphSnapshot sourceParagraph,
            ParagraphSnapshot targetParagraph,
            ComparisonWorkBudget comparisonWorkBudget) {
            if (!string.Equals(sourceParagraph.PartKind, targetParagraph.PartKind, StringComparison.Ordinal)) {
                return 0;
            }

            return Math.Max(
                GetTextSimilarity(sourceParagraph.MatchText, targetParagraph.MatchText, comparisonWorkBudget),
                GetTextSimilarity(sourceParagraph.ComparisonText, targetParagraph.ComparisonText, comparisonWorkBudget));
        }

        private static double GetParagraphAlignmentPrefilterSimilarity(
            ParagraphSnapshot sourceParagraph,
            ParagraphSnapshot targetParagraph,
            ComparisonWorkBudget comparisonWorkBudget) {
            if (!string.Equals(sourceParagraph.PartKind, targetParagraph.PartKind, StringComparison.Ordinal)) {
                return 0;
            }

            return Math.Max(
                GetBoundedTextSimilarity(sourceParagraph.MatchText, targetParagraph.MatchText, comparisonWorkBudget),
                GetBoundedTextSimilarity(sourceParagraph.ComparisonText, targetParagraph.ComparisonText, comparisonWorkBudget));
        }

        private static IReadOnlyList<int> SelectBoundedAlignmentCandidates(
            int start,
            int end,
            Func<int, double> getPrefilterSimilarity,
            ComparisonWorkBudget comparisonWorkBudget) {
            var candidates = new SortedSet<AlignmentCandidate>(AlignmentCandidateComparer.Instance);
            for (int index = start; index < end; index++) {
                if (!comparisonWorkBudget.TryConsume(1)) {
                    break;
                }

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

        private static double GetParagraphVisibleTextSimilarity(
            ParagraphSnapshot sourceParagraph,
            ParagraphSnapshot targetParagraph,
            ComparisonWorkBudget comparisonWorkBudget) {
            if (!string.Equals(sourceParagraph.PartKind, targetParagraph.PartKind, StringComparison.Ordinal)) {
                return 0;
            }

            return GetContainmentAwareTextSimilarity(sourceParagraph.ComparisonText, targetParagraph.ComparisonText, comparisonWorkBudget);
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

        private static List<ParagraphSnapshot> GetLogicalBodyParagraphs(WordPackageView document, WordComparisonOptions options) {
            var snapshots = new List<ParagraphSnapshot>();
            WordPartView? mainPart = document.Main;
            AddParagraphSnapshots(snapshots, mainPart, BodyOf(mainPart), BodyPartKey, BodyPartOrderBase, options);

            if (mainPart != null) {
                int headerIndex = 0;
                foreach (KeyValuePair<WordPartView, string> headerPartKey in CreateOrderedHeaderPartKeys(mainPart)) {
                    AddParagraphSnapshots(snapshots, headerPartKey.Key, HeaderRoot(headerPartKey.Key), headerPartKey.Value, HeaderPartOrderBase + (headerIndex * RelatedPartOrderStride), options);
                    headerIndex++;
                }

                int footerIndex = 0;
                foreach (KeyValuePair<WordPartView, string> footerPartKey in CreateOrderedFooterPartKeys(mainPart)) {
                    AddParagraphSnapshots(snapshots, footerPartKey.Key, FooterRoot(footerPartKey.Key), footerPartKey.Value, FooterPartOrderBase + (footerIndex * RelatedPartOrderStride), options);
                    footerIndex++;
                }

                List<OfficeXmlNode> footnotes = GetReferencedFootnotes(mainPart);
                for (int footnoteIndex = 0; footnoteIndex < footnotes.Count; footnoteIndex++) {
                    string noteId = GetNotePartKeyId(footnotes[footnoteIndex], footnoteIndex);
                    AddParagraphSnapshots(snapshots, FootnotesPart(mainPart), footnotes[footnoteIndex], FootnotePartKeyPrefix + noteId, FootnotePartOrderBase + (footnoteIndex * RelatedPartOrderStride), options);
                }

                List<OfficeXmlNode> endnotes = GetReferencedEndnotes(mainPart);
                for (int endnoteIndex = 0; endnoteIndex < endnotes.Count; endnoteIndex++) {
                    string noteId = GetNotePartKeyId(endnotes[endnoteIndex], endnoteIndex);
                    AddParagraphSnapshots(snapshots, EndnotesPart(mainPart), endnotes[endnoteIndex], EndnotePartKeyPrefix + noteId, EndnotePartOrderBase + (endnoteIndex * RelatedPartOrderStride), options);
                }
            }

            return snapshots;
        }

        private static void AddParagraphSnapshots(List<ParagraphSnapshot> snapshots, WordPartView? part, OfficeXmlNode container, string partKind, int orderBase, WordComparisonOptions options) {
            foreach (OrderedElement ordered in EnumerateDescendantsWithOrder(container, orderBase)) {
                OfficeXmlNode paragraph = ordered.Element;
                if (!paragraph.Is(W, "p")) {
                    continue;
                }

                if (HasAncestorW(paragraph, "tc")) {
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

        private static string GetParagraphStyleId(OfficeXmlNode paragraph) {
            return ChildVal(paragraph.Child(W, "pPr"), "pStyle") ?? string.Empty;
        }

        private static string GetParagraphText(OfficeXmlNode paragraph) {
            return GetParagraphTextSnapshot(paragraph, null).Text;
        }

        private static string GetParagraphMatchText(OfficeXmlNode paragraph) {
            return GetParagraphTextSnapshot(paragraph, null).MatchText;
        }

        private static string GetParagraphMatchText(OfficeXmlNode paragraph, WordPartView? part) {
            return GetParagraphTextSnapshot(paragraph, part).MatchText;
        }

        private static string GetParagraphMatchText(OfficeXmlNode paragraph, WordPartView? part, WordComparisonOptions options) {
            return GetParagraphTextSnapshot(paragraph, part, options).MatchText;
        }

        private static ParagraphTextSnapshot GetParagraphTextSnapshot(OfficeXmlNode paragraph, WordPartView? part) {
            return GetParagraphTextSnapshot(paragraph, part, WordComparisonOptions.Default);
        }

        /// <summary>
        /// The <c>w:type</c> of a break as the SDK path read it: "textWrapping" when absent, a format error for a value
        /// the SDK doesn't parse.
        /// </summary>
        private static string GetBreakType(OfficeXmlNode breakNode) =>
            EnumToken(breakNode.Attribute(W, "type"), "page", "column", "textWrapping") ?? "textWrapping";

        private static ParagraphTextSnapshot GetParagraphTextSnapshot(OfficeXmlNode paragraph, WordPartView? part, WordComparisonOptions options) {
            var textBuilder = new StringBuilder();
            var matchBuilder = new StringBuilder();
            var pendingTextBuilder = new StringBuilder();
            AppendNumberingMatchToken(matchBuilder, paragraph, part, options);
            foreach (OfficeXmlNode element in EnumerateComparableDescendants(paragraph)) {
                if (element.NamespaceId != W || AncestorW(element, "p") != paragraph) {
                    continue;
                }

                switch (element.LocalName) {
                    case "t":
                        string text = element.InnerText;
                        textBuilder.Append(text);
                        pendingTextBuilder.Append(text);
                        break;
                    case "delText":
                        FlushPendingTextToken(matchBuilder, pendingTextBuilder, options);
                        string deletedValue = element.InnerText;
                        textBuilder.Append("[Deleted:");
                        textBuilder.Append(deletedValue);
                        textBuilder.Append(']');
                        AppendMatchToken(matchBuilder, "deletedText", NormalizeComparisonText(deletedValue, options));
                        break;
                    case "hyperlink":
                        FlushPendingTextToken(matchBuilder, pendingTextBuilder, options);
                        AppendMatchToken(matchBuilder, "hyperlink", NormalizeComparisonText(GetHyperlinkSignature(part, element), options));
                        break;
                    case "fldSimple":
                        FlushPendingTextToken(matchBuilder, pendingTextBuilder, options);
                        AppendMatchToken(matchBuilder, "simpleField", NormalizeComparisonText(element.Attribute(W, "instr") ?? string.Empty, options));
                        break;
                    case "instrText":
                        FlushPendingTextToken(matchBuilder, pendingTextBuilder, options);
                        AppendMatchToken(matchBuilder, "fieldCode", NormalizeComparisonText(element.InnerText, options));
                        break;
                    case "tab" when IsTabChar(element):
                        FlushPendingTextToken(matchBuilder, pendingTextBuilder, options);
                        textBuilder.Append('\t');
                        AppendMatchToken(matchBuilder, "tab", string.Empty);
                        break;
                    case "br":
                        FlushPendingTextToken(matchBuilder, pendingTextBuilder, options);
                        switch (GetBreakType(element)) {
                            case "page":
                                textBuilder.Append("[PageBreak]");
                                AppendMatchToken(matchBuilder, "break", "page");
                                break;
                            case "column":
                                textBuilder.Append("[ColumnBreak]");
                                AppendMatchToken(matchBuilder, "break", "column");
                                break;
                            default:
                                textBuilder.Append('\n');
                                AppendMatchToken(matchBuilder, "break", "textWrapping");
                                break;
                        }

                        break;
                    case "footnoteReference":
                        FlushPendingTextToken(matchBuilder, pendingTextBuilder, options);
                        long? footnoteId = GetNoteId(element);
                        textBuilder.Append("[FootnoteReference:");
                        textBuilder.Append(footnoteId?.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty);
                        textBuilder.Append(']');
                        AppendMatchToken(matchBuilder, "footnoteReference", NormalizeComparisonText(GetFootnoteReferenceSignature(part, footnoteId, options), options));
                        break;
                    case "endnoteReference":
                        FlushPendingTextToken(matchBuilder, pendingTextBuilder, options);
                        long? endnoteId = GetNoteId(element);
                        textBuilder.Append("[EndnoteReference:");
                        textBuilder.Append(endnoteId?.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty);
                        textBuilder.Append(']');
                        AppendMatchToken(matchBuilder, "endnoteReference", NormalizeComparisonText(GetEndnoteReferenceSignature(part, endnoteId, options), options));
                        break;
                    case "sym":
                        FlushPendingTextToken(matchBuilder, pendingTextBuilder, options);
                        string symbolFont = element.Attribute(W, "font") ?? string.Empty;
                        string symbolChar = element.Attribute(W, "char") ?? string.Empty;
                        textBuilder.Append("[Symbol:");
                        textBuilder.Append(symbolFont);
                        textBuilder.Append(':');
                        textBuilder.Append(symbolChar);
                        textBuilder.Append(']');
                        AppendMatchToken(matchBuilder, "symbol", symbolFont + ":" + symbolChar);
                        break;
                    case "noBreakHyphen":
                        FlushPendingTextToken(matchBuilder, pendingTextBuilder, options);
                        textBuilder.Append('-');
                        AppendMatchToken(matchBuilder, "noBreakHyphen", string.Empty);
                        break;
                    case "softHyphen":
                        FlushPendingTextToken(matchBuilder, pendingTextBuilder, options);
                        textBuilder.Append("[SoftHyphen]");
                        AppendMatchToken(matchBuilder, "softHyphen", string.Empty);
                        break;
                    case "cr":
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

        private static void AppendNumberingMatchToken(StringBuilder builder, OfficeXmlNode paragraph, WordPartView? part, WordComparisonOptions options) {
            OfficeXmlNode numberingProperties = paragraph.Child(W, "pPr").Child(W, "numPr");
            if (numberingProperties.IsNone) {
                return;
            }

            string level = Int32Text(ChildVal(numberingProperties, "ilvl")) ?? string.Empty;
            int? numberingId = ParseInt32(ChildVal(numberingProperties, "numId"));
            AppendMatchToken(builder, "numbering", "level=" + level + ";definition=" + NormalizeComparisonText(GetNumberingDefinitionSignature(part, numberingId, level), options));
        }

        private static string GetNumberingDefinitionSignature(WordPartView? part, int? numberingId, string level) {
            if (numberingId == null) {
                return string.Empty;
            }

            return WordPackageView.SdkNumberingDefinitionSignature(
                GetMainDocumentPart(part)?.FirstPart(WordPackageView.NumberingRelationshipType), numberingId.Value, level);
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

        private static string GetHyperlinkSignature(WordPartView? part, OfficeXmlNode hyperlink) {
            var builder = new StringBuilder();
            if (hyperlink.Attribute(OfficeIMO.Core.Internal.OpenXml.OfficeXmlNamespaces.Relationships, "id") is string relationshipId) {
                builder.Append("id:");
                builder.Append(GetRelationshipTarget(part, relationshipId));
            }

            if (hyperlink.Attribute(W, "anchor") is string anchor) {
                builder.Append("#");
                builder.Append(anchor);
            }

            if (hyperlink.Attribute(W, "docLocation") is string docLocation) {
                builder.Append("@");
                builder.Append(docLocation);
            }

            if (ParseOnOff(hyperlink.Attribute(W, "history")) is bool history) {
                builder.Append("|history=");
                builder.Append(history ? "true" : "false");
            }

            return builder.ToString();
        }

        private static string GetRelationshipTarget(WordPartView? part, string relationshipId) {
            if (part == null) {
                return relationshipId;
            }

            string? hyperlinkTarget = part.HyperlinkTarget(relationshipId);
            if (hyperlinkTarget != null) {
                return hyperlinkTarget;
            }

            if (part.TryGetExternalRelationship(relationshipId, out _, out string uri)) {
                return uri;
            }

            return relationshipId;
        }

        private static string GetFootnoteReferenceSignature(WordPartView? part, long? noteId) {
            return GetFootnoteReferenceSignature(part, noteId, WordComparisonOptions.Default);
        }

        private static string GetFootnoteReferenceSignature(WordPartView? part, long? noteId, WordComparisonOptions options) {
            if (part == null || part != part.Package.Main || noteId == null) {
                return noteId?.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty;
            }

            WordPartView? notesPart = FootnotesPart(part);
            OfficeXmlNode footnote = FindVisibleNote(FootnotesRoot(notesPart), "footnote", noteId.Value);
            return footnote.IsNone ? string.Empty : GetNoteContentSignature(notesPart, footnote, options);
        }

        private static string GetEndnoteReferenceSignature(WordPartView? part, long? noteId) {
            return GetEndnoteReferenceSignature(part, noteId, WordComparisonOptions.Default);
        }

        private static string GetEndnoteReferenceSignature(WordPartView? part, long? noteId, WordComparisonOptions options) {
            if (part == null || part != part.Package.Main || noteId == null) {
                return noteId?.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty;
            }

            WordPartView? notesPart = EndnotesPart(part);
            OfficeXmlNode endnote = FindVisibleNote(EndnotesRoot(notesPart), "endnote", noteId.Value);
            return endnote.IsNone ? string.Empty : GetNoteContentSignature(notesPart, endnote, options);
        }

        private static OfficeXmlNode FindVisibleNote(OfficeXmlNode notesRoot, string noteName, long noteId) {
            if (notesRoot.IsNone) {
                return default;
            }

            foreach (OfficeXmlNode note in notesRoot.Children(W, noteName)) {
                if (GetNoteId(note) == noteId && IsVisibleNote(note)) {
                    return note;
                }
            }

            return default;
        }

        private static string GetNoteContentSignature(WordPartView? notePart, OfficeXmlNode note) {
            return GetNoteContentSignature(notePart, note, WordComparisonOptions.Default);
        }

        private static string GetNoteContentSignature(WordPartView? notePart, OfficeXmlNode note, WordComparisonOptions options) {
            return string.Join(
                TableRowSeparator,
                ChildrenW(note, "p")
                    .Select(paragraph => GetParagraphMatchText(paragraph, notePart, options))
                    .ToArray());
        }

        private static bool HasImageContent(OfficeXmlNode paragraph) {
            foreach (OfficeXmlNode element in TypedDescendants(paragraph)) {
                if (element.Is(W, "drawing") || element.Is(Vml, "imagedata")) {
                    return true;
                }
            }

            return false;
        }

        private static string ParagraphLocation(int paragraphIndex) {
            return "paragraph[" + paragraphIndex.ToString(System.Globalization.CultureInfo.InvariantCulture) + "]";
        }

        private sealed class ParagraphSnapshot : IComparisonFingerprint {
            internal ParagraphSnapshot(string partKind, WordPartView? part, OfficeXmlNode paragraph, string text, string comparisonText, string matchText, string styleId, string formatSignature, int documentOrder) {
                PartKind = partKind;
                Part = part;
                Paragraph = paragraph;
                Text = text;
                ComparisonText = comparisonText;
                MatchText = matchText;
                StyleId = styleId;
                FormatSignature = formatSignature;
                DocumentOrder = documentOrder;
                ComparisonFingerprint = CombineComparisonFingerprints(
                    GetOrdinalTextFingerprint(partKind),
                    GetOrdinalTextFingerprint(matchText));
            }

            internal string PartKind { get; }

            internal WordPartView? Part { get; }

            internal OfficeXmlNode Paragraph { get; }

            internal string Text { get; }

            internal string ComparisonText { get; }

            internal string MatchText { get; }

            internal string StyleId { get; }

            internal string FormatSignature { get; }

            internal int DocumentOrder { get; }

            public ulong ComparisonFingerprint { get; }
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
