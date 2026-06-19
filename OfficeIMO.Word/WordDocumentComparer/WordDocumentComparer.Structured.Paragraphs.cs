using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text;
using V = DocumentFormat.OpenXml.Vml;

namespace OfficeIMO.Word {
    public static partial class WordDocumentComparer {
        private static void AnalyzeParagraphs(WordDocument source, WordDocument target, WordComparisonResult result) {
            List<ParagraphSnapshot> sourceParagraphs = GetLogicalBodyParagraphs(source);
            List<ParagraphSnapshot> targetParagraphs = GetLogicalBodyParagraphs(target);
            IReadOnlyList<MatchedIndexPair> matchedParagraphs = FindMatchingIndexes(
                sourceParagraphs,
                targetParagraphs,
                ParagraphSnapshotEqualityComparer.Instance);

            int sourceStart = 0;
            int targetStart = 0;

            foreach (MatchedIndexPair match in matchedParagraphs) {
                AddParagraphRangeFindings(sourceParagraphs, targetParagraphs, sourceStart, match.SourceIndex, targetStart, match.TargetIndex, result);
                sourceStart = match.SourceIndex + 1;
                targetStart = match.TargetIndex + 1;
            }

            AddParagraphRangeFindings(sourceParagraphs, targetParagraphs, sourceStart, sourceParagraphs.Count, targetStart, targetParagraphs.Count, result);
        }

        private static void AddParagraphRangeFindings(
            IReadOnlyList<ParagraphSnapshot> sourceParagraphs,
            IReadOnlyList<ParagraphSnapshot> targetParagraphs,
            int sourceStart,
            int sourceEnd,
            int targetStart,
            int targetEnd,
            WordComparisonResult result) {
            int sourceIndex = sourceStart;
            int targetIndex = targetStart;

            while (sourceIndex < sourceEnd && targetIndex < targetEnd) {
                int betterTargetIndex = FindBetterTargetAlignmentIndex(sourceParagraphs[sourceIndex], targetParagraphs, targetIndex, targetEnd);
                if (targetEnd - targetIndex > sourceEnd - sourceIndex &&
                    betterTargetIndex > targetIndex) {
                    while (targetIndex < betterTargetIndex) {
                        AddInsertedParagraphFinding(targetParagraphs, targetIndex, result);
                        targetIndex++;
                    }

                    continue;
                }

                int betterSourceIndex = FindBetterSourceAlignmentIndex(sourceParagraphs, sourceIndex, sourceEnd, targetParagraphs[targetIndex]);
                if (sourceEnd - sourceIndex > targetEnd - targetIndex &&
                    betterSourceIndex > sourceIndex) {
                    while (sourceIndex < betterSourceIndex) {
                        AddDeletedParagraphFinding(sourceParagraphs, sourceIndex, result);
                        sourceIndex++;
                    }

                    continue;
                }

                string sourceText = sourceParagraphs[sourceIndex].Text;
                string targetText = targetParagraphs[targetIndex].Text;

                if (!string.Equals(sourceParagraphs[sourceIndex].PartKind, targetParagraphs[targetIndex].PartKind, StringComparison.Ordinal)) {
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

        private static int FindBetterTargetAlignmentIndex(ParagraphSnapshot sourceParagraph, IReadOnlyList<ParagraphSnapshot> targetParagraphs, int targetStart, int targetEnd) {
            int bestIndex = targetStart;
            double bestSimilarity = GetParagraphSimilarity(sourceParagraph, targetParagraphs[targetStart]);

            for (int index = targetStart + 1; index < targetEnd; index++) {
                double similarity = GetParagraphSimilarity(sourceParagraph, targetParagraphs[index]);
                if (similarity <= bestSimilarity) {
                    continue;
                }

                bestSimilarity = similarity;
                bestIndex = index;
                if (similarity >= 1) {
                    break;
                }
            }

            return bestIndex;
        }

        private static int FindBetterSourceAlignmentIndex(IReadOnlyList<ParagraphSnapshot> sourceParagraphs, int sourceStart, int sourceEnd, ParagraphSnapshot targetParagraph) {
            int bestIndex = sourceStart;
            double bestSimilarity = GetParagraphSimilarity(sourceParagraphs[sourceStart], targetParagraph);

            for (int index = sourceStart + 1; index < sourceEnd; index++) {
                double similarity = GetParagraphSimilarity(sourceParagraphs[index], targetParagraph);
                if (similarity <= bestSimilarity) {
                    continue;
                }

                bestSimilarity = similarity;
                bestIndex = index;
                if (similarity >= 1) {
                    break;
                }
            }

            return bestIndex;
        }

        private static double GetParagraphSimilarity(ParagraphSnapshot sourceParagraph, ParagraphSnapshot targetParagraph) {
            if (!string.Equals(sourceParagraph.PartKind, targetParagraph.PartKind, StringComparison.Ordinal)) {
                return 0;
            }

            return GetTextSimilarity(sourceParagraph.MatchText, targetParagraph.MatchText);
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

        private static List<ParagraphSnapshot> GetLogicalBodyParagraphs(WordDocument document) {
            var snapshots = new List<ParagraphSnapshot>();
            MainDocumentPart? mainPart = document._wordprocessingDocument.MainDocumentPart;
            AddParagraphSnapshots(snapshots, mainPart, mainPart?.Document?.Body, BodyPartKey, BodyPartOrderBase);

            if (mainPart != null) {
                Dictionary<HeaderPart, string> headerPartKeys = CreateHeaderPartKeys(mainPart);
                int headerIndex = 0;
                foreach (HeaderPart headerPart in mainPart.HeaderParts) {
                    if (!headerPartKeys.TryGetValue(headerPart, out string? headerPartKey)) {
                        continue;
                    }

                    AddParagraphSnapshots(snapshots, headerPart, headerPart.Header, headerPartKey, HeaderPartOrderBase + (headerIndex * RelatedPartOrderStride));
                    headerIndex++;
                }

                Dictionary<FooterPart, string> footerPartKeys = CreateFooterPartKeys(mainPart);
                int footerIndex = 0;
                foreach (FooterPart footerPart in mainPart.FooterParts) {
                    if (!footerPartKeys.TryGetValue(footerPart, out string? footerPartKey)) {
                        continue;
                    }

                    AddParagraphSnapshots(snapshots, footerPart, footerPart.Footer, footerPartKey, FooterPartOrderBase + (footerIndex * RelatedPartOrderStride));
                    footerIndex++;
                }

                int footnoteIndex = 0;
                foreach (Footnote footnote in mainPart.FootnotesPart?.Footnotes?.Elements<Footnote>() ?? Enumerable.Empty<Footnote>()) {
                    if (!IsVisibleNote(footnote)) {
                        continue;
                    }

                    string noteId = footnoteIndex.ToString(System.Globalization.CultureInfo.InvariantCulture);
                    AddParagraphSnapshots(snapshots, mainPart.FootnotesPart, footnote, FootnotePartKeyPrefix + noteId, FootnotePartOrderBase + (footnoteIndex * RelatedPartOrderStride));
                    footnoteIndex++;
                }

                int endnoteIndex = 0;
                foreach (Endnote endnote in mainPart.EndnotesPart?.Endnotes?.Elements<Endnote>() ?? Enumerable.Empty<Endnote>()) {
                    if (!IsVisibleNote(endnote)) {
                        continue;
                    }

                    string noteId = endnoteIndex.ToString(System.Globalization.CultureInfo.InvariantCulture);
                    AddParagraphSnapshots(snapshots, mainPart.EndnotesPart, endnote, EndnotePartKeyPrefix + noteId, EndnotePartOrderBase + (endnoteIndex * RelatedPartOrderStride));
                    endnoteIndex++;
                }
            }

            return snapshots;
        }

        private static void AddParagraphSnapshots(List<ParagraphSnapshot> snapshots, OpenXmlPart? part, OpenXmlElement? container, string partKind, int orderBase) {
            foreach (OrderedElement ordered in EnumerateDescendantsWithOrder(container, orderBase)) {
                if (ordered.Element is not Paragraph paragraph) {
                    continue;
                }

                if (paragraph.Ancestors<TableCell>().Any()) {
                    continue;
                }

                ParagraphTextSnapshot text = GetParagraphTextSnapshot(paragraph, part);
                if (text.Text.Length == 0 && HasImageContent(paragraph)) {
                    continue;
                }

                snapshots.Add(new ParagraphSnapshot(partKind, text.Text, text.MatchText, ordered.DocumentOrder));
            }
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

        private static ParagraphTextSnapshot GetParagraphTextSnapshot(Paragraph paragraph, OpenXmlPart? part) {
            var textBuilder = new StringBuilder();
            var matchBuilder = new StringBuilder();
            var pendingTextBuilder = new StringBuilder();
            foreach (OpenXmlElement element in paragraph.Descendants()) {
                if (element.Ancestors<Paragraph>().FirstOrDefault() != paragraph) {
                    continue;
                }

                switch (element) {
                    case Text text:
                        textBuilder.Append(text.Text);
                        pendingTextBuilder.Append(text.Text);
                        break;
                    case Hyperlink hyperlink:
                        FlushPendingTextToken(matchBuilder, pendingTextBuilder);
                        AppendMatchToken(matchBuilder, "hyperlink", GetHyperlinkSignature(part, hyperlink));
                        break;
                    case SimpleField simpleField:
                        FlushPendingTextToken(matchBuilder, pendingTextBuilder);
                        AppendMatchToken(matchBuilder, "simpleField", simpleField.Instruction?.Value ?? string.Empty);
                        break;
                    case FieldCode fieldCode:
                        FlushPendingTextToken(matchBuilder, pendingTextBuilder);
                        AppendMatchToken(matchBuilder, "fieldCode", fieldCode.Text ?? string.Empty);
                        break;
                    case TabChar:
                        FlushPendingTextToken(matchBuilder, pendingTextBuilder);
                        textBuilder.Append('\t');
                        AppendMatchToken(matchBuilder, "tab", string.Empty);
                        break;
                    case Break breakNode:
                        FlushPendingTextToken(matchBuilder, pendingTextBuilder);
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
                        FlushPendingTextToken(matchBuilder, pendingTextBuilder);
                        string footnoteReferenceId = footnoteReference.Id?.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty;
                        textBuilder.Append("[FootnoteReference:");
                        textBuilder.Append(footnoteReferenceId);
                        textBuilder.Append(']');
                        AppendMatchToken(matchBuilder, "footnoteReference", GetFootnoteReferenceSignature(part, footnoteReference.Id?.Value));
                        break;
                    case EndnoteReference endnoteReference:
                        FlushPendingTextToken(matchBuilder, pendingTextBuilder);
                        string endnoteReferenceId = endnoteReference.Id?.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty;
                        textBuilder.Append("[EndnoteReference:");
                        textBuilder.Append(endnoteReferenceId);
                        textBuilder.Append(']');
                        AppendMatchToken(matchBuilder, "endnoteReference", GetEndnoteReferenceSignature(part, endnoteReference.Id?.Value));
                        break;
                    case SymbolChar symbol:
                        FlushPendingTextToken(matchBuilder, pendingTextBuilder);
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
                        FlushPendingTextToken(matchBuilder, pendingTextBuilder);
                        textBuilder.Append('-');
                        AppendMatchToken(matchBuilder, "noBreakHyphen", string.Empty);
                        break;
                    case SoftHyphen:
                        FlushPendingTextToken(matchBuilder, pendingTextBuilder);
                        textBuilder.Append("[SoftHyphen]");
                        AppendMatchToken(matchBuilder, "softHyphen", string.Empty);
                        break;
                    case CarriageReturn:
                        FlushPendingTextToken(matchBuilder, pendingTextBuilder);
                        textBuilder.Append('\n');
                        AppendMatchToken(matchBuilder, "carriageReturn", string.Empty);
                        break;
                }
            }

            FlushPendingTextToken(matchBuilder, pendingTextBuilder);
            return new ParagraphTextSnapshot(textBuilder.ToString(), matchBuilder.ToString());
        }

        private static void FlushPendingTextToken(StringBuilder builder, StringBuilder pendingTextBuilder) {
            if (pendingTextBuilder.Length == 0) {
                return;
            }

            AppendMatchToken(builder, "text", pendingTextBuilder.ToString());
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
            if (part is not MainDocumentPart mainPart || noteId == null) {
                return noteId?.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty;
            }

            Footnote? footnote = mainPart.FootnotesPart?.Footnotes?.Elements<Footnote>()
                .FirstOrDefault(item => item.Id?.Value == noteId.Value && IsVisibleNote(item));
            return footnote == null ? string.Empty : GetNoteContentSignature(mainPart.FootnotesPart, footnote);
        }

        private static string GetEndnoteReferenceSignature(OpenXmlPart? part, long? noteId) {
            if (part is not MainDocumentPart mainPart || noteId == null) {
                return noteId?.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty;
            }

            Endnote? endnote = mainPart.EndnotesPart?.Endnotes?.Elements<Endnote>()
                .FirstOrDefault(item => item.Id?.Value == noteId.Value && IsVisibleNote(item));
            return endnote == null ? string.Empty : GetNoteContentSignature(mainPart.EndnotesPart, endnote);
        }

        private static string GetNoteContentSignature(OpenXmlPart? notePart, OpenXmlElement note) {
            return string.Join(
                TableRowSeparator,
                note.Elements<Paragraph>()
                    .Select(paragraph => GetParagraphMatchText(paragraph, notePart))
                    .ToArray());
        }

        private static bool HasImageContent(Paragraph paragraph) {
            return paragraph.Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>().Any() ||
                   paragraph.Descendants<V.ImageData>().Any();
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
            internal ParagraphSnapshot(string partKind, string text, string matchText, int documentOrder) {
                PartKind = partKind;
                Text = text;
                MatchText = matchText;
                DocumentOrder = documentOrder;
            }

            internal string PartKind { get; }

            internal string Text { get; }

            internal string MatchText { get; }

            internal int DocumentOrder { get; }
        }

        private readonly struct ParagraphTextSnapshot {
            internal ParagraphTextSnapshot(string text, string matchText) {
                Text = text;
                MatchText = matchText;
            }

            internal string Text { get; }

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
