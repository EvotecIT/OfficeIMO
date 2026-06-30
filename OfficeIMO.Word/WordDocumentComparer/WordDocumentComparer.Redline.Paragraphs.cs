using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public static partial class WordDocumentComparer {
        private static HashSet<int> ApplyParagraphFindings(
            WordprocessingDocument sourceDocument,
            WordprocessingDocument targetDocument,
            WordComparisonResult result,
            WordComparisonRedlineOptions options,
            out IReadOnlyList<RedlineParagraphEntry> sourceParagraphs,
            out IReadOnlyList<RedlineParagraphEntry> targetParagraphs) {
            sourceParagraphs = GetRedlineParagraphEntries(sourceDocument);
            targetParagraphs = GetRedlineParagraphEntries(targetDocument);
            var rewrittenParagraphs = new HashSet<int>();
            var deletedParagraphOffsets = new Dictionary<string, int>(StringComparer.Ordinal);

            foreach (WordComparisonFinding finding in result.Findings) {
                if (!ShouldTrackFinding(finding, options) ||
                    finding.Scope != WordComparisonScope.Paragraph ||
                    !IsParagraphTextRedlineFinding(finding) ||
                    !HasTrackedText(finding)) {
                    continue;
                }

                switch (finding.ChangeKind) {
                    case WordComparisonChangeKind.Modified:
                        if (finding.TargetIndex is int modifiedIndex &&
                            modifiedIndex >= 0 &&
                            modifiedIndex < targetParagraphs.Count) {
                            if (IsContentControlParagraph(targetParagraphs[modifiedIndex].Paragraph)) {
                                break;
                            }

                            if (string.Equals(finding.SourceText, finding.TargetText, StringComparison.Ordinal)) {
                                break;
                            }

                            RewriteParagraphWithTrackedText(targetParagraphs[modifiedIndex].Paragraph, finding.SourceText, finding.TargetText, options);
                            rewrittenParagraphs.Add(modifiedIndex);
                        }

                        break;
                    case WordComparisonChangeKind.Inserted:
                        if (finding.TargetIndex is int insertedIndex &&
                            insertedIndex >= 0 &&
                            insertedIndex < targetParagraphs.Count &&
                            !string.IsNullOrEmpty(finding.TargetText) &&
                            !IsContentControlParagraph(targetParagraphs[insertedIndex].Paragraph)) {
                            RewriteParagraphWithTrackedText(targetParagraphs[insertedIndex].Paragraph, null, finding.TargetText, options);
                            rewrittenParagraphs.Add(insertedIndex);
                        }

                        break;
                    case WordComparisonChangeKind.Deleted:
                        if (finding.SourceIndex is int sourceIndex &&
                            sourceIndex >= 0 &&
                            sourceIndex < sourceParagraphs.Count &&
                            !string.IsNullOrEmpty(finding.SourceText) &&
                            !IsContentControlParagraph(sourceParagraphs[sourceIndex].Paragraph)) {
                            RedlineParagraphEntry sourceEntry = sourceParagraphs[sourceIndex];
                            deletedParagraphOffsets.TryGetValue(sourceEntry.PartKey, out int deletedOffset);
                            int targetLocalIndex = Math.Max(0, sourceEntry.LocalIndex - deletedOffset);
                            InsertDeletedParagraph(targetDocument, targetParagraphs, sourceEntry, targetLocalIndex, finding.SourceText!, options);
                            deletedParagraphOffsets[sourceEntry.PartKey] = deletedOffset + 1;
                        }

                        break;
                }
            }

            ApplyParagraphFormattingFindings(sourceParagraphs, targetParagraphs, result, options);
            return rewrittenParagraphs;
        }

        private static void ApplyParagraphFormattingFindings(
            IReadOnlyList<RedlineParagraphEntry> sourceParagraphs,
            IReadOnlyList<RedlineParagraphEntry> targetParagraphs,
            WordComparisonResult result,
            WordComparisonRedlineOptions options) {
            foreach (WordComparisonFinding finding in result.Findings) {
                if (!ShouldTrackFinding(finding, options) ||
                    finding.Scope != WordComparisonScope.Paragraph ||
                    !IsParagraphFormattingRedlineFinding(finding) ||
                    finding.ChangeKind != WordComparisonChangeKind.Modified ||
                    finding.TargetIndex is not int targetIndex ||
                    targetIndex < 0 ||
                    targetIndex >= targetParagraphs.Count) {
                    continue;
                }

                int sourceIndex = finding.SourceIndex ?? targetIndex;
                if (sourceIndex < 0 || sourceIndex >= sourceParagraphs.Count) {
                    continue;
                }

                ApplyParagraphFormattingFinding(sourceParagraphs[sourceIndex].Paragraph, targetParagraphs[targetIndex].Paragraph, options);
            }
        }

        private static bool IsParagraphTextRedlineFinding(WordComparisonFinding finding) {
            return string.Equals(finding.Message, "Paragraph text changed.", StringComparison.Ordinal) ||
                   string.Equals(finding.Message, "Paragraph inserted.", StringComparison.Ordinal) ||
                   string.Equals(finding.Message, "Paragraph deleted.", StringComparison.Ordinal);
        }

        private static bool IsParagraphFormattingRedlineFinding(WordComparisonFinding finding) {
            return string.Equals(finding.Message, "Paragraph style id changed.", StringComparison.Ordinal) ||
                   string.Equals(finding.Message, "Paragraph effective formatting changed.", StringComparison.Ordinal);
        }

        private static void ApplyParagraphFormattingFinding(Paragraph sourceParagraph, Paragraph targetParagraph, WordComparisonRedlineOptions options) {
            targetParagraph.ParagraphProperties ??= new ParagraphProperties();
            foreach (ParagraphPropertiesChange existingChange in targetParagraph.ParagraphProperties.Elements<ParagraphPropertiesChange>().ToList()) {
                existingChange.Remove();
            }

            targetParagraph.ParagraphProperties.ParagraphPropertiesChange = CreateParagraphPropertiesChange(sourceParagraph.ParagraphProperties, options);
        }

        private static void ApplyRunFindings(
            IReadOnlyList<RedlineParagraphEntry> sourceParagraphs,
            IReadOnlyList<RedlineParagraphEntry> targetParagraphs,
            WordComparisonResult result,
            WordComparisonRedlineOptions options,
            HashSet<int> rewrittenParagraphs) {
            foreach (WordComparisonFinding finding in result.Findings) {
                if (!ShouldTrackFinding(finding, options) ||
                    finding.Scope != WordComparisonScope.Run ||
                    !TryParseParagraphRunLocation(finding.Location, out int paragraphIndex, out int runIndex) ||
                    paragraphIndex < 0 ||
                    paragraphIndex >= targetParagraphs.Count ||
                    rewrittenParagraphs.Contains(paragraphIndex) ||
                    IsContentControlParagraph(targetParagraphs[paragraphIndex].Paragraph)) {
                    continue;
                }

                if (IsFormattingFinding(finding)) {
                    ApplyRunFormattingFinding(sourceParagraphs, targetParagraphs[paragraphIndex].Paragraph, paragraphIndex, runIndex, finding, options);
                } else {
                    ApplyRunFinding(targetParagraphs[paragraphIndex].Paragraph, runIndex, finding, options);
                }
            }
        }

        private static void ApplyRunFormattingFinding(
            IReadOnlyList<RedlineParagraphEntry> sourceParagraphs,
            Paragraph targetParagraph,
            int paragraphIndex,
            int runIndex,
            WordComparisonFinding finding,
            WordComparisonRedlineOptions options) {
            int sourceParagraphIndex = FindSourceParagraphIndex(sourceParagraphs, targetParagraph, paragraphIndex);
            if (!string.Equals(finding.Message, "Run formatting changed.", StringComparison.Ordinal) ||
                finding.ChangeKind != WordComparisonChangeKind.Modified ||
                sourceParagraphIndex < 0 ||
                sourceParagraphIndex >= sourceParagraphs.Count) {
                return;
            }

            int sourceRunIndex = finding.SourceIndex ?? runIndex;
            List<Run> sourceRuns = GetDirectParagraphRuns(sourceParagraphs[sourceParagraphIndex].Paragraph);
            List<Run> targetRuns = GetDirectParagraphRuns(targetParagraph);
            if (sourceRunIndex < 0 ||
                sourceRunIndex >= sourceRuns.Count ||
                runIndex < 0 ||
                runIndex >= targetRuns.Count) {
                return;
            }

            Run targetRun = targetRuns[runIndex];
            targetRun.RunProperties ??= new RunProperties();
            foreach (RunPropertiesChange existingChange in targetRun.RunProperties.Elements<RunPropertiesChange>().ToList()) {
                existingChange.Remove();
            }

            targetRun.RunProperties.RunPropertiesChange = CreateRunPropertiesChange(sourceRuns[sourceRunIndex].RunProperties, options);
        }

        private static int FindSourceParagraphIndex(IReadOnlyList<RedlineParagraphEntry> sourceParagraphs, Paragraph targetParagraph, int fallbackIndex) {
            string targetText = GetParagraphText(targetParagraph);
            if (!string.IsNullOrEmpty(targetText)) {
                if (fallbackIndex >= 0 &&
                    fallbackIndex < sourceParagraphs.Count &&
                    string.Equals(GetParagraphText(sourceParagraphs[fallbackIndex].Paragraph), targetText, StringComparison.Ordinal)) {
                    return fallbackIndex;
                }

                for (int index = 0; index < sourceParagraphs.Count; index++) {
                    if (string.Equals(GetParagraphText(sourceParagraphs[index].Paragraph), targetText, StringComparison.Ordinal)) {
                        return index;
                    }
                }
            }

            return fallbackIndex;
        }

        private static List<Run> GetDirectParagraphRuns(Paragraph paragraph) {
            return paragraph.Descendants<Run>()
                .Where(run => run.Ancestors<Paragraph>().FirstOrDefault() == paragraph)
                .ToList();
        }

        private static RunPropertiesChange CreateRunPropertiesChange(RunProperties? sourceProperties, WordComparisonRedlineOptions options) {
            var previousProperties = new PreviousRunProperties();
            if (sourceProperties != null) {
                RunProperties clonedProperties = (RunProperties)sourceProperties.CloneNode(true);
                foreach (RunPropertiesChange existingChange in clonedProperties.Elements<RunPropertiesChange>().ToList()) {
                    existingChange.Remove();
                }

                foreach (OpenXmlElement child in clonedProperties.ChildElements.ToList()) {
                    previousProperties.Append(child.CloneNode(true));
                }
            }

            return new RunPropertiesChange(previousProperties) {
                Author = options.Author,
                Date = options.DateTime ?? DateTime.Now,
                Id = WordHeadersAndFooters.GenerateRevisionId()
            };
        }

        private static ParagraphPropertiesChange CreateParagraphPropertiesChange(ParagraphProperties? sourceProperties, WordComparisonRedlineOptions options) {
            var previousProperties = new PreviousParagraphProperties();
            if (sourceProperties != null) {
                ParagraphProperties clonedProperties = (ParagraphProperties)sourceProperties.CloneNode(true);
                foreach (ParagraphPropertiesChange existingChange in clonedProperties.Elements<ParagraphPropertiesChange>().ToList()) {
                    existingChange.Remove();
                }

                foreach (OpenXmlElement child in clonedProperties.ChildElements.ToList()) {
                    previousProperties.Append(child.CloneNode(true));
                }
            }

            return new ParagraphPropertiesChange(previousProperties) {
                Author = options.Author,
                Date = options.DateTime ?? DateTime.Now,
                Id = WordHeadersAndFooters.GenerateRevisionId()
            };
        }

        private static List<RedlineParagraphEntry> GetRedlineParagraphEntries(WordprocessingDocument document) {
            var entries = new List<RedlineParagraphEntry>();
            MainDocumentPart? mainPart = document.MainDocumentPart;
            AddRedlineParagraphEntries(entries, BodyPartKey, mainPart?.Document?.Body, BodyPartOrderBase);

            if (mainPart == null) {
                return entries;
            }

            int headerIndex = 0;
            foreach (KeyValuePair<HeaderPart, string> headerPartKey in CreateOrderedHeaderPartKeys(mainPart)) {
                AddRedlineParagraphEntries(entries, headerPartKey.Value, headerPartKey.Key.Header, HeaderPartOrderBase + (headerIndex * RelatedPartOrderStride));
                headerIndex++;
            }

            int footerIndex = 0;
            foreach (KeyValuePair<FooterPart, string> footerPartKey in CreateOrderedFooterPartKeys(mainPart)) {
                AddRedlineParagraphEntries(entries, footerPartKey.Value, footerPartKey.Key.Footer, FooterPartOrderBase + (footerIndex * RelatedPartOrderStride));
                footerIndex++;
            }

            List<Footnote> footnotes = GetReferencedFootnotes(mainPart);
            for (int footnoteIndex = 0; footnoteIndex < footnotes.Count; footnoteIndex++) {
                string noteId = footnoteIndex.ToString(System.Globalization.CultureInfo.InvariantCulture);
                AddRedlineParagraphEntries(entries, FootnotePartKeyPrefix + noteId, footnotes[footnoteIndex], FootnotePartOrderBase + (footnoteIndex * RelatedPartOrderStride));
            }

            List<Endnote> endnotes = GetReferencedEndnotes(mainPart);
            for (int endnoteIndex = 0; endnoteIndex < endnotes.Count; endnoteIndex++) {
                string noteId = endnoteIndex.ToString(System.Globalization.CultureInfo.InvariantCulture);
                AddRedlineParagraphEntries(entries, EndnotePartKeyPrefix + noteId, endnotes[endnoteIndex], EndnotePartOrderBase + (endnoteIndex * RelatedPartOrderStride));
            }

            return entries;
        }

        private static void AddRedlineParagraphEntries(List<RedlineParagraphEntry> entries, string partKey, OpenXmlCompositeElement? container, int orderBase) {
            if (container == null) {
                return;
            }

            int localIndex = 0;
            foreach (OrderedElement ordered in EnumerateDescendantsWithOrder(container, orderBase)) {
                if (ordered.Element is not Paragraph paragraph ||
                    paragraph.Ancestors<TableCell>().Any()) {
                    continue;
                }

                if (GetParagraphText(paragraph).Length == 0 && HasImageContent(paragraph)) {
                    continue;
                }

                entries.Add(new RedlineParagraphEntry(partKey, container, paragraph, localIndex));
                localIndex++;
            }
        }

        private static bool IsContentControlParagraph(Paragraph paragraph) {
            return paragraph.Ancestors<SdtElement>().Any() || paragraph.Descendants<SdtElement>().Any();
        }

        private static void InsertDeletedParagraph(WordprocessingDocument targetDocument, IReadOnlyList<RedlineParagraphEntry> targetParagraphs, RedlineParagraphEntry sourceEntry, int targetLocalIndex, string deletedText, WordComparisonRedlineOptions options) {
            RedlineParagraphEntry? firstTargetPartEntry = targetParagraphs.FirstOrDefault(entry => string.Equals(entry.PartKey, sourceEntry.PartKey, StringComparison.Ordinal));
            var paragraph = new Paragraph();
            paragraph.Append(CreateDeletedRun(deletedText, options));

            List<RedlineParagraphEntry> targetPartEntries = targetParagraphs
                .Where(entry => string.Equals(entry.PartKey, sourceEntry.PartKey, StringComparison.Ordinal))
                .ToList();

            if (targetLocalIndex >= 0 && targetLocalIndex < targetPartEntries.Count) {
                targetPartEntries[targetLocalIndex].Paragraph.InsertBeforeSelf(paragraph);
                return;
            }

            OpenXmlCompositeElement? targetContainer = firstTargetPartEntry?.Container ?? GetRedlineContainerByPartKey(targetDocument, sourceEntry.PartKey);
            if (targetContainer != null) {
                AppendRedlineParagraph(targetContainer, paragraph);
            }
        }

        private static void AppendRedlineParagraph(OpenXmlCompositeElement container, Paragraph paragraph) {
            if (container is Body body) {
                SectionProperties? sectionProperties = body.Elements<SectionProperties>().LastOrDefault();
                if (sectionProperties != null) {
                    body.InsertBefore(paragraph, sectionProperties);
                    return;
                }
            }

            container.Append(paragraph);
        }

        private static bool TryParseParagraphRunLocation(string location, out int paragraphIndex, out int runIndex) {
            paragraphIndex = -1;
            runIndex = -1;
            const string paragraphPrefix = "paragraph[";
            const string runToken = "]/run[";
            if (!location.StartsWith(paragraphPrefix, StringComparison.Ordinal)) {
                return false;
            }

            int runTokenIndex = location.IndexOf(runToken, StringComparison.Ordinal);
            if (runTokenIndex < 0 || !location.EndsWith("]", StringComparison.Ordinal)) {
                return false;
            }

            string paragraphText = location.Substring(paragraphPrefix.Length, runTokenIndex - paragraphPrefix.Length);
            string runText = location.Substring(runTokenIndex + runToken.Length, location.Length - runTokenIndex - runToken.Length - 1);
            return int.TryParse(paragraphText, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out paragraphIndex) &&
                   int.TryParse(runText, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out runIndex);
        }

        private sealed class RedlineParagraphEntry {
            internal RedlineParagraphEntry(string partKey, OpenXmlCompositeElement container, Paragraph paragraph, int localIndex) {
                PartKey = partKey;
                Container = container;
                Paragraph = paragraph;
                LocalIndex = localIndex;
            }

            internal string PartKey { get; }

            internal OpenXmlCompositeElement Container { get; }

            internal Paragraph Paragraph { get; }

            internal int LocalIndex { get; }
        }
    }
}
