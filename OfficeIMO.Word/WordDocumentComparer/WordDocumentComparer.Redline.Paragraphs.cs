using OfficeIMO.Word.OpenXml;

namespace OfficeIMO.Word {
    public static partial class WordDocumentComparer {
        private static HashSet<int> ApplyParagraphFindings(
            WordPackageView sourceDocument,
            WordPackageView targetDocument,
            WordComparisonResult result,
            WordComparisonRedlineOptions options,
            out IReadOnlyList<RedlineParagraphEntry> sourceParagraphs,
            out IReadOnlyList<RedlineParagraphEntry> targetParagraphs) {
            sourceParagraphs = GetRedlineParagraphEntries(sourceDocument);
            targetParagraphs = GetRedlineParagraphEntries(targetDocument);
            var rewrittenParagraphs = new HashSet<int>();

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
                            List<RedlineParagraphEntry> sourcePartEntries = sourceParagraphs
                                .Where(entry => string.Equals(entry.PartKey, sourceEntry.PartKey, StringComparison.Ordinal))
                                .ToList();
                            List<RedlineParagraphEntry> targetPartEntries = targetParagraphs
                                .Where(entry => string.Equals(entry.PartKey, sourceEntry.PartKey, StringComparison.Ordinal))
                                .ToList();
                            int targetLocalIndex = FindTargetGapByNeighborIdentity(
                                sourcePartEntries,
                                targetPartEntries,
                                sourceEntry.LocalIndex,
                                entry => GetParagraphText(entry.Paragraph),
                                entry => GetParagraphText(entry.Paragraph));
                            InsertDeletedParagraph(targetDocument, targetParagraphs, sourceEntry, targetLocalIndex, finding.SourceText!, options);
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

        private static void ApplyParagraphFormattingFinding(OfficeXmlNode sourceParagraph, OfficeXmlNode targetParagraph, WordComparisonRedlineOptions options) {
            OfficeXmlNode properties = targetParagraph.Child(W, "pPr");
            if (properties.IsNone) {
                properties = targetParagraph.InsertOrdered(targetParagraph.CreateElement(W, "pPr"), WPView.Order);
            }

            properties.RemoveChildren(W, "pPrChange");
            SetPropertyChild(properties, CreatePropertiesChange(properties, "pPrChange", sourceParagraph.Child(W, "pPr"), options), order: null);
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
            OfficeXmlNode targetParagraph,
            int paragraphIndex,
            int runIndex,
            WordComparisonFinding finding,
            WordComparisonRedlineOptions options) {
            int sourceParagraphIndex = TryGetDetailedSourceParagraphIndex(finding, out int detailedSourceParagraphIndex)
                ? detailedSourceParagraphIndex
                : FindSourceParagraphIndex(sourceParagraphs, targetParagraph, paragraphIndex);
            if (!string.Equals(finding.Message, "Run formatting changed.", StringComparison.Ordinal) ||
                finding.ChangeKind != WordComparisonChangeKind.Modified ||
                sourceParagraphIndex < 0 ||
                sourceParagraphIndex >= sourceParagraphs.Count) {
                return;
            }

            int sourceRunIndex = finding.SourceIndex ?? runIndex;
            List<OfficeXmlNode> sourceRuns = GetDirectParagraphRuns(sourceParagraphs[sourceParagraphIndex].Paragraph);
            List<OfficeXmlNode> targetRuns = GetDirectParagraphRuns(targetParagraph);
            if (sourceRunIndex < 0 ||
                sourceRunIndex >= sourceRuns.Count ||
                runIndex < 0 ||
                runIndex >= targetRuns.Count) {
                return;
            }

            OfficeXmlNode targetRun = targetRuns[runIndex];
            OfficeXmlNode properties = targetRun.Child(W, "rPr");
            if (properties.IsNone) {
                properties = targetRun.InsertOrdered(targetRun.CreateElement(W, "rPr"), WRView.Order);
            }

            properties.RemoveChildren(W, "rPrChange");
            SetPropertyChild(properties, CreatePropertiesChange(properties, "rPrChange", sourceRuns[sourceRunIndex].Child(W, "rPr"), options), order: null);
        }

        private static bool TryGetDetailedSourceParagraphIndex(WordComparisonFinding finding, out int sourceParagraphIndex) {
            sourceParagraphIndex = -1;
            const string token = "/sourceParagraph[";
            int tokenIndex = finding.DetailedLocation.IndexOf(token, StringComparison.Ordinal);
            if (tokenIndex < 0 || !finding.DetailedLocation.EndsWith("]", StringComparison.Ordinal)) {
                return false;
            }

            int valueStart = tokenIndex + token.Length;
            string value = finding.DetailedLocation.Substring(valueStart, finding.DetailedLocation.Length - valueStart - 1);
            return int.TryParse(value, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out sourceParagraphIndex);
        }

        private static int FindSourceParagraphIndex(IReadOnlyList<RedlineParagraphEntry> sourceParagraphs, OfficeXmlNode targetParagraph, int fallbackIndex) {
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

        private static List<OfficeXmlNode> GetDirectParagraphRuns(OfficeXmlNode paragraph) {
            return DescendantsW(paragraph, "r")
                .Where(run => AncestorW(run, "p") == paragraph)
                .ToList();
        }

        /// <summary>
        /// A <c>w:rPrChange</c> or <c>w:pPrChange</c> recording <paramref name="sourceProperties"/> as the previous
        /// properties, without their own revision record, made for <paramref name="properties"/> to hold.
        /// </summary>
        private static OfficeXmlNode CreatePropertiesChange(OfficeXmlNode properties, string changeName, OfficeXmlNode sourceProperties, WordComparisonRedlineOptions options) {
            OfficeXmlNode change = properties.CreateElement(W, changeName);
            string previousName = changeName == "rPrChange" ? "rPr" : "pPr";
            OfficeXmlNode previousProperties = change.CreateElement(W, previousName);
            if (!sourceProperties.IsNone) {
                OfficeXmlNode clonedProperties = sourceProperties.Clone();
                clonedProperties.RemoveChildren(W, changeName);
                foreach (OfficeXmlNode child in clonedProperties.Children().ToList()) {
                    previousProperties.Append(child.Clone());
                }
            }

            change.Append(previousProperties);
            SetRevisionAttributes(change, options.Author, options.DateTime ?? DateTime.Now);
            return change;
        }

        /// <summary>
        /// Puts a property element in place of the existing one with its name, or in schema order, as a typed SDK property
        /// does. A revision record (<c>w:rPrChange</c>, <c>w:pPrChange</c>) comes last in its schema, after extension
        /// properties too, so without an order it is appended.
        /// </summary>
        private static OfficeXmlNode SetPropertyChild(OfficeXmlNode properties, OfficeXmlNode child, OfficeIMO.Core.Internal.OpenXml.OfficeXmlChildOrder? order) {
            OfficeXmlNode existing = properties.Child(child.NamespaceId, child.LocalName);
            if (existing.IsNone) {
                return order == null ? properties.Append(child) : properties.InsertOrdered(child, order);
            }

            OfficeXmlNode inserted = properties.InsertBefore(child, existing);
            existing.Remove();
            return inserted;
        }

        private static List<RedlineParagraphEntry> GetRedlineParagraphEntries(WordPackageView document) {
            var entries = new List<RedlineParagraphEntry>();
            WordPartView? mainPart = document.Main;
            AddRedlineParagraphEntries(entries, BodyPartKey, BodyOf(mainPart), BodyPartOrderBase);

            if (mainPart == null) {
                return entries;
            }

            int headerIndex = 0;
            foreach (KeyValuePair<WordPartView, string> headerPartKey in CreateOrderedHeaderPartKeys(mainPart)) {
                AddRedlineParagraphEntries(entries, headerPartKey.Value, HeaderRoot(headerPartKey.Key), HeaderPartOrderBase + (headerIndex * RelatedPartOrderStride));
                headerIndex++;
            }

            int footerIndex = 0;
            foreach (KeyValuePair<WordPartView, string> footerPartKey in CreateOrderedFooterPartKeys(mainPart)) {
                AddRedlineParagraphEntries(entries, footerPartKey.Value, FooterRoot(footerPartKey.Key), FooterPartOrderBase + (footerIndex * RelatedPartOrderStride));
                footerIndex++;
            }

            List<OfficeXmlNode> footnotes = GetReferencedFootnotes(mainPart);
            for (int footnoteIndex = 0; footnoteIndex < footnotes.Count; footnoteIndex++) {
                string noteId = GetNotePartKeyId(footnotes[footnoteIndex], footnoteIndex);
                AddRedlineParagraphEntries(entries, FootnotePartKeyPrefix + noteId, footnotes[footnoteIndex], FootnotePartOrderBase + (footnoteIndex * RelatedPartOrderStride));
            }

            List<OfficeXmlNode> endnotes = GetReferencedEndnotes(mainPart);
            for (int endnoteIndex = 0; endnoteIndex < endnotes.Count; endnoteIndex++) {
                string noteId = GetNotePartKeyId(endnotes[endnoteIndex], endnoteIndex);
                AddRedlineParagraphEntries(entries, EndnotePartKeyPrefix + noteId, endnotes[endnoteIndex], EndnotePartOrderBase + (endnoteIndex * RelatedPartOrderStride));
            }

            return entries;
        }

        private static void AddRedlineParagraphEntries(List<RedlineParagraphEntry> entries, string partKey, OfficeXmlNode container, int orderBase) {
            if (container.IsNone) {
                return;
            }

            int localIndex = 0;
            foreach (OrderedElement ordered in EnumerateDescendantsWithOrder(container, orderBase)) {
                OfficeXmlNode paragraph = ordered.Element;
                if (!paragraph.Is(W, "p") || HasAncestorW(paragraph, "tc")) {
                    continue;
                }

                if (GetParagraphText(paragraph).Length == 0 && HasImageContent(paragraph)) {
                    continue;
                }

                entries.Add(new RedlineParagraphEntry(partKey, container, paragraph, localIndex));
                localIndex++;
            }
        }

        private static bool IsContentControlParagraph(OfficeXmlNode paragraph) {
            return HasAncestorW(paragraph, "sdt") || HasDescendant(paragraph, W, "sdt");
        }

        private static void InsertDeletedParagraph(WordPackageView targetDocument, IReadOnlyList<RedlineParagraphEntry> targetParagraphs, RedlineParagraphEntry sourceEntry, int targetLocalIndex, string deletedText, WordComparisonRedlineOptions options) {
            RedlineParagraphEntry? firstTargetPartEntry = targetParagraphs.FirstOrDefault(entry => string.Equals(entry.PartKey, sourceEntry.PartKey, StringComparison.Ordinal));
            // The paragraph is made before its place is known, as before, so revision ids are handed out in the same order.
            OfficeXmlNode context = firstTargetPartEntry?.Container ?? BodyOf(targetDocument.Main);
            if (context.IsNone) {
                context = targetParagraphs.Count > 0 ? targetParagraphs[0].Container : default;
            }

            if (context.IsNone) {
                return;
            }

            OfficeXmlNode paragraph = CreateDeletedParagraph(context, deletedText, options);

            List<RedlineParagraphEntry> targetPartEntries = targetParagraphs
                .Where(entry => string.Equals(entry.PartKey, sourceEntry.PartKey, StringComparison.Ordinal))
                .ToList();

            if (targetLocalIndex >= 0 && targetLocalIndex < targetPartEntries.Count) {
                OfficeXmlNode anchor = targetPartEntries[targetLocalIndex].Paragraph;
                anchor.Parent.InsertBefore(paragraph, anchor);
                return;
            }

            OfficeXmlNode targetContainer = firstTargetPartEntry?.Container ?? GetRedlineContainerByPartKey(targetDocument, sourceEntry.PartKey);
            if (!targetContainer.IsNone) {
                AppendRedlineParagraph(targetContainer, paragraph);
            }
        }

        private static OfficeXmlNode CreateDeletedParagraph(OfficeXmlNode context, string deletedText, WordComparisonRedlineOptions options) {
            OfficeXmlNode paragraph = context.CreateElement(W, "p");
            paragraph.Append(CreateDeletedRun(paragraph, deletedText, options));
            return paragraph;
        }

        private static void AppendRedlineParagraph(OfficeXmlNode container, OfficeXmlNode paragraph) {
            if (container.Is(W, "body")) {
                OfficeXmlNode sectionProperties = container.Children(W, "sectPr").Last();
                if (!sectionProperties.IsNone) {
                    container.InsertBefore(paragraph, sectionProperties);
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
            internal RedlineParagraphEntry(string partKey, OfficeXmlNode container, OfficeXmlNode paragraph, int localIndex) {
                PartKey = partKey;
                Container = container;
                Paragraph = paragraph;
                LocalIndex = localIndex;
            }

            internal string PartKey { get; }

            internal OfficeXmlNode Container { get; }

            internal OfficeXmlNode Paragraph { get; }

            internal int LocalIndex { get; }
        }
    }
}
