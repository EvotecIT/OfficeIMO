using AngleSharp.Dom;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Threading;

namespace OfficeIMO.Word.Html {
    internal partial class WordToHtmlConverter {
        private static bool TryAppendNoteReference(
            IDocument htmlDoc,
            WordParagraph run,
            WordToHtmlOptions options,
            bool processNotes,
            List<INode> nodes,
            List<(int Number, WordFootNote Note)> footnotes,
            Dictionary<long, int> footnoteMap,
            List<(int Number, WordEndNote Note)> endnotes,
            Dictionary<long, int> endnoteMap,
            OpenXmlElement? artifactElement = null) {
            if (!processNotes) {
                return false;
            }

            WordFootNote? footnote = artifactElement is FootnoteReference footnoteReference
                ? new WordFootNote(run._document, run._paragraph, new Run((FootnoteReference)footnoteReference.CloneNode(true)))
                : artifactElement == null ? run.FootNote : null;
            if (options.ExportFootnotes && footnote != null) {
                if (IsBlockquoteCiteReference(run.CharacterStyleId)) {
                    return true;
                }

                var note = footnote;
                if (!TryReplaceLastNodeWithAbbreviation(run.CharacterStyleId, htmlDoc, nodes, note.Paragraphs?.Skip(1).Select(r => r.Text))) {
                    long id = note.ReferenceId ?? 0;
                    if (!footnoteMap.TryGetValue(id, out int number)) {
                        number = footnotes.Count + 1;
                        footnoteMap[id] = number;
                        footnotes.Add((number, note));
                    }
                    var sup = CreateOutputElement(htmlDoc, "sup");
                    var a = CreateOutputElement(htmlDoc, "a");
                    string numberText = number.ToString(System.Globalization.CultureInfo.InvariantCulture);
                    SetOutputAttribute(htmlDoc, a, "href", "#fn" + numberText, "FootnoteReference:href");
                    SetOutputAttribute(htmlDoc, a, "id", "fnref" + numberText, "FootnoteReference:id");
                    SetOutputText(htmlDoc, a, numberText, "FootnoteReference:text");
                    sup.AppendChild(a);
                    nodes.Add(sup);
                }

                return true;
            }

            WordEndNote? endnote = artifactElement is EndnoteReference endnoteReference
                ? new WordEndNote(run._document, run._paragraph, new Run((EndnoteReference)endnoteReference.CloneNode(true)))
                : artifactElement == null ? run.EndNote : null;
            if (options.ExportEndnotes && endnote != null) {
                if (IsBlockquoteCiteReference(run.CharacterStyleId)) {
                    return true;
                }

                var note = endnote;
                if (!TryReplaceLastNodeWithAbbreviation(run.CharacterStyleId, htmlDoc, nodes, note.Paragraphs?.Skip(1).Select(r => r.Text))) {
                    long id = note.ReferenceId ?? 0;
                    if (!endnoteMap.TryGetValue(id, out int number)) {
                        number = endnotes.Count + 1;
                        endnoteMap[id] = number;
                        endnotes.Add((number, note));
                    }
                    var sup = CreateOutputElement(htmlDoc, "sup");
                    var a = CreateOutputElement(htmlDoc, "a");
                    string numberText = number.ToString(System.Globalization.CultureInfo.InvariantCulture);
                    SetOutputAttribute(htmlDoc, a, "href", "#en" + numberText, "EndnoteReference:href");
                    SetOutputAttribute(htmlDoc, a, "id", "enref" + numberText, "EndnoteReference:id");
                    SetOutputText(htmlDoc, a, numberText, "EndnoteReference:text");
                    sup.AppendChild(a);
                    nodes.Add(sup);
                }

                return true;
            }

            return false;
        }

        private static bool IsBlockquoteCiteReference(string? characterStyleId) {
            return string.Equals(characterStyleId, HtmlSemanticStyleIds.BlockquoteCite, StringComparison.OrdinalIgnoreCase);
        }

        private static bool TryGetBlockquoteCiteAttribute(WordParagraph paragraph, out string cite) {
            if (HtmlSemanticMetadata.TryGetBlockquoteCite(paragraph, out cite)) {
                return true;
            }

            foreach (var run in paragraph.GetRuns()) {
                if (!IsBlockquoteCiteReference(run.CharacterStyleId)) {
                    continue;
                }

                if (run.FootNote != null && TryGetNoteCitation(run.FootNote.Paragraphs, out cite)) {
                    return true;
                }

                if (run.EndNote != null && TryGetNoteCitation(run.EndNote.Paragraphs, out cite)) {
                    return true;
                }
            }

            cite = string.Empty;
            return false;
        }

        private static bool TryGetNoteCitation(IEnumerable<WordParagraph>? noteParagraphs, out string cite) {
            foreach (var paragraph in noteParagraphs?.Skip(1) ?? Enumerable.Empty<WordParagraph>()) {
                if (paragraph.Hyperlink?.Uri != null) {
                    cite = paragraph.Hyperlink.Uri.ToString();
                    return true;
                }

                foreach (var run in paragraph.GetRuns()) {
                    if (run.Hyperlink?.Uri != null) {
                        cite = run.Hyperlink.Uri.ToString();
                        return true;
                    }
                }

                if (!string.IsNullOrWhiteSpace(paragraph.Text)) {
                    cite = paragraph.Text.Trim();
                    return true;
                }
            }

            cite = string.Empty;
            return false;
        }

        private static void DiscoverTransitiveNotes(
            List<(int Number, WordFootNote Note)> footnotes, Dictionary<long, int> footnoteMap,
            List<(int Number, WordEndNote Note)> endnotes, Dictionary<long, int> endnoteMap,
            WordToHtmlOptions options, CancellationToken cancellationToken) {
            int footIndex = 0, endIndex = 0;
            while (footIndex < footnotes.Count || endIndex < endnotes.Count) {
                cancellationToken.ThrowIfCancellationRequested();
                if (footIndex < footnotes.Count)
                    Scan(footnotes[footIndex++].Note.Paragraphs?.Skip(1));
                if (endIndex < endnotes.Count)
                    Scan(endnotes[endIndex++].Note.Paragraphs?.Skip(1));
            }

            void Scan(IEnumerable<WordParagraph>? paragraphs) {
                foreach (WordParagraph paragraph in (paragraphs ?? Enumerable.Empty<WordParagraph>())
                    .Distinct(ParagraphElementComparer.Instance)) {
                    foreach (WordParagraph run in paragraph.GetRuns()) {
                        if (IsBlockquoteCiteReference(run.CharacterStyleId) ||
                            string.Equals(run.CharacterStyleId, "HtmlAbbr", StringComparison.OrdinalIgnoreCase))
                            continue;
                        IEnumerable<OpenXmlElement> children = run._visibleRunSourceChildren ??
                            (IEnumerable<OpenXmlElement>?)run._run?.ChildElements ?? Array.Empty<OpenXmlElement>();
                        foreach (OpenXmlElement child in children) {
                            foreach (FootnoteReference reference in child is FootnoteReference directFootnote
                                ? new[] { directFootnote } : child.Descendants<FootnoteReference>()) {
                                if (!options.ExportFootnotes) continue;
                                var note = new WordFootNote(run._document, run._paragraph,
                                    new Run((FootnoteReference)reference.CloneNode(true)));
                                long id = note.ReferenceId ?? 0;
                                if (footnoteMap.ContainsKey(id)) continue;
                                int number = footnotes.Count + 1;
                                footnoteMap.Add(id, number);
                                footnotes.Add((number, note));
                            }
                            foreach (EndnoteReference reference in child is EndnoteReference directEndnote
                                ? new[] { directEndnote } : child.Descendants<EndnoteReference>()) {
                                if (!options.ExportEndnotes) continue;
                                var note = new WordEndNote(run._document, run._paragraph,
                                    new Run((EndnoteReference)reference.CloneNode(true)));
                                long id = note.ReferenceId ?? 0;
                                if (endnoteMap.ContainsKey(id)) continue;
                                int number = endnotes.Count + 1;
                                endnoteMap.Add(id, number);
                                endnotes.Add((number, note));
                            }
                        }
                    }
                }
            }
        }

        private static bool TryReplaceLastNodeWithAbbreviation(string? characterStyleId, IDocument htmlDoc, List<INode> nodes, IEnumerable<string?>? noteParagraphs) {
            if (!string.Equals(characterStyleId, "HtmlAbbr", StringComparison.OrdinalIgnoreCase) || nodes.Count == 0) {
                return false;
            }

            string text = string.Join(string.Empty, noteParagraphs ?? Enumerable.Empty<string?>());
            var abbr = CreateOutputElement(htmlDoc, "abbr");
            SetOutputAttribute(htmlDoc, abbr, "title", text, "NoteAbbreviation:title");
            var lastNode = nodes[nodes.Count - 1];
            abbr.AppendChild(lastNode);
            nodes[nodes.Count - 1] = abbr;
            return true;
        }

        private static void AppendFootnotes(
            IDocument htmlDoc,
            IElement body,
            List<(int Number, WordFootNote Note)> footnotes,
            WordToHtmlOptions options,
            CancellationToken cancellationToken, AppendWordParagraphHtml appendParagraph, Action resetParagraphFlow) {
            if (!options.ExportFootnotes || footnotes.Count == 0) {
                return;
            }

            foreach (var (number, note) in footnotes) {
                cancellationToken.ThrowIfCancellationRequested();
                foreach (var paragraph in note.Paragraphs?.Skip(1).GroupBy(item => item._paragraph).Select(group => group.First()) ?? Enumerable.Empty<WordParagraph>())
                    ReserveOutputCharacters(htmlDoc, MeasureOutputContentCharacters(paragraph._paragraph, options.TrackedChangePolicy),
                        "Referenced footnote content exceeds the configured HTML output-character limit before DOM construction.",
                        "Footnote:" + number.ToString(System.Globalization.CultureInfo.InvariantCulture));
            }

            var footSection = CreateOutputElement(htmlDoc, "section");
            SetOutputAttribute(htmlDoc, footSection, "class", "footnotes", "Footnotes:class");
            var hr = CreateOutputElement(htmlDoc, "hr");
            footSection.AppendChild(hr);
            var ol = CreateOutputElement(htmlDoc, "ol");
            foreach (var (number, note) in footnotes) {
                cancellationToken.ThrowIfCancellationRequested();
                var li = CreateOutputElement(htmlDoc, "li");
                SetOutputAttribute(htmlDoc, li, "id", "fn" + number.ToString(System.Globalization.CultureInfo.InvariantCulture), "Footnote:id");
                AppendNoteParagraphs(htmlDoc, li, note.Paragraphs?.Skip(1), appendParagraph, resetParagraphFlow);
                ol.AppendChild(li);
            }
            footSection.AppendChild(ol);
            body.AppendChild(footSection);
        }

        private static void AppendEndnotes(
            IDocument htmlDoc,
            IElement body,
            List<(int Number, WordEndNote Note)> endnotes,
            WordToHtmlOptions options,
            CancellationToken cancellationToken, AppendWordParagraphHtml appendParagraph, Action resetParagraphFlow) {
            if (!options.ExportEndnotes || endnotes.Count == 0) {
                return;
            }

            foreach (var (number, note) in endnotes) {
                cancellationToken.ThrowIfCancellationRequested();
                foreach (var paragraph in note.Paragraphs?.Skip(1).GroupBy(item => item._paragraph).Select(group => group.First()) ?? Enumerable.Empty<WordParagraph>())
                    ReserveOutputCharacters(htmlDoc, MeasureOutputContentCharacters(paragraph._paragraph, options.TrackedChangePolicy),
                        "Referenced endnote content exceeds the configured HTML output-character limit before DOM construction.",
                        "Endnote:" + number.ToString(System.Globalization.CultureInfo.InvariantCulture));
            }

            var endSection = CreateOutputElement(htmlDoc, "section");
            SetOutputAttribute(htmlDoc, endSection, "class", "endnotes", "Endnotes:class");
            var hr = CreateOutputElement(htmlDoc, "hr");
            endSection.AppendChild(hr);
            var ol = CreateOutputElement(htmlDoc, "ol");
            foreach (var (number, note) in endnotes) {
                cancellationToken.ThrowIfCancellationRequested();
                var li = CreateOutputElement(htmlDoc, "li");
                SetOutputAttribute(htmlDoc, li, "id", "en" + number.ToString(System.Globalization.CultureInfo.InvariantCulture), "Endnote:id");
                AppendNoteParagraphs(htmlDoc, li, note.Paragraphs?.Skip(1), appendParagraph, resetParagraphFlow);
                ol.AppendChild(li);
            }
            endSection.AppendChild(ol);
            body.AppendChild(endSection);
        }

        private static void AppendNoteParagraphs(IDocument htmlDoc, IElement li, IEnumerable<WordParagraph>? noteParagraphs,
            AppendWordParagraphHtml appendParagraph, Action resetParagraphFlow) {
            var paragraphs = noteParagraphs?.GroupBy(paragraph => paragraph._paragraph).Select(group => group.First()).ToArray() ?? Array.Empty<WordParagraph>();
            resetParagraphFlow();
            if (paragraphs.Length == 0) {
                li.AppendChild(CreateOutputElement(htmlDoc, "p"));
            }
            foreach (var paragraph in paragraphs) appendParagraph(li, paragraph);
            resetParagraphFlow();
        }
    }
}
