using AngleSharp.Dom;
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
            Dictionary<long, int> endnoteMap) {
            if (!processNotes) {
                return false;
            }

            if (options.ExportFootnotes && run.FootNote != null) {
                var note = run.FootNote;
                if (string.Equals(run.CharacterStyleId, "HtmlAbbr", StringComparison.OrdinalIgnoreCase) && nodes.Count > 0) {
                    string text = string.Join(string.Empty, note.Paragraphs?.Skip(1).Select(r => r.Text) ?? Enumerable.Empty<string>());
                    var abbr = htmlDoc.CreateElement("abbr");
                    abbr.SetAttribute("title", text);
                    var lastNode = nodes[nodes.Count - 1];
                    abbr.AppendChild(lastNode);
                    nodes[nodes.Count - 1] = abbr;
                } else {
                    long id = note.ReferenceId ?? 0;
                    if (!footnoteMap.TryGetValue(id, out int number)) {
                        number = footnotes.Count + 1;
                        footnoteMap[id] = number;
                        footnotes.Add((number, note));
                    }
                    var sup = htmlDoc.CreateElement("sup");
                    var a = htmlDoc.CreateElement("a");
                    a.SetAttribute("href", $"#fn{number}");
                    a.SetAttribute("id", $"fnref{number}");
                    a.TextContent = number.ToString();
                    sup.AppendChild(a);
                    nodes.Add(sup);
                }

                return true;
            }

            if (options.ExportEndnotes && run.EndNote != null) {
                var note = run.EndNote;
                long id = note.ReferenceId ?? 0;
                if (!endnoteMap.TryGetValue(id, out int number)) {
                    number = endnotes.Count + 1;
                    endnoteMap[id] = number;
                    endnotes.Add((number, note));
                }
                var sup = htmlDoc.CreateElement("sup");
                var a = htmlDoc.CreateElement("a");
                a.SetAttribute("href", $"#en{number}");
                a.SetAttribute("id", $"enref{number}");
                a.TextContent = number.ToString();
                sup.AppendChild(a);
                nodes.Add(sup);
                return true;
            }

            return false;
        }

        private static void AppendFootnotes(
            IDocument htmlDoc,
            IElement body,
            List<(int Number, WordFootNote Note)> footnotes,
            WordToHtmlOptions options,
            CancellationToken cancellationToken) {
            if (!options.ExportFootnotes || footnotes.Count == 0) {
                return;
            }

            var footSection = htmlDoc.CreateElement("section");
            footSection.SetAttribute("class", "footnotes");
            var hr = htmlDoc.CreateElement("hr");
            footSection.AppendChild(hr);
            var ol = htmlDoc.CreateElement("ol");
            foreach (var (number, note) in footnotes) {
                cancellationToken.ThrowIfCancellationRequested();
                var li = htmlDoc.CreateElement("li");
                li.SetAttribute("id", $"fn{number}");
                var p = htmlDoc.CreateElement("p");
                string text = string.Join(string.Empty, note.Paragraphs?.Skip(1).Select(r => r.Text) ?? Enumerable.Empty<string>());
                p.TextContent = text;
                li.AppendChild(p);
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
            CancellationToken cancellationToken) {
            if (!options.ExportEndnotes || endnotes.Count == 0) {
                return;
            }

            var endSection = htmlDoc.CreateElement("section");
            endSection.SetAttribute("class", "endnotes");
            var hr = htmlDoc.CreateElement("hr");
            endSection.AppendChild(hr);
            var ol = htmlDoc.CreateElement("ol");
            foreach (var (number, note) in endnotes) {
                cancellationToken.ThrowIfCancellationRequested();
                var li = htmlDoc.CreateElement("li");
                li.SetAttribute("id", $"en{number}");
                var p = htmlDoc.CreateElement("p");
                string text = string.Join(string.Empty, note.Paragraphs?.Skip(1).Select(r => r.Text) ?? Enumerable.Empty<string>());
                p.TextContent = text;
                li.AppendChild(p);
                ol.AppendChild(li);
            }
            endSection.AppendChild(ol);
            body.AppendChild(endSection);
        }
    }
}
