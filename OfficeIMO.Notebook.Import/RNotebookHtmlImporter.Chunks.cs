using System.Text;
using OfficeIMO.Html.Dom;

namespace OfficeIMO.Notebook.Import;

internal static partial class RNotebookHtmlImporter {
    private sealed class SavedChunk {
        internal StringBuilder Source { get; } = new();
        internal List<NotebookOutput> Outputs { get; } = new();
    }

    private sealed class ChunkReader {
        private readonly NotebookImportContext _context;
        private readonly List<SavedChunk> _chunks = new();
        private readonly Stack<Annotation> _annotations = new();
        private SavedChunk? _chunk;
        private int _outputs;

        internal ChunkReader(NotebookImportContext context) => _context = context;
        internal IReadOnlyList<SavedChunk> Chunks => _chunks;

        internal void Visit(HtmlNode node, bool inPreformattedText = false) {
            _context.CancellationToken.ThrowIfCancellationRequested();
            if (node.Kind == HtmlNodeKind.Comment) {
                ReadComment(node.TextContent);
                return;
            }
            Annotation? annotation = _annotations.Count == 0 ? null : _annotations.Peek();
            if (node is HtmlElement element) {
                if (element.Id == "rmd-source-code") return;
                if (IsUnsupportedSavedOutputElement(element.LocalName)) {
                    if (annotation?.Kind != AnnotationKind.Unsupported && _chunk != null) {
                        Report("notebook-import-rnb-output", "The interactive or unsupported presentation of a saved HTML element was omitted.");
                    }
                    if (IsIgnoredSavedMarkup(element.LocalName)) return;
                }
                if (annotation != null && ReadElement(element, annotation)) return;
                inPreformattedText |= element.LocalName is "pre" or "code";
            } else if (node.Kind == HtmlNodeKind.Text && annotation?.Kind is AnnotationKind.Source or AnnotationKind.Output) {
                annotation.AppendText(node.TextContent, inPreformattedText);
                return;
            }
            foreach (HtmlNode child in node.ChildNodes) Visit(child, inPreformattedText);
        }

        internal void Complete() {
            if (_chunk == null) return;
            Report("notebook-import-rnb-marker", "The rendered notebook contains an unterminated chunk annotation.");
            FinishChunk();
        }

        private void ReadComment(string content) {
            if (!TryReadMarker(content, out string label, out bool starts)) return;
            if (label == "chunk") {
                if (starts) {
                    if (_chunk != null) throw new InvalidDataException("R Notebook chunk annotations cannot be nested.");
                    if (_chunks.Count >= _context.Limits.MaxCells) throw new InvalidDataException("R Notebook chunk count exceeds the import limit.");
                    _chunk = new SavedChunk();
                } else if (_chunk != null) {
                    if (_annotations.Count > 0) {
                        Report("notebook-import-rnb-marker", "The rendered chunk contains an unterminated output annotation.");
                    }
                    FinishChunk();
                } else {
                    Report("notebook-import-rnb-marker", "A rendered chunk ending has no matching start marker.");
                }
                return;
            }
            if (_chunk == null) return;

            if (starts) {
                if (_annotations.Count >= 64) throw new InvalidDataException("R Notebook annotation nesting exceeds the import limit.");
                if (_annotations.Count > 0) FlushText(_annotations.Peek());
                AnnotationKind kind = Classify(label);
                if (_annotations.Count > 0 && _annotations.Peek().Kind == AnnotationKind.Unsupported) {
                    kind = AnnotationKind.Unsupported;
                } else if (kind == AnnotationKind.Unsupported) {
                    Report("notebook-import-rnb-output", "A saved interactive or unknown notebook output was omitted.");
                }
                _annotations.Push(new Annotation(label, kind));
                return;
            }

            if (_annotations.Count == 0 || _annotations.Peek().Label != label) {
                Report("notebook-import-rnb-marker", "A rendered output ending has no matching start marker.");
                return;
            }
            FlushText(_annotations.Pop());
        }

        private bool ReadElement(HtmlElement element, Annotation annotation) {
            if (annotation.Kind == AnnotationKind.Unsupported) return false;
            if (annotation.Kind == AnnotationKind.Source) return false;
            if (element.LocalName == "img") {
                FlushText(annotation);
                NotebookOutput? image = ReadImage(element);
                if (image != null) AddOutput(image);
                else Report("notebook-import-rnb-image", "A saved figure was omitted because it is external, unsupported, or exceeds the image limit.");
                return true;
            }
            if (element.LocalName == "table") {
                FlushText(annotation);
                NotebookOutput? table = ReadTable(element, _context.CancellationToken, out TableOmissions omissions);
                if (table == null) {
                    Report("notebook-import-rnb-table", "A saved table exceeded the bounded HTML output limits.");
                    return true;
                }

                AddOutput(table);
                if ((omissions & TableOmissions.Resource) != 0) {
                    Report("notebook-import-rnb-table-resource", "A resource inside a saved table could not be retained in the static table.");
                }
                if ((omissions & TableOmissions.Structure) != 0) {
                    Report("notebook-import-rnb-table-content", "An active or unsupported element or attribute inside a saved table could not be retained.");
                }
                return true;
            }
            if (element.LocalName == "br") {
                annotation.AppendText("\n", inPreformattedText: true);
                return true;
            }
            return false;
        }

        private void FlushText(Annotation annotation) {
            string text = annotation.TakeText();
            if (text.Length == 0) return;
            if (annotation.Kind == AnnotationKind.Source) {
                text = text.Trim('\r', '\n');
                if (string.IsNullOrWhiteSpace(text)) return;
                StringBuilder source = _chunk!.Source;
                if (source.Length > 0) source.Append('\n');
                source.Append(text);
            } else if (annotation.Kind == AnnotationKind.Output) {
                bool error = annotation.Label is "error" or "warning" or "message";
                AddOutput(NotebookOutput.CreateStream(error ? "stderr" : "stdout", text));
            }
        }

        private void AddOutput(NotebookOutput output) {
            if (_outputs >= _context.Limits.MaxTotalOutputs || _chunk!.Outputs.Count >= _context.Limits.MaxOutputsPerCell) {
                throw new InvalidDataException("R Notebook saved outputs exceed the import limit.");
            }
            _outputs++;
            _chunk.Outputs.Add(output);
        }

        private void FinishChunk() {
            foreach (Annotation unfinished in _annotations) FlushText(unfinished);
            _annotations.Clear();
            _chunks.Add(_chunk!);
            _chunk = null;
        }

        private void Report(string code, string message) => _context.Report(code, message, ChunkLocation(_chunks.Count));

        private static AnnotationKind Classify(string label) => label switch {
            "source" => AnnotationKind.Source,
            "output" or "plot" or "text" or "html" or "error" or "warning" or "message" => AnnotationKind.Output,
            _ => AnnotationKind.Unsupported
        };

        private static bool TryReadMarker(string content, out string label, out bool starts) {
            label = string.Empty;
            starts = false;
            ReadOnlySpan<char> value = content.AsSpan().TrimStart();
            if (!value.StartsWith("rnb-".AsSpan(), StringComparison.Ordinal)) return false;
            int whitespace = value.IndexOfAny(" \r\n\t".AsSpan());
            if (whitespace >= 0) value = value.Slice(0, whitespace);
            if (value.EndsWith("-begin".AsSpan(), StringComparison.Ordinal)) starts = true;
            else if (!value.EndsWith("-end".AsSpan(), StringComparison.Ordinal)) return false;
            int length = value.Length - 4 - (starts ? 6 : 4);
            if (length < 1 || length > 32) return false;
            label = value.Slice(4, length).ToString();
            return true;
        }
    }

    private enum AnnotationKind {
        Source,
        Output,
        Unsupported
    }

    private sealed class Annotation {
        private readonly StringBuilder _text = new();
        private int _textEnd;

        internal Annotation(string label, AnnotationKind kind) {
            Label = label;
            Kind = kind;
        }

        internal string Label { get; }
        internal AnnotationKind Kind { get; }
        internal void AppendText(string text, bool inPreformattedText) {
            // Preserve whitespace between text nodes and inside saved preformatted blocks.
            // Bare whitespace before or after rendered output is HTML layout.
            bool layout = Kind == AnnotationKind.Output && !inPreformattedText && string.IsNullOrWhiteSpace(text);
            if (!layout || _text.Length > 0) _text.Append(text);
            if (!layout) _textEnd = _text.Length;
        }

        internal string TakeText() {
            if (_textEnd == 0) return string.Empty;
            string text = _text.ToString(0, _textEnd);
            _text.Clear();
            _textEnd = 0;
            return text;
        }
    }
}
