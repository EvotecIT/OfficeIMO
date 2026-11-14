using System.Globalization;
using System.Text;
using System.Xml;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.Reader.Word;

/// <summary>
/// Reads the plain text of a DOCX body for the text projection with one forward pass over its main part, without
/// building the Word document model. It writes the lines the text-and-metadata projection's GetText() writes: each
/// body paragraph's visible run text, and each table as a column line followed by its rows with tab-separated cells.
/// Any construct whose text depends on the full model (headers and footers, notes, fields, content controls, drawings,
/// revisions, merged or nested table cells, long paragraphs) makes it decline; the caller then builds the model and
/// the read scope reduces that result to the same text chunks.
/// </summary>
internal static class WordStreamingTextReader {
    private const string WordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string XmlNamespace = "http://www.w3.org/XML/1998/namespace";

    /// <summary>Whether the read asked for the text projection and nothing this pass cannot honor.</summary>
    internal static bool Applies(ReaderOptions readerOptions) =>
        DocumentReaderEngine.IsTextProjection
        && string.IsNullOrEmpty(readerOptions.OpenPassword)
        && readerOptions.Inspection == null
        && FormatInternals.OpenXmlReaderAdmission.Policy(readerOptions, global::OfficeIMO.Word.WordLoadOptions.DefaultMaxInputBytes) == null;

    /// <summary>Returns the streamed result, or null to read the document through the Word model.</summary>
    internal static OfficeDocumentReadResult? TryRead(Stream stream, string sourceName, ReaderOptions readerOptions, CancellationToken cancellationToken) {
        if (!stream.CanSeek) return null;
        long start = stream.Position;
        if (stream.Length - start > global::OfficeIMO.Word.WordLoadOptions.DefaultMaxInputBytes) return null;
        try {
            List<string>? lines = ReadLines(stream, readerOptions, cancellationToken);
            return lines == null ? null : CreateResult(lines, sourceName);
        } catch (OperationCanceledException) {
            throw;
        } catch (Exception exception) when (exception is not OutOfMemoryException) {
            // The Word model reports damaged or unusual packages with its own errors.
            return null;
        } finally {
            stream.Position = start;
        }
    }

    private static List<string>? ReadLines(Stream stream, ReaderOptions readerOptions, CancellationToken cancellationToken) {
        using var package = WordprocessingDocument.Open(stream, false);
        MainDocumentPart? main = package.MainDocumentPart;
        if (main == null) return null;
        var context = new Context(readerOptions, cancellationToken);
        // The model load caps a part at the same character count; a larger part declines and the model reports it.
        var settings = new XmlReaderSettings {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            MaxCharactersInDocument = readerOptions.OpenXmlMaxCharactersInPart ?? 0
        };
        using (Stream part = main.GetStream(FileMode.Open, FileAccess.Read))
        using (XmlReader reader = XmlReader.Create(part, settings)) {
            if (!ReadBody(reader, context)) return null;
        }
        foreach (string line in context.Lines) {
            if (line.Length > readerOptions.MaxChars) return null;
        }
        return context.Lines;
    }

    private static OfficeDocumentReadResult CreateResult(List<string> lines, string sourceName) {
        string fileName = Path.GetFileName(sourceName);
        var chunks = new List<ReaderChunk>();
        foreach (string line in lines) {
            if (line.Length == 0) continue;
            int index = chunks.Count;
            chunks.Add(new ReaderChunk {
                Id = "word:" + fileName + ":b" + index.ToString("D4", CultureInfo.InvariantCulture),
                Kind = ReaderInputKind.Word,
                Location = new ReaderLocation { Path = sourceName, BlockIndex = index },
                Text = line
            });
        }
        return DocumentReaderEngine.CreateTextResult(chunks, ReaderInputKind.Word,
            new[] { OfficeDocumentReaderBuilderWordExtensions.HandlerId, "officeimo.reader.word.streamed-text" });
    }

    private sealed class Context {
        internal Context(ReaderOptions options, CancellationToken cancellationToken) {
            MaxTableRows = options.MaxTableRows;
            CancellationToken = cancellationToken;
        }

        internal int MaxTableRows { get; }
        internal CancellationToken CancellationToken { get; }
        internal List<string> Lines { get; } = new();
    }

    private static bool ReadBody(XmlReader reader, Context context) {
        while (reader.Read()) {
            if (reader.NodeType == XmlNodeType.Element && reader.NamespaceURI == WordNamespace && reader.LocalName == "body") {
                return ReadBlocks(reader, context);
            }
        }
        return false;
    }

    private static bool ReadBlocks(XmlReader reader, Context context) {
        if (reader.IsEmptyElement) return true;
        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth) return true;
            if (reader.NodeType != XmlNodeType.Element) continue;
            if (reader.NamespaceURI != WordNamespace) return false;
            context.CancellationToken.ThrowIfCancellationRequested();
            switch (reader.LocalName) {
                case "p":
                    if (!ReadParagraph(reader, out string? text)) return false;
                    context.Lines.Add(text!);
                    break;
                case "tbl":
                    if (!ReadTable(reader, context)) return false;
                    break;
                case "sectPr":
                    if (!ReadSectionProperties(reader)) return false;
                    break;
                case "bookmarkStart": case "bookmarkEnd": case "proofErr":
                    Skip(reader);
                    break;
                default:
                    return false;
            }
        }
        return false;
    }

    // Running headers and footers are read into the text after each section, so a section that has them declines.
    private static bool ReadSectionProperties(XmlReader reader) {
        if (reader.IsEmptyElement) return true;
        int depth = reader.Depth;
        while (reader.Read() && !(reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth)) {
            if (reader.NodeType == XmlNodeType.Element && reader.NamespaceURI == WordNamespace
                && reader.LocalName is "headerReference" or "footerReference") return false;
        }
        return true;
    }

    private static bool ReadParagraph(XmlReader reader, out string? text) {
        text = null;
        var builder = new StringBuilder();
        if (!reader.IsEmptyElement) {
            int depth = reader.Depth;
            bool closed = false;
            while (reader.Read()) {
                if (reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth) { closed = true; break; }
                if (reader.NodeType != XmlNodeType.Element) continue;
                if (reader.NamespaceURI != WordNamespace) return false;
                switch (reader.LocalName) {
                    case "pPr":
                        if (!ReadParagraphProperties(reader)) return false;
                        break;
                    case "r":
                        if (!ReadRun(reader, builder)) return false;
                        break;
                    case "hyperlink":
                        // Its runs are part of the paragraph's text; the loop reads them as it descends.
                        break;
                    case "bookmarkStart": case "bookmarkEnd": case "proofErr": case "permStart": case "permEnd":
                    case "commentRangeStart": case "commentRangeEnd":
                    case "moveFromRangeStart": case "moveFromRangeEnd": case "moveToRangeStart": case "moveToRangeEnd":
                        Skip(reader);
                        break;
                    default:
                        return false;
                }
            }
            if (!closed) return false;
        }
        text = builder.ToString();
        return true;
    }

    // List numbering never reaches the text; a section break inside a paragraph is checked like the body's.
    private static bool ReadParagraphProperties(XmlReader reader) {
        if (reader.IsEmptyElement) return true;
        int depth = reader.Depth;
        while (reader.Read() && !(reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth)) {
            if (reader.NodeType == XmlNodeType.Element && reader.NamespaceURI == WordNamespace && reader.LocalName == "sectPr"
                && !ReadSectionProperties(reader)) return false;
        }
        return true;
    }

    // Matches WordParagraph.Text for a run: text honours xml:space, tabs and breaks become characters, and a hidden
    // run that only carries a comment reference contributes nothing.
    private static bool ReadRun(XmlReader reader, StringBuilder run) {
        if (reader.IsEmptyElement) return true;
        int runStart = run.Length;
        bool hidden = false, commentReference = false;
        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth) {
                if (hidden && commentReference) run.Length = runStart;
                return true;
            }
            if (reader.NodeType != XmlNodeType.Element) continue;
            if (reader.NamespaceURI != WordNamespace) return false;
            switch (reader.LocalName) {
                case "rPr":
                    if (!reader.IsEmptyElement) {
                        int propertiesDepth = reader.Depth;
                        while (reader.Read() && !(reader.NodeType == XmlNodeType.EndElement && reader.Depth == propertiesDepth)) {
                            if (reader.NodeType == XmlNodeType.Element && reader.Depth == propertiesDepth + 1
                                && reader.NamespaceURI == WordNamespace && reader.LocalName == "vanish") hidden = IsOn(reader);
                        }
                    }
                    break;
                case "t":
                    bool preserve = reader.GetAttribute("space", XmlNamespace) == "preserve";
                    if (!ReadText(reader, run, preserve)) return false;
                    break;
                case "tab": case "ptab":
                    run.Append('\t');
                    Skip(reader);
                    break;
                case "cr":
                    run.Append('\n');
                    Skip(reader);
                    break;
                case "noBreakHyphen":
                    run.Append('\u2011');
                    Skip(reader);
                    break;
                case "softHyphen":
                    run.Append('\u00ad');
                    Skip(reader);
                    break;
                case "br":
                    string? type = reader.GetAttribute("type", WordNamespace);
                    run.Append(type is null or "textWrapping" ? '\n' : '\u2028');
                    Skip(reader);
                    break;
                case "commentReference":
                    commentReference = true;
                    Skip(reader);
                    break;
                case "lastRenderedPageBreak":
                    Skip(reader);
                    break;
                case "drawing":
                    // A picture adds no text; a drawing that holds a text box or other text declines.
                    if (!SkipWithoutText(reader)) return false;
                    break;
                default:
                    return false;
            }
        }
        return false;
    }

    private static bool SkipWithoutText(XmlReader reader) {
        if (reader.IsEmptyElement) return true;
        int depth = reader.Depth;
        while (reader.Read() && !(reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth)) {
            if (reader.NodeType == XmlNodeType.Element && (reader.NamespaceURI == WordNamespace
                || reader.LocalName is "txbx" or "txbxContent" or "AlternateContent")) return false;
        }
        return true;
    }

    // Appends the element's text; without xml:space="preserve" the XML white space at either edge is dropped.
    private static bool ReadText(XmlReader reader, StringBuilder builder, bool preserve) {
        if (reader.IsEmptyElement) return true;
        int start = builder.Length;
        int depth = reader.Depth;
        while (reader.Read() && !(reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth)) {
            if (reader.NodeType is XmlNodeType.Text or XmlNodeType.Whitespace or XmlNodeType.SignificantWhitespace or XmlNodeType.CDATA) {
                builder.Append(reader.Value);
            } else if (reader.NodeType == XmlNodeType.Element) {
                return false;
            }
        }
        if (!preserve) {
            int end = builder.Length;
            while (end > start && IsXmlEdgeWhitespace(builder[end - 1])) end--;
            builder.Length = end;
            int first = start;
            while (first < end && IsXmlEdgeWhitespace(builder[first])) first++;
            if (first > start) builder.Remove(start, first - start);
        }
        return true;
    }

    private static bool IsXmlEdgeWhitespace(char value) => value is ' ' or '\t' or '\r' or '\n';

    // Matches WordTableProjection: a column line (header row values, or "Column N"), then one line per row with
    // tab-separated cells; a cell is its non-blank paragraphs joined by spaces.
    private static bool ReadTable(XmlReader reader, Context context) {
        if (reader.IsEmptyElement) return true;
        int depth = reader.Depth;
        bool repeatsHeader = false;
        var rows = new List<List<string>>();
        while (reader.Read()) {
            if (reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth) break;
            if (reader.NodeType != XmlNodeType.Element) continue;
            if (reader.NamespaceURI != WordNamespace) return false;
            context.CancellationToken.ThrowIfCancellationRequested();
            switch (reader.LocalName) {
                case "tblPr":
                case "tblGrid":
                    // The model sizes the table from its widest row, not its grid.
                    Skip(reader);
                    break;
                case "tr":
                    var cells = new List<string>();
                    if (!ReadRow(reader, cells, out bool header, context.CancellationToken)) return false;
                    if (rows.Count == 0) repeatsHeader = header;
                    rows.Add(cells);
                    break;
                case "bookmarkStart": case "bookmarkEnd": case "proofErr":
                    Skip(reader);
                    break;
                default:
                    return false;
            }
        }
        int columns = 0;
        foreach (List<string> row in rows) columns = Math.Max(columns, row.Count);
        bool hasHeader = repeatsHeader && rows.Count > 0;
        if (columns > 0) {
            var names = new string[columns];
            for (int index = 0; index < columns; index++) {
                string value = hasHeader ? Cell(rows[0], index) : string.Empty;
                names[index] = string.IsNullOrWhiteSpace(value) ? "Column " + (index + 1).ToString(CultureInfo.InvariantCulture) : value;
            }
            context.Lines.Add(string.Join("\t", names));
        }
        int written = 0;
        for (int rowIndex = hasHeader ? 1 : 0; rowIndex < rows.Count; rowIndex++) {
            if (context.MaxTableRows > 0 && written == context.MaxTableRows) break;
            var values = new string[columns];
            for (int index = 0; index < columns; index++) values[index] = Cell(rows[rowIndex], index);
            context.Lines.Add(string.Join("\t", values));
            written++;
        }
        return true;
    }

    private static string Cell(List<string> row, int index) => index < row.Count ? row[index] : string.Empty;

    private static bool ReadRow(XmlReader reader, List<string> cells, out bool repeatsHeader, CancellationToken cancellationToken) {
        repeatsHeader = false;
        if (reader.IsEmptyElement) return true;
        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth) return true;
            if (reader.NodeType != XmlNodeType.Element) continue;
            if (reader.NamespaceURI != WordNamespace) return false;
            switch (reader.LocalName) {
                case "trPr": case "tblPrEx":
                    if (!reader.IsEmptyElement) {
                        bool rowProperties = reader.LocalName == "trPr";
                        bool headerSeen = false;
                        int propertiesDepth = reader.Depth;
                        while (reader.Read() && !(reader.NodeType == XmlNodeType.EndElement && reader.Depth == propertiesDepth)) {
                            if (reader.NodeType != XmlNodeType.Element || reader.NamespaceURI != WordNamespace) continue;
                            // Revised row properties are a revision; the model read handles those.
                            if (reader.LocalName is "gridBefore" or "gridAfter" or "trPrChange" or "tblPrExChange") return false;
                            // Like the model, only the first direct tblHeader of the row's own properties counts.
                            if (rowProperties && !headerSeen && reader.LocalName == "tblHeader" && reader.Depth == propertiesDepth + 1) {
                                repeatsHeader = IsOn(reader);
                                headerSeen = true;
                            }
                        }
                    }
                    break;
                case "tc":
                    cancellationToken.ThrowIfCancellationRequested();
                    if (!ReadCell(reader, out string? value)) return false;
                    cells.Add(value!);
                    break;
                case "bookmarkStart": case "bookmarkEnd": case "proofErr":
                    Skip(reader);
                    break;
                default:
                    return false;
            }
        }
        return false;
    }

    private static bool ReadCell(XmlReader reader, out string? value) {
        value = null;
        var parts = new List<string>();
        if (!reader.IsEmptyElement) {
            int depth = reader.Depth;
            bool closed = false;
            while (reader.Read()) {
                if (reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth) { closed = true; break; }
                if (reader.NodeType != XmlNodeType.Element) continue;
                if (reader.NamespaceURI != WordNamespace) return false;
                switch (reader.LocalName) {
                    case "tcPr":
                        if (!reader.IsEmptyElement) {
                            int propertiesDepth = reader.Depth;
                            while (reader.Read() && !(reader.NodeType == XmlNodeType.EndElement && reader.Depth == propertiesDepth)) {
                                if (reader.NodeType == XmlNodeType.Element && reader.NamespaceURI == WordNamespace
                                    && reader.LocalName is "gridSpan" or "vMerge" or "hMerge") return false;
                            }
                        }
                        break;
                    case "p":
                        if (!ReadParagraph(reader, out string? text)) return false;
                        if (!string.IsNullOrWhiteSpace(text)) parts.Add(text!);
                        break;
                    case "bookmarkStart": case "bookmarkEnd": case "proofErr":
                        Skip(reader);
                        break;
                    default:
                        return false;
                }
            }
            if (!closed) return false;
        }
        value = string.Join(" ", parts);
        return true;
    }

    private static bool IsOn(XmlReader reader) => reader.GetAttribute("val", WordNamespace) is null or "1" or "true" or "on";

    private static void Skip(XmlReader reader) {
        if (reader.IsEmptyElement) return;
        int depth = reader.Depth;
        while (reader.Read() && !(reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth)) { }
    }
}
