using System.Globalization;
using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.Word;

/// <summary>
/// Reads the plain text of a DOCX body with one forward pass over its main part, without building the Word document
/// model or loading the Open XML SDK. It writes the lines the Reader's text-and-metadata projection writes: each body
/// paragraph's visible run text, and each table as a column line followed by its rows with tab-separated cells. Any
/// construct whose text depends on the full model (headers and footers, notes, fields, content controls, drawings,
/// revisions, markup compatibility, merged or nested table cells) makes it decline.
/// </summary>
internal static class WordStreamedBodyText {
    private const int W = OfficeXmlNamespaces.WordprocessingML;

    /// <summary>Returns the body lines, or null when the body uses a construct this pass does not read.</summary>
    /// <exception cref="InvalidDataException">The package is not a readable Open XML package.</exception>
    /// <exception cref="System.Xml.XmlException">The main document part is not well-formed.</exception>
    internal static List<string>? TryReadLines(Stream stream, int maxChars, int maxTableRows, long? maxCharactersInPart,
        CancellationToken cancellationToken) {
        using OfficeOpcPackage package = OfficeOpcPackage.Open(stream, leaveOpen: true);
        string? main = package.GetRelatedPartName(null, OfficeOpcPackage.OfficeDocumentRelationshipType);
        if (main == null) return null;
        var context = new Context(maxTableRows, cancellationToken);
        OfficeXmlPullReader reader = package.OpenXml(main, out OfficePooledBytes content, processMarkupCompatibility: false);
        using (content) {
            // The model load caps a part's characters; a part that may exceed the cap declines and the model reports it.
            if (maxCharactersInPart > 0 && content.Length > maxCharactersInPart) return null;
            if (!ReadBody(reader, context)) return null;
        }
        foreach (string line in context.Lines) {
            if (line.Length > maxChars) return null;
        }
        return context.Lines;
    }

    private sealed class Context {
        internal Context(int maxTableRows, CancellationToken cancellationToken) {
            MaxTableRows = maxTableRows;
            CancellationToken = cancellationToken;
        }

        internal int MaxTableRows { get; }
        internal CancellationToken CancellationToken { get; }
        internal List<string> Lines { get; } = new();
        internal OfficeCharBuffer Paragraph { get; } = new();
    }

    private static bool ReadBody(OfficeXmlPullReader reader, Context context) {
        while (reader.Read()) {
            if (reader.NodeKind == OfficeXmlNodeKind.StartElement && reader.NamespaceId == W && reader.LocalName == "body") {
                return ReadBlocks(reader, context);
            }
        }
        return false;
    }

    private static bool ReadBlocks(OfficeXmlPullReader reader, Context context) {
        if (reader.IsEmptyElement) return true;
        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeKind == OfficeXmlNodeKind.EndElement && reader.Depth == depth) return true;
            if (reader.NodeKind != OfficeXmlNodeKind.StartElement) continue;
            if (reader.NamespaceId != W) return false;
            context.CancellationToken.ThrowIfCancellationRequested();
            switch (reader.LocalName) {
                case "p":
                    if (!ReadParagraph(reader, context, out string? text)) return false;
                    context.Lines.Add(text!);
                    break;
                case "tbl":
                    if (!ReadTable(reader, context)) return false;
                    break;
                case "sectPr":
                    if (!ReadSectionProperties(reader)) return false;
                    break;
                case "bookmarkStart": case "bookmarkEnd": case "proofErr":
                    reader.SkipElement();
                    break;
                default:
                    return false;
            }
        }
        return false;
    }

    // Running headers and footers are read into the text after each section, so a section that has them declines.
    private static bool ReadSectionProperties(OfficeXmlPullReader reader) {
        if (reader.IsEmptyElement) return true;
        int depth = reader.Depth;
        while (reader.Read() && !(reader.NodeKind == OfficeXmlNodeKind.EndElement && reader.Depth == depth)) {
            if (reader.NodeKind == OfficeXmlNodeKind.StartElement && reader.NamespaceId == W
                && reader.LocalName is "headerReference" or "footerReference") return false;
        }
        return true;
    }

    private static bool ReadParagraph(OfficeXmlPullReader reader, Context context, out string? text) {
        text = null;
        OfficeCharBuffer builder = context.Paragraph;
        builder.Clear();
        if (!reader.IsEmptyElement) {
            int depth = reader.Depth;
            bool closed = false;
            while (reader.Read()) {
                if (reader.NodeKind == OfficeXmlNodeKind.EndElement && reader.Depth == depth) { closed = true; break; }
                if (reader.NodeKind != OfficeXmlNodeKind.StartElement) continue;
                if (reader.NamespaceId != W) return false;
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
                        reader.SkipElement();
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
    private static bool ReadParagraphProperties(OfficeXmlPullReader reader) {
        if (reader.IsEmptyElement) return true;
        int depth = reader.Depth;
        while (reader.Read() && !(reader.NodeKind == OfficeXmlNodeKind.EndElement && reader.Depth == depth)) {
            if (reader.NodeKind == OfficeXmlNodeKind.StartElement && reader.NamespaceId == W && reader.LocalName == "sectPr"
                && !ReadSectionProperties(reader)) return false;
        }
        return true;
    }

    // Matches WordParagraph.Text for a run: text honours xml:space, tabs and breaks become characters, and a hidden
    // run that only carries a comment reference contributes nothing.
    private static bool ReadRun(OfficeXmlPullReader reader, OfficeCharBuffer run) {
        if (reader.IsEmptyElement) return true;
        int runStart = run.Length;
        bool hidden = false, commentReference = false;
        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeKind == OfficeXmlNodeKind.EndElement && reader.Depth == depth) {
                if (hidden && commentReference) run.Truncate(runStart);
                return true;
            }
            if (reader.NodeKind != OfficeXmlNodeKind.StartElement) continue;
            if (reader.NamespaceId != W) return false;
            switch (reader.LocalName) {
                case "rPr":
                    if (!reader.IsEmptyElement) {
                        int propertiesDepth = reader.Depth;
                        while (reader.Read() && !(reader.NodeKind == OfficeXmlNodeKind.EndElement && reader.Depth == propertiesDepth)) {
                            if (reader.NodeKind == OfficeXmlNodeKind.StartElement && reader.Depth == propertiesDepth + 1
                                && reader.NamespaceId == W && reader.LocalName == "vanish") hidden = IsOn(reader);
                        }
                    }
                    break;
                case "t":
                    bool preserve = reader.TryGetAttribute(OfficeXmlNamespaces.Xml, "space", out string space) && space == "preserve";
                    if (!ReadText(reader, run, preserve)) return false;
                    break;
                case "tab": case "ptab":
                    run.Append('\t');
                    reader.SkipElement();
                    break;
                case "cr":
                    run.Append('\n');
                    reader.SkipElement();
                    break;
                case "noBreakHyphen":
                    run.Append('\u2011');
                    reader.SkipElement();
                    break;
                case "softHyphen":
                    run.Append('\u00ad');
                    reader.SkipElement();
                    break;
                case "br":
                    string? type = reader.TryGetAttribute(W, "type", out string typeValue) ? typeValue : null;
                    run.Append(type is null or "textWrapping" ? '\n' : '\u2028');
                    reader.SkipElement();
                    break;
                case "commentReference":
                    commentReference = true;
                    reader.SkipElement();
                    break;
                case "lastRenderedPageBreak":
                    reader.SkipElement();
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

    private static bool SkipWithoutText(OfficeXmlPullReader reader) {
        if (reader.IsEmptyElement) return true;
        int depth = reader.Depth;
        while (reader.Read() && !(reader.NodeKind == OfficeXmlNodeKind.EndElement && reader.Depth == depth)) {
            if (reader.NodeKind == OfficeXmlNodeKind.StartElement && (reader.NamespaceId == W
                || reader.LocalName is "txbx" or "txbxContent" or "AlternateContent")) return false;
        }
        return true;
    }

    // Appends the element's text; without xml:space="preserve" the XML white space at either edge is dropped.
    private static bool ReadText(OfficeXmlPullReader reader, OfficeCharBuffer builder, bool preserve) {
        if (reader.IsEmptyElement) return true;
        int start = builder.Length;
        int depth = reader.Depth;
        while (reader.Read() && !(reader.NodeKind == OfficeXmlNodeKind.EndElement && reader.Depth == depth)) {
            if (reader.NodeKind == OfficeXmlNodeKind.Text) {
                reader.AppendText(builder);
            } else if (reader.NodeKind == OfficeXmlNodeKind.StartElement) {
                return false;
            }
        }
        if (!preserve) {
            int end = builder.Length;
            while (end > start && IsXmlEdgeWhitespace(builder[end - 1])) end--;
            builder.Truncate(end);
            int first = start;
            while (first < end && IsXmlEdgeWhitespace(builder[first])) first++;
            if (first > start) builder.Remove(start, first - start);
        }
        return true;
    }

    private static bool IsXmlEdgeWhitespace(char value) => value is ' ' or '\t' or '\r' or '\n';

    // Matches WordTableProjection: a column line (header row values, or "Column N"), then one line per row with
    // tab-separated cells; a cell is its non-blank paragraphs joined by spaces.
    private static bool ReadTable(OfficeXmlPullReader reader, Context context) {
        if (reader.IsEmptyElement) return true;
        int depth = reader.Depth;
        bool repeatsHeader = false;
        var rows = new List<List<string>>();
        while (reader.Read()) {
            if (reader.NodeKind == OfficeXmlNodeKind.EndElement && reader.Depth == depth) break;
            if (reader.NodeKind != OfficeXmlNodeKind.StartElement) continue;
            if (reader.NamespaceId != W) return false;
            context.CancellationToken.ThrowIfCancellationRequested();
            switch (reader.LocalName) {
                case "tblPr":
                case "tblGrid":
                    // The model sizes the table from its widest row, not its grid.
                    reader.SkipElement();
                    break;
                case "tr":
                    var cells = new List<string>();
                    if (!ReadRow(reader, context, cells, out bool header)) return false;
                    if (rows.Count == 0) repeatsHeader = header;
                    rows.Add(cells);
                    break;
                case "bookmarkStart": case "bookmarkEnd": case "proofErr":
                    reader.SkipElement();
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

    private static bool ReadRow(OfficeXmlPullReader reader, Context context, List<string> cells, out bool repeatsHeader) {
        repeatsHeader = false;
        if (reader.IsEmptyElement) return true;
        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeKind == OfficeXmlNodeKind.EndElement && reader.Depth == depth) return true;
            if (reader.NodeKind != OfficeXmlNodeKind.StartElement) continue;
            if (reader.NamespaceId != W) return false;
            switch (reader.LocalName) {
                case "trPr": case "tblPrEx":
                    if (!reader.IsEmptyElement) {
                        bool rowProperties = reader.LocalName == "trPr";
                        bool headerSeen = false;
                        int propertiesDepth = reader.Depth;
                        while (reader.Read() && !(reader.NodeKind == OfficeXmlNodeKind.EndElement && reader.Depth == propertiesDepth)) {
                            if (reader.NodeKind != OfficeXmlNodeKind.StartElement || reader.NamespaceId != W) continue;
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
                    context.CancellationToken.ThrowIfCancellationRequested();
                    if (!ReadCell(reader, context, out string? value)) return false;
                    cells.Add(value!);
                    break;
                case "bookmarkStart": case "bookmarkEnd": case "proofErr":
                    reader.SkipElement();
                    break;
                default:
                    return false;
            }
        }
        return false;
    }

    private static bool ReadCell(OfficeXmlPullReader reader, Context context, out string? value) {
        value = null;
        var parts = new List<string>();
        if (!reader.IsEmptyElement) {
            int depth = reader.Depth;
            bool closed = false;
            while (reader.Read()) {
                if (reader.NodeKind == OfficeXmlNodeKind.EndElement && reader.Depth == depth) { closed = true; break; }
                if (reader.NodeKind != OfficeXmlNodeKind.StartElement) continue;
                if (reader.NamespaceId != W) return false;
                switch (reader.LocalName) {
                    case "tcPr":
                        if (!reader.IsEmptyElement) {
                            int propertiesDepth = reader.Depth;
                            while (reader.Read() && !(reader.NodeKind == OfficeXmlNodeKind.EndElement && reader.Depth == propertiesDepth)) {
                                if (reader.NodeKind == OfficeXmlNodeKind.StartElement && reader.NamespaceId == W
                                    && reader.LocalName is "gridSpan" or "vMerge" or "hMerge") return false;
                            }
                        }
                        break;
                    case "p":
                        if (!ReadParagraph(reader, context, out string? text)) return false;
                        if (!string.IsNullOrWhiteSpace(text)) parts.Add(text!);
                        break;
                    case "bookmarkStart": case "bookmarkEnd": case "proofErr":
                        reader.SkipElement();
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

    private static bool IsOn(OfficeXmlPullReader reader) =>
        !reader.TryGetAttribute(W, "val", out string value) || value is "1" or "true" or "on";

}
