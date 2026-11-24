using System.Globalization;
using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.Word;

/// <summary>
/// Reads the plain text of a DOCX with one forward pass over its main part, without building the Word document model or
/// loading the Open XML SDK. It writes the lines the Reader's text-and-metadata projection writes, in its order: each
/// section's paragraphs and tables, then that section's headers and footers, then the text of the footnotes and
/// endnotes the document refers to. A table is a column line followed by its rows with tab-separated cells.
/// </summary>
/// <remarks>
/// The text follows the Word model's inspection snapshot: field instructions are hidden and their results kept, with
/// field state carried across the paragraphs of a story; paragraphs the snapshot drops are dropped; insertions,
/// deletions, moves and smart tags contribute no text; text boxes contribute their text inline from the compatibility
/// branch the model selects; content controls contribute their content, or their runs when they hold a field; block
/// content controls, alternative-format chunks, paragraph-level math, nested tables in cells and merged-cell markup are
/// left out the way the snapshot leaves them out. Constructs whose text needs more of the model (math inside runs or
/// drawings, compatibility markup outside runs, revised row properties) make the pass decline.
/// </remarks>
internal static class WordStreamedBodyText {
    private const int W = OfficeXmlNamespaces.WordprocessingML;
    private const string HeaderRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "header";
    private const string FooterRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "footer";
    private const string FootnotesRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "footnotes";
    private const string EndnotesRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "endnotes";

    /// <summary>Returns the document's text lines, or null when it uses a construct this pass does not read.</summary>
    /// <exception cref="InvalidDataException">The package is not a readable Open XML package.</exception>
    /// <exception cref="System.Xml.XmlException">A part is not well-formed.</exception>
    /// <param name="stream">The seekable package, read from its current position.</param>
    /// <param name="maxChars">The longest line the caller accepts; a longer one declines.</param>
    /// <param name="maxTableRows">Data rows written per table; 0 writes them all.</param>
    /// <param name="maxCharactersInPart">The character limit of the model load; a larger part declines.</param>
    /// <param name="includeNotes">Writes the text of referenced footnotes and endnotes after the sections.</param>
    /// <param name="cancellationToken">Checked between blocks, rows and cells.</param>
    /// <param name="declineReason">Why the pass declined, for diagnostics; null when it returns lines.</param>
    internal static List<string>? TryReadLines(Stream stream, int maxChars, int maxTableRows, long? maxCharactersInPart,
        bool includeNotes, CancellationToken cancellationToken, out string? declineReason) {
        declineReason = null;
        using OfficeOpcPackage package = OfficeOpcPackage.Open(stream, leaveOpen: true);
        string? main = package.GetRelatedPartName(null, OfficeOpcPackage.OfficeDocumentRelationshipType);
        if (main == null) {
            declineReason = "no main document part";
            return null;
        }
        package.RequireContentType(main,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml",
            "application/vnd.ms-word.document.macroEnabled.main+xml",
            "application/vnd.ms-word.template.macroEnabledTemplate.main+xml");
        var context = new Context(package, main, maxChars, maxTableRows, maxCharactersInPart, cancellationToken);
        bool read = context.Read(main, reader => ReadBody(reader, context));
        if (read && includeNotes) read = WriteNotes(context);
        if (!read) {
            declineReason = context.DeclineReason ?? "no body";
            return null;
        }
        foreach (string line in context.Lines) {
            if (line.Length > maxChars) {
                declineReason = "line longer than MaxChars";
                return null;
            }
        }
        return context.Lines;
    }

    private sealed class Context {
        private readonly Dictionary<string, PartLines> _headerFooters = new(StringComparer.Ordinal);
        private readonly HashSet<string> _noteReferences = new(StringComparer.Ordinal);

        internal Context(OfficeOpcPackage package, string mainPart, int maxChars, int maxTableRows, long? maxCharactersInPart, CancellationToken cancellationToken) {
            Package = package;
            MainPart = mainPart;
            MaxChars = maxChars;
            MaxTableRows = maxTableRows;
            MaxCharactersInPart = maxCharactersInPart;
            CancellationToken = cancellationToken;
        }

        internal OfficeOpcPackage Package { get; }
        internal string MainPart { get; }
        internal int MaxChars { get; }
        internal int MaxTableRows { get; }
        internal long? MaxCharactersInPart { get; }
        internal CancellationToken CancellationToken { get; }
        internal List<string> Lines { get; private set; } = new();
        internal OfficeCharBuffer Paragraph { get; } = new();

        /// <summary>Field state of the document story, which spans body paragraphs and table cells.</summary>
        internal ParagraphState BodyFields { get; } = new();

        /// <summary>Note references in the order the projection meets them: body, then that section's headers and footers.</summary>
        internal List<(bool Endnote, string Id)> NoteReferences { get; private set; } = new();

        // The even header and footer carry over from earlier sections; default and first do not.
        internal string? EvenHeader { get; set; }
        internal string? EvenFooter { get; set; }

        /// <summary>The first construct the pass could not read.</summary>
        internal string? DeclineReason { get; private set; }

        internal bool Decline(string reason) {
            DeclineReason ??= reason;
            return false;
        }

        internal void AddNoteReference(bool endnote, string id) => NoteReferences.Add((endnote, id));

        internal bool FirstNoteReference(bool endnote, string id) => _noteReferences.Add((endnote ? "endnote:" : "footnote:") + id);

        /// <summary>Reads a part of the package through <paramref name="read"/>.</summary>
        internal bool Read(string partName, Func<OfficeXmlPullReader, bool> read) {
            // A part whose declared size alone puts it over the limit declines before anything is inflated.
            if (MaxCharactersInPart > 0 && Package.GetPartLength(partName) > OfficeXmlPullReader.MaximumBytesFor(MaxCharactersInPart.Value)) {
                return Decline("part over the character limit");
            }
            OfficeXmlPullReader reader = Package.OpenXml(partName, out OfficePooledBytes content, processMarkupCompatibility: false);
            using (content) {
                // The model load caps a part's characters; a part that may exceed the cap declines and the model reports it.
                if (MaxCharactersInPart > 0 && content.Length > MaxCharactersInPart
                    && OfficeXmlPullReader.CountCharacters(content.Array, content.Length) > MaxCharactersInPart) {
                    return Decline("part over the character limit");
                }
                return read(reader);
            }
        }

        /// <summary>Writes the lines of a header or footer part, reading each part once however many sections use it.</summary>
        internal bool WriteHeaderFooter(string relationshipId, string relationshipType) {
            OfficeOpcRelationship? relationship = Package.FindRelationship(MainPart, relationshipId);
            // The model keeps a reference to a missing part as an empty header.
            if (relationship == null || relationship.TargetPartName == null || !Package.PartExists(relationship.TargetPartName)) return true;
            if (!string.Equals(relationship.Type, relationshipType, StringComparison.Ordinal)) return Decline("header or footer reference to another part type");
            string partName = relationship.TargetPartName;
            if (!_headerFooters.TryGetValue(partName, out PartLines? part)) {
                List<string> sectionLines = Lines;
                List<(bool, string)> sectionNotes = NoteReferences;
                Lines = new List<string>();
                NoteReferences = new List<(bool, string)>();
                bool read = Read(partName, reader => ReadHeaderFooter(reader, this));
                part = new PartLines(Lines, NoteReferences);
                Lines = sectionLines;
                NoteReferences = sectionNotes;
                if (!read) return false;
                _headerFooters.Add(partName, part);
            }
            Lines.AddRange(part.Lines);
            NoteReferences.AddRange(part.NoteReferences);
            return true;
        }
    }

    private sealed class PartLines {
        internal PartLines(List<string> lines, List<(bool, string)> noteReferences) {
            Lines = lines;
            NoteReferences = noteReferences;
        }

        internal List<string> Lines { get; }
        internal List<(bool Endnote, string Id)> NoteReferences { get; }
    }

    /// <summary>The header and footer references of one section.</summary>
    private sealed class SectionReferences {
        internal string? DefaultHeader, FirstHeader, EvenHeader, DefaultFooter, FirstFooter, EvenFooter;
    }

    private static bool ReadBody(OfficeXmlPullReader reader, Context context) {
        while (reader.Read()) {
            if (reader.NodeKind == OfficeXmlNodeKind.StartElement && reader.NamespaceId == W && reader.LocalName == "body") {
                if (!ReadBlocks(reader, context)) return false;
                // Reading to the end lets the parser check the rest of the part, such as a second root element.
                while (reader.Read()) { }
                return true;
            }
        }
        return context.Decline("no body");
    }

    private static bool ReadBlocks(OfficeXmlPullReader reader, Context context) {
        if (reader.IsEmptyElement) return true;
        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeKind == OfficeXmlNodeKind.EndElement && reader.Depth == depth) return true;
            if (reader.NodeKind != OfficeXmlNodeKind.StartElement) continue;
            if (reader.NamespaceId != W) return context.Decline(Outside(reader));
            context.CancellationToken.ThrowIfCancellationRequested();
            switch (reader.LocalName) {
                case "p":
                    if (!ReadParagraph(reader, context, context.BodyFields, recordNotes: true, canDrop: true, out string? text, out SectionReferences? section)) return false;
                    if (text != null) context.Lines.Add(text);
                    // A paragraph whose properties end a section closes it: its headers and footers follow its text.
                    if (section != null && !WriteSection(section, context)) return false;
                    break;
                case "tbl":
                    if (!ReadTable(reader, context, context.BodyFields, recordNotes: true)) return false;
                    break;
                case "sectPr":
                    if (!ReadSectionProperties(reader, context, out SectionReferences final)) return false;
                    if (!WriteSection(final, context)) return false;
                    break;
                case "sdt": case "altChunk": case "customXml":
                    // The inspection snapshot keeps only paragraphs and tables of a section.
                    if (!SkipBlockWithoutFields(reader, context)) return false;
                    break;
                case "bookmarkStart": case "bookmarkEnd": case "proofErr": case "permStart": case "permEnd":
                case "commentRangeStart": case "commentRangeEnd":
                case "moveFromRangeStart": case "moveFromRangeEnd": case "moveToRangeStart": case "moveToRangeEnd":
                    reader.SkipElement();
                    break;
                default:
                    return context.Decline("body element w:" + reader.LocalName);
            }
        }
        return context.Decline("unclosed body");
    }

    // The projection's order for a section's headers and footers.
    private static bool WriteSection(SectionReferences section, Context context) {
        if (section.EvenHeader != null) context.EvenHeader = section.EvenHeader;
        if (section.EvenFooter != null) context.EvenFooter = section.EvenFooter;
        return Write(section.DefaultHeader, HeaderRelationshipType)
            && Write(section.FirstHeader, HeaderRelationshipType)
            && Write(context.EvenHeader, HeaderRelationshipType)
            && Write(section.DefaultFooter, FooterRelationshipType)
            && Write(section.FirstFooter, FooterRelationshipType)
            && Write(context.EvenFooter, FooterRelationshipType);

        bool Write(string? relationshipId, string type) => relationshipId == null || context.WriteHeaderFooter(relationshipId, type);
    }

    private static bool ReadSectionProperties(OfficeXmlPullReader reader, Context context, out SectionReferences section) {
        section = new SectionReferences();
        if (reader.IsEmptyElement) return true;
        int depth = reader.Depth;
        while (reader.Read() && !(reader.NodeKind == OfficeXmlNodeKind.EndElement && reader.Depth == depth)) {
            if (reader.NodeKind != OfficeXmlNodeKind.StartElement || reader.NamespaceId != W || reader.Depth != depth + 1) continue;
            bool header = reader.LocalName == "headerReference";
            if (!header && reader.LocalName != "footerReference") continue;
            if (!reader.TryGetAttribute(OfficeXmlNamespaces.Relationships, "id", out string id)) continue;
            string type = reader.TryGetAttribute(W, "type", out string value) ? value : "default";
            switch (type) {
                case "default": if (header) section.DefaultHeader ??= id; else section.DefaultFooter ??= id; break;
                case "first": if (header) section.FirstHeader ??= id; else section.FirstFooter ??= id; break;
                case "even": if (header) section.EvenHeader ??= id; else section.EvenFooter ??= id; break;
                default: return context.Decline("header or footer reference type " + type);
            }
        }
        return true;
    }

    // A header or footer contributes its direct paragraphs and tables; content controls and other markup do not.
    private static bool ReadHeaderFooter(OfficeXmlPullReader reader, Context context) {
        var fields = new ParagraphState();
        while (reader.Read()) {
            if (reader.NodeKind != OfficeXmlNodeKind.StartElement) continue;
            if (reader.Depth == 1) {
                if (reader.NamespaceId != W || reader.LocalName is not ("hdr" or "ftr")) return context.Decline("header part without w:hdr or w:ftr");
                continue;
            }
            if (reader.Depth != 2) continue;
            if (reader.NamespaceId == W && reader.LocalName == "p") {
                // Header and footer paragraphs are kept even when they hold no text.
                if (!ReadParagraph(reader, context, fields, recordNotes: true, canDrop: false, out string? text, out SectionReferences? section)) return false;
                if (section != null) return context.Decline("section properties in a header or footer");
                context.Lines.Add(text!);
            } else if (reader.NamespaceId == W && reader.LocalName == "tbl") {
                if (!ReadTable(reader, context, fields, recordNotes: true)) return false;
            } else if (!SkipBlockWithoutFields(reader, context)) {
                return false;
            }
        }
        return true;
    }

    // Field state of a story: each open field is hidden while it reads its instruction and visible once it reaches its
    // result. A field may continue across paragraphs, so the state belongs to the story, not the paragraph.
    private sealed class ParagraphState {
        private readonly List<bool> _fields = new();
        // Open fields still in their instruction; text is visible only when none are.
        private int _instructions;

        internal bool RecordNotes { get; set; }
        internal bool Visible => _instructions == 0;
        internal bool Balanced => _fields.Count == 0;
        internal void Begin() {
            _fields.Add(false);
            _instructions++;
        }
        internal bool Separate() {
            if (_fields.Count == 0) return false;
            if (!_fields[_fields.Count - 1]) {
                _fields[_fields.Count - 1] = true;
                _instructions--;
            }
            return true;
        }
        internal bool End() {
            if (_fields.Count == 0) return false;
            if (!_fields[_fields.Count - 1]) _instructions--;
            _fields.RemoveAt(_fields.Count - 1);
            return true;
        }
    }

    // The model reads note ids as numbers, so "01" and "1" name the same note.
    private static string NoteId(string id) =>
        long.TryParse(id, System.Globalization.NumberStyles.Integer, CultureInfo.InvariantCulture, out long value)
            ? value.ToString(CultureInfo.InvariantCulture)
            : id;

    // Text is null when the model drops the paragraph: a paragraph whose children give the model no part to keep, such
    // as one whose only runs follow a field begin that does not end in the same paragraph.
    private static bool ReadParagraph(OfficeXmlPullReader reader, Context context, ParagraphState fields, bool recordNotes, bool canDrop,
        out string? text, out SectionReferences? section) {
        text = null;
        section = null;
        OfficeCharBuffer builder = context.Paragraph;
        builder.Clear();
        fields.RecordNotes = recordNotes;
        int noteMark = context.NoteReferences.Count;
        int children = 0;
        string? first = null, second = null;
        bool hasPart = false, buffering = false;
        int hyperlinkDepth = -1;
        if (!reader.IsEmptyElement) {
            int depth = reader.Depth;
            bool closed = false;
            while (reader.Read()) {
                if (reader.NodeKind == OfficeXmlNodeKind.EndElement) {
                    if (reader.Depth == depth) {
                        closed = true;
                        break;
                    }
                    if (reader.Depth == hyperlinkDepth) hyperlinkDepth = -1;
                    continue;
                }
                if (reader.NodeKind != OfficeXmlNodeKind.StartElement) continue;
                if (reader.NamespaceId == OfficeXmlNamespaces.Math && reader.Depth == depth + 1
                    && (reader.LocalNameEquals("oMath") || reader.LocalNameEquals("oMathPara"))) {
                    // Paragraph-level math is no part of the paragraph's text, but it keeps the paragraph.
                    children++;
                    if (children == 1) first = "oMath";
                    else if (children == 2) second = "oMath";
                    hasPart = true;
                    reader.SkipElement();
                    continue;
                }
                if (reader.NamespaceId != W) return context.Decline(Outside(reader));
                string name = reader.LocalName;
                if (reader.Depth == depth + 1) {
                    children++;
                    if (children == 1) first = name;
                    else if (children == 2) second = name;
                }
                // Runs inside a hyperlink belong to the hyperlink's part.
                bool partsCount = hyperlinkDepth < 0;
                switch (name) {
                    case "pPr":
                        if (!ReadParagraphProperties(reader, context, out section)) return false;
                        break;
                    case "r":
                        if (!ReadRun(reader, context, builder, fields, out string? fieldChar)) return false;
                        if (partsCount) TrackRunPart(fieldChar, ref buffering, ref hasPart);
                        break;
                    case "hyperlink":
                        // Its runs are part of the paragraph's text; the loop reads them as it descends.
                        if (partsCount) hasPart = true;
                        if (partsCount && !reader.IsEmptyElement) hyperlinkDepth = reader.Depth;
                        break;
                    case "customXml":
                        break;
                    case "fldSimple":
                        // A simple field shows its result runs; references inside it do not collect notes.
                        if (partsCount) hasPart = true;
                        if (!ReadSimpleField(reader, context, builder, fields)) return false;
                        break;
                    case "sdt":
                        if (partsCount) hasPart = true;
                        if (!ReadInlineContentControl(reader, context, builder, fields)) return false;
                        break;
                    case "ins": case "moveTo":
                        // The model reads no text from them, but their runs decide whether it keeps the paragraph.
                        if (!ScanRevisionParts(reader, context, partsCount, ref buffering, ref hasPart)) return false;
                        break;
                    case "bookmarkStart":
                        if (partsCount) hasPart = true;
                        reader.SkipElement();
                        break;
                    case "del": case "moveFrom": case "smartTag":
                    case "bookmarkEnd": case "proofErr": case "permStart": case "permEnd":
                    case "commentRangeStart": case "commentRangeEnd":
                    case "moveFromRangeStart": case "moveFromRangeEnd": case "moveToRangeStart": case "moveToRangeEnd":
                        reader.SkipElement();
                        break;
                    default:
                        return context.Decline("paragraph element w:" + name);
                }
            }
            if (!closed) return context.Decline("unclosed paragraph");
        }
        // The model keeps an empty paragraph, a lone run or properties, and properties followed by one run as they are.
        bool kept = !canDrop || children == 0 || hasPart
            || children == 1 && (first == "r" || first == "pPr")
            || children == 2 && first == "pPr" && second == "r";
        if (!kept) {
            context.NoteReferences.RemoveRange(noteMark, context.NoteReferences.Count - noteMark);
            return true;
        }
        text = builder.ToString();
        return true;
    }

    // Runs after a field begin wait for its end in the same paragraph; runs still waiting when the paragraph ends give
    // the model no part. Only the first field character of a run counts, as in the model.
    private static void TrackRunPart(string? fieldChar, ref bool buffering, ref bool hasPart) {
        if (buffering) {
            if (fieldChar == "end") {
                buffering = false;
                hasPart = true;
            }
        } else if (fieldChar == "begin") {
            buffering = true;
        } else {
            hasPart = true;
        }
    }

    private static bool ScanRevisionParts(OfficeXmlPullReader reader, Context context, bool partsCount, ref bool buffering, ref bool hasPart) {
        if (reader.IsEmptyElement) return true;
        int depth = reader.Depth;
        while (reader.Read() && !(reader.NodeKind == OfficeXmlNodeKind.EndElement && reader.Depth == depth)) {
            if (reader.NodeKind != OfficeXmlNodeKind.StartElement) continue;
            if (reader.NamespaceId != W) return context.Decline(Outside(reader));
            switch (reader.LocalName) {
                case "r":
                    string? fieldChar = FirstFieldChar(reader);
                    if (partsCount) TrackRunPart(fieldChar, ref buffering, ref hasPart);
                    break;
                case "hyperlink": case "fldSimple": case "sdt": case "bookmarkStart":
                    if (partsCount) hasPart = true;
                    reader.SkipElement();
                    break;
                case "customXml": case "ins": case "moveTo":
                    break;
                default:
                    reader.SkipElement();
                    break;
            }
        }
        return true;
    }

    // The model follows complex field markers through every paragraph of a story, those inside skipped block
    // content controls included. A skipped block that holds one would leave the field state behind, so it declines.
    private static bool SkipBlockWithoutFields(OfficeXmlPullReader reader, Context context) {
        if (reader.IsEmptyElement) return true;
        int depth = reader.Depth;
        while (reader.Read() && !(reader.NodeKind == OfficeXmlNodeKind.EndElement && reader.Depth == depth)) {
            if (reader.NodeKind == OfficeXmlNodeKind.StartElement && reader.NamespaceId == W && reader.LocalName == "fldChar") {
                return context.Decline("complex field marker in a skipped block");
            }
        }
        return true;
    }

    private static string? FirstFieldChar(OfficeXmlPullReader reader) {
        string? fieldChar = null;
        if (reader.IsEmptyElement) return null;
        int depth = reader.Depth;
        while (reader.Read() && !(reader.NodeKind == OfficeXmlNodeKind.EndElement && reader.Depth == depth)) {
            if (fieldChar == null && reader.NodeKind == OfficeXmlNodeKind.StartElement && reader.Depth == depth + 1
                && reader.NamespaceId == W && reader.LocalName == "fldChar") {
                fieldChar = reader.TryGetAttribute(W, "fldCharType", out string type) ? type : string.Empty;
            }
        }
        return fieldChar;
    }

    // List numbering never reaches the text; a section break inside a paragraph returns its references.
    private static bool ReadParagraphProperties(OfficeXmlPullReader reader, Context context, out SectionReferences? section) {
        section = null;
        if (reader.IsEmptyElement) return true;
        int depth = reader.Depth;
        while (reader.Read() && !(reader.NodeKind == OfficeXmlNodeKind.EndElement && reader.Depth == depth)) {
            if (reader.NodeKind == OfficeXmlNodeKind.StartElement && reader.NamespaceId == W && reader.LocalName == "sectPr" && reader.Depth == depth + 1) {
                if (!ReadSectionProperties(reader, context, out SectionReferences references)) return false;
                section = references;
            }
        }
        return true;
    }

    private static bool ReadSimpleField(OfficeXmlPullReader reader, Context context, OfficeCharBuffer builder, ParagraphState state) {
        if (reader.IsEmptyElement) return true;
        var inner = new ParagraphState();
        int depth = reader.Depth;
        while (reader.Read() && !(reader.NodeKind == OfficeXmlNodeKind.EndElement && reader.Depth == depth)) {
            if (reader.NodeKind != OfficeXmlNodeKind.StartElement) continue;
            if (reader.NamespaceId != W) return context.Decline(Outside(reader));
            if (reader.LocalName == "r") {
                int start = builder.Length;
                if (!ReadRun(reader, context, builder, inner, out _)) return false;
                if (!state.Visible) builder.Truncate(start);
            } else if (reader.LocalName is not ("bookmarkStart" or "bookmarkEnd" or "proofErr")) {
                return context.Decline("simple field element w:" + reader.LocalName);
            } else {
                reader.SkipElement();
            }
        }
        return inner.Balanced || context.Decline("field inside a simple field");
    }

    // A content control that holds a field contributes its content's runs with the story's field state, inserted runs
    // left out; any other contributes the visible text of its whole content when the field state shows it, and a check
    // box contributes nothing. Which applies depends on the whole control, so a forked reader looks ahead first.
    private static bool ReadInlineContentControl(OfficeXmlPullReader reader, Context context, OfficeCharBuffer builder, ParagraphState state) {
        if (reader.IsEmptyElement) return true;
        bool hasField = false, hasNote = false, checkBox = false, inProperties = false;
        OfficeXmlPullReader ahead = reader.Fork();
        int depth = ahead.Depth;
        int checkBoxNamespace = ahead.Namespaces.GetId(Word2010Namespace);
        while (ahead.Read() && !(ahead.NodeKind == OfficeXmlNodeKind.EndElement && ahead.Depth == depth)) {
            if (ahead.NodeKind == OfficeXmlNodeKind.EndElement) {
                if (ahead.Depth == depth + 1) inProperties = false;
                continue;
            }
            if (ahead.NodeKind != OfficeXmlNodeKind.StartElement) continue;
            if (ahead.Depth == depth + 1) {
                inProperties = ahead.NamespaceId == W && ahead.LocalNameEquals("sdtPr") && !ahead.IsEmptyElement;
            } else if (ahead.NamespaceId == W && ahead.LocalName == "fldChar") {
                hasField = true;
            } else if (ahead.NamespaceId == W && ahead.LocalName is "footnoteReference" or "endnoteReference") {
                hasNote = true;
            } else if (inProperties && ahead.Depth == depth + 2 && ahead.NamespaceId == checkBoxNamespace && ahead.LocalNameEquals("checkbox")) {
                // Only this control's own properties make it a check box, as in the model.
                checkBox = true;
            }
        }
        if (!hasField) {
            if (!state.Visible || checkBox) {
                reader.SkipElement();
                return true;
            }
            return AppendVisibleChildren(reader, context, builder);
        }
        if (hasNote) return context.Decline("note reference in a content control with a field");
        return ReadContentControlRuns(reader, context, builder, state);
    }

    // Enumerates runs the way the model's run list does: through custom XML, hyperlinks, simple fields and nested
    // content controls, but not inserted runs.
    private static bool ReadContentControlRuns(OfficeXmlPullReader reader, Context context, OfficeCharBuffer builder, ParagraphState state) {
        if (reader.IsEmptyElement) return true;
        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeKind == OfficeXmlNodeKind.EndElement && reader.Depth == depth) return true;
            if (reader.NodeKind != OfficeXmlNodeKind.StartElement || reader.Depth != depth + 1) continue;
            if (reader.NamespaceId != W) {
                reader.SkipElement();
                continue;
            }
            switch (reader.LocalName) {
                case "sdtContent": case "customXml": case "hyperlink":
                    if (!ReadContentControlRuns(reader, context, builder, state)) return false;
                    break;
                case "r":
                    bool recordNotes = state.RecordNotes;
                    state.RecordNotes = false;
                    bool read = ReadRun(reader, context, builder, state, out _);
                    state.RecordNotes = recordNotes;
                    if (!read) return false;
                    break;
                case "fldSimple":
                    if (!state.Visible) return context.Decline("simple field inside a hidden field");
                    if (!ReadContentControlRuns(reader, context, builder, state)) return false;
                    break;
                case "sdt":
                    if (!ReadInlineContentControl(reader, context, builder, state)) return false;
                    break;
                default:
                    reader.SkipElement();
                    break;
            }
        }
        return context.Decline("unclosed content control");
    }

    // Matches WordParagraph.Text for a run: text honours xml:space, tabs and breaks become characters, field
    // instructions are hidden, and a hidden run that only carries a comment reference contributes nothing.
    private static bool ReadRun(OfficeXmlPullReader reader, Context context, OfficeCharBuffer run, ParagraphState state, out string? firstFieldChar) {
        firstFieldChar = null;
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
            if (reader.NamespaceId == OfficeXmlNamespaces.MarkupCompatibility && reader.LocalNameEquals("AlternateContent")) {
                // A drawing or text box in compatibility markup: the model reads the branch it can interpret.
                if (!state.Visible) reader.SkipElement();
                else if (!AppendVisible(reader, context, run)) return false;
                continue;
            }
            if (reader.NamespaceId != W) return context.Decline(Outside(reader));
            string name = reader.LocalName;
            if (name == "fldChar") {
                string type = reader.TryGetAttribute(W, "fldCharType", out string value) ? value : string.Empty;
                firstFieldChar ??= type;
                bool known = type switch {
                    "begin" => Begin(state),
                    "separate" => state.Separate(),
                    "end" => state.End(),
                    _ => false
                };
                if (!known) return context.Decline("unmatched field character " + type);
                reader.SkipElement();
                continue;
            }
            if (name == "rPr") {
                if (!reader.IsEmptyElement) {
                    int propertiesDepth = reader.Depth;
                    while (reader.Read() && !(reader.NodeKind == OfficeXmlNodeKind.EndElement && reader.Depth == propertiesDepth)) {
                        if (reader.NodeKind != OfficeXmlNodeKind.StartElement) continue;
                        // A compatibility branch could change whether the run is hidden.
                        if (reader.NamespaceId == OfficeXmlNamespaces.MarkupCompatibility) return context.Decline("markup compatibility in run properties");
                        if (reader.Depth == propertiesDepth + 1 && reader.NamespaceId == W && reader.LocalName == "vanish") hidden = IsOn(reader);
                    }
                }
                continue;
            }
            bool visible = state.Visible;
            switch (name) {
                case "t":
                    bool preserve = reader.TryGetAttribute(OfficeXmlNamespaces.Xml, "space", out string space) && space == "preserve";
                    if (!visible) {
                        reader.SkipElement();
                    } else if (!ReadText(reader, context, run, preserve)) {
                        return false;
                    }
                    break;
                case "tab": case "ptab":
                    if (visible) run.Append('\t');
                    reader.SkipElement();
                    break;
                case "cr":
                    if (visible) run.Append('\n');
                    reader.SkipElement();
                    break;
                case "noBreakHyphen":
                    if (visible) run.Append('\u2011');
                    reader.SkipElement();
                    break;
                case "softHyphen":
                    if (visible) run.Append('\u00ad');
                    reader.SkipElement();
                    break;
                case "br":
                    string? type = reader.TryGetAttribute(W, "type", out string typeValue) ? typeValue : null;
                    if (visible) run.Append(type is null or "textWrapping" ? '\n' : '\u2028');
                    reader.SkipElement();
                    break;
                case "footnoteReference": case "endnoteReference":
                    if (visible && state.RecordNotes && reader.TryGetAttribute(W, "id", out string noteId)) {
                        context.AddNoteReference(name == "endnoteReference", NoteId(noteId));
                    }
                    reader.SkipElement();
                    break;
                case "commentReference":
                    commentReference = true;
                    reader.SkipElement();
                    break;
                case "instrText": case "delInstrText": case "delText": case "sym": case "lastRenderedPageBreak":
                case "footnoteRef": case "endnoteRef": case "separator": case "continuationSeparator": case "annotationRef":
                case "dayShort": case "dayLong": case "monthShort": case "monthLong": case "yearShort": case "yearLong":
                case "pgNum":
                    reader.SkipElement();
                    break;
                case "drawing": case "pict": case "object":
                    // Text boxes contribute their text inline, with no separator, as the model reads them.
                    if (!visible) reader.SkipElement();
                    else if (!AppendVisible(reader, context, run)) return false;
                    break;
                default:
                    return context.Decline("run element w:" + name);
            }
        }
        return context.Decline("unclosed run");

        static bool Begin(ParagraphState state) {
            state.Begin();
            return true;
        }
    }

    // Choice branches the model interprets: those requiring only these Office extension namespaces.
    private static readonly HashSet<string> SupportedChoiceNamespaces = new(StringComparer.Ordinal) {
        "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
        "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
        "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
        "http://schemas.microsoft.com/office/word/2010/wordprocessingInk",
        "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
        "http://schemas.microsoft.com/office/word/2010/wordml",
        "http://schemas.microsoft.com/office/word/2012/wordml",
        "http://schemas.microsoft.com/office/word/2015/wordml/symex",
        "http://schemas.microsoft.com/office/word/2016/wordml/cid",
        "http://schemas.microsoft.com/office/word/2018/wordml",
        "http://schemas.microsoft.com/office/word/2018/wordml/cex",
        "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash"
    };

    // Prefixes the model accepts undeclared, as older documents omitted the declaration in header and footer parts.
    private static readonly Dictionary<string, string> SupportedChoicePrefixes = new(StringComparer.Ordinal) {
        ["wpc"] = "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
        ["wp14"] = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
        ["wpg"] = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
        ["wpi"] = "http://schemas.microsoft.com/office/word/2010/wordprocessingInk",
        ["wps"] = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
        ["w14"] = "http://schemas.microsoft.com/office/word/2010/wordml",
        ["w15"] = "http://schemas.microsoft.com/office/word/2012/wordml",
        ["w16se"] = "http://schemas.microsoft.com/office/word/2015/wordml/symex",
        ["w16cid"] = "http://schemas.microsoft.com/office/word/2016/wordml/cid",
        ["w16"] = "http://schemas.microsoft.com/office/word/2018/wordml",
        ["w16cex"] = "http://schemas.microsoft.com/office/word/2018/wordml/cex",
        ["w16sdtdh"] = "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash"
    };

    private const string Word2010Namespace = "http://schemas.microsoft.com/office/word/2010/wordml";
    private const string WordprocessingGroupNamespace = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup";

    // Appends the visible text of an element and its descendants the way the model reads drawings, text boxes and
    // content controls: text, tabs and breaks of every nested run in document order, with no separators; field
    // characters and instructions, comment references, deletions, moves from and grouped shapes add nothing; markup
    // compatibility contributes its selected branch.
    private static bool AppendVisible(OfficeXmlPullReader reader, Context context, OfficeCharBuffer builder) {
        int namespaceId = reader.NamespaceId;
        if (namespaceId == W) {
            switch (reader.LocalName) {
                case "t":
                    bool preserve = reader.TryGetAttribute(OfficeXmlNamespaces.Xml, "space", out string space) && space == "preserve";
                    return ReadText(reader, context, builder, preserve);
                case "tab": case "ptab":
                    builder.Append('\t');
                    reader.SkipElement();
                    return true;
                case "cr":
                    builder.Append('\n');
                    reader.SkipElement();
                    return true;
                case "noBreakHyphen":
                    builder.Append('\u2011');
                    reader.SkipElement();
                    return true;
                case "softHyphen":
                    builder.Append('\u00ad');
                    reader.SkipElement();
                    return true;
                case "br":
                    string? type = reader.TryGetAttribute(W, "type", out string typeValue) ? typeValue : null;
                    builder.Append(type is null or "textWrapping" ? '\n' : '\u2028');
                    reader.SkipElement();
                    return true;
                case "commentReference": case "fldChar": case "instrText": case "del": case "moveFrom":
                    reader.SkipElement();
                    return true;
                case "r":
                    return AppendVisibleRun(reader, context, builder);
            }
        } else if (namespaceId == OfficeXmlNamespaces.Math) {
            return context.Decline("math in a drawing or content control");
        } else if (namespaceId == OfficeXmlNamespaces.MarkupCompatibility && reader.LocalNameEquals("AlternateContent")) {
            return AppendAlternateContent(reader, context, builder);
        } else if (namespaceId == OfficeXmlNamespaces.Vml && reader.LocalNameEquals("group")
            || (reader.LocalNameEquals("wgp") || reader.LocalNameEquals("grpSp"))
            && reader.Namespaces.GetUri(namespaceId) == WordprocessingGroupNamespace) {
            // A grouped drawing owns its labels; they are not part of the enclosing paragraph's text.
            reader.SkipElement();
            return true;
        }
        return AppendVisibleChildren(reader, context, builder);
    }

    private static bool AppendVisibleChildren(OfficeXmlPullReader reader, Context context, OfficeCharBuffer builder) {
        if (reader.IsEmptyElement) return true;
        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeKind == OfficeXmlNodeKind.EndElement && reader.Depth == depth) return true;
            if (reader.NodeKind == OfficeXmlNodeKind.StartElement && reader.Depth == depth + 1 && !AppendVisible(reader, context, builder)) return false;
        }
        return context.Decline("unclosed element");
    }

    // A nested run hidden with vanish that carries a comment reference adds nothing, as in the host paragraph.
    private static bool AppendVisibleRun(OfficeXmlPullReader reader, Context context, OfficeCharBuffer builder) {
        if (reader.IsEmptyElement) return true;
        int start = builder.Length;
        bool hidden = false, commentReference = false;
        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeKind == OfficeXmlNodeKind.EndElement && reader.Depth == depth) {
                if (hidden && commentReference) builder.Truncate(start);
                return true;
            }
            if (reader.NodeKind != OfficeXmlNodeKind.StartElement || reader.Depth != depth + 1) continue;
            if (reader.NamespaceId == W && reader.LocalName == "rPr") {
                if (!reader.IsEmptyElement) {
                    int propertiesDepth = reader.Depth;
                    while (reader.Read() && !(reader.NodeKind == OfficeXmlNodeKind.EndElement && reader.Depth == propertiesDepth)) {
                        if (reader.NodeKind != OfficeXmlNodeKind.StartElement) continue;
                        if (reader.NamespaceId == OfficeXmlNamespaces.MarkupCompatibility) return context.Decline("markup compatibility in run properties");
                        if (reader.Depth == propertiesDepth + 1 && reader.NamespaceId == W && reader.LocalName == "vanish") hidden = IsOn(reader);
                    }
                }
                continue;
            }
            if (reader.NamespaceId == W && reader.LocalName == "commentReference") commentReference = true;
            if (!AppendVisible(reader, context, builder)) return false;
        }
        return context.Decline("unclosed run");
    }

    // The first Choice whose required namespaces are all supported, else the Fallback.
    private static bool AppendAlternateContent(OfficeXmlPullReader reader, Context context, OfficeCharBuffer builder) {
        if (reader.IsEmptyElement) return true;
        bool selected = false, fallbackSeen = false;
        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeKind == OfficeXmlNodeKind.EndElement && reader.Depth == depth) return true;
            if (reader.NodeKind != OfficeXmlNodeKind.StartElement || reader.Depth != depth + 1) continue;
            bool choice = reader.NamespaceId == OfficeXmlNamespaces.MarkupCompatibility && reader.LocalNameEquals("Choice");
            bool fallback = reader.NamespaceId == OfficeXmlNamespaces.MarkupCompatibility && reader.LocalNameEquals("Fallback");
            if (choice && fallbackSeen) return context.Decline("compatibility choice after its fallback");
            fallbackSeen |= fallback;
            if (!selected && (fallback || choice && RequirementsSupported(reader))) {
                selected = true;
                if (!AppendVisibleChildren(reader, context, builder)) return false;
            } else {
                reader.SkipElement();
            }
        }
        return context.Decline("unclosed compatibility markup");
    }

    private static bool RequirementsSupported(OfficeXmlPullReader reader) {
        if (!reader.TryGetAttribute(OfficeXmlNamespaces.None, "Requires", out string requires)) return false;
        string[] prefixes = requires.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
        if (prefixes.Length == 0) return false;
        foreach (string prefix in prefixes) {
            string? uri = reader.LookupNamespace(prefix);
            if (uri == null) SupportedChoicePrefixes.TryGetValue(prefix, out uri);
            if (uri == null || !SupportedChoiceNamespaces.Contains(uri)) return false;
        }
        return true;
    }

    // Appends the element's text; without xml:space="preserve" the XML white space at either edge is dropped.
    private static bool ReadText(OfficeXmlPullReader reader, Context context, OfficeCharBuffer builder, bool preserve) {
        if (reader.IsEmptyElement) return true;
        int start = builder.Length;
        int depth = reader.Depth;
        while (reader.Read() && !(reader.NodeKind == OfficeXmlNodeKind.EndElement && reader.Depth == depth)) {
            if (reader.NodeKind == OfficeXmlNodeKind.Text) {
                // The model reads a text element's value up to the first comment or processing instruction.
                if (!reader.FollowsCommentOrInstruction) reader.AppendText(builder);
            } else if (reader.NodeKind == OfficeXmlNodeKind.StartElement) {
                return context.Decline("element inside w:t");
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
    // tab-separated cells; a cell is its non-blank paragraphs joined by spaces. Cells are read as written: the model
    // ignores grid spans and keeps vertically merged cells as their own, usually empty, cells.
    private static bool ReadTable(OfficeXmlPullReader reader, Context context, ParagraphState fields, bool recordNotes) {
        if (reader.IsEmptyElement) return true;
        int depth = reader.Depth;
        bool repeatsHeader = false;
        var rows = new List<List<string>>();
        while (reader.Read()) {
            if (reader.NodeKind == OfficeXmlNodeKind.EndElement && reader.Depth == depth) break;
            if (reader.NodeKind != OfficeXmlNodeKind.StartElement) continue;
            if (reader.NamespaceId != W) return context.Decline(Outside(reader));
            context.CancellationToken.ThrowIfCancellationRequested();
            switch (reader.LocalName) {
                case "tblPr":
                case "tblGrid":
                    // The model sizes the table from its widest row, not its grid.
                    reader.SkipElement();
                    break;
                case "tr":
                    var cells = new List<string>();
                    if (!ReadRow(reader, context, fields, recordNotes, cells, out bool header)) return false;
                    if (rows.Count == 0) repeatsHeader = header;
                    rows.Add(cells);
                    break;
                case "bookmarkStart": case "bookmarkEnd": case "proofErr":
                    reader.SkipElement();
                    break;
                default:
                    return context.Decline("table element w:" + reader.LocalName);
            }
        }
        int columns = 0;
        foreach (List<string> row in rows) columns = Math.Max(columns, row.Count);
        // Every line carries the separators of the widest row, so a table wider than the line limit cannot stream.
        if (columns - 1 > context.MaxChars) return context.Decline("line longer than MaxChars");
        bool hasHeader = repeatsHeader && rows.Count > 0;
        if (columns > 0) {
            var names = new string[columns];
            for (int index = 0; index < columns; index++) {
                string value = hasHeader ? Cell(rows[0], index) : string.Empty;
                names[index] = string.IsNullOrWhiteSpace(value) ? "Column " + (index + 1).ToString(CultureInfo.InvariantCulture) : value;
            }
            string line = string.Join("\t", names);
            if (line.Length > context.MaxChars) return context.Decline("line longer than MaxChars");
            context.Lines.Add(line);
        }
        int written = 0;
        for (int rowIndex = hasHeader ? 1 : 0; rowIndex < rows.Count; rowIndex++) {
            if (context.MaxTableRows > 0 && written == context.MaxTableRows) break;
            var values = new string[columns];
            for (int index = 0; index < columns; index++) values[index] = Cell(rows[rowIndex], index);
            string line = string.Join("\t", values);
            if (line.Length > context.MaxChars) return context.Decline("line longer than MaxChars");
            context.Lines.Add(line);
            written++;
        }
        return true;
    }

    private static string Cell(List<string> row, int index) => index < row.Count ? row[index] : string.Empty;

    private static bool ReadRow(OfficeXmlPullReader reader, Context context, ParagraphState fields, bool recordNotes, List<string> cells, out bool repeatsHeader) {
        repeatsHeader = false;
        if (reader.IsEmptyElement) return true;
        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeKind == OfficeXmlNodeKind.EndElement && reader.Depth == depth) return true;
            if (reader.NodeKind != OfficeXmlNodeKind.StartElement) continue;
            if (reader.NamespaceId != W) return context.Decline(Outside(reader));
            switch (reader.LocalName) {
                case "trPr": case "tblPrEx":
                    if (!reader.IsEmptyElement) {
                        bool rowProperties = reader.LocalName == "trPr";
                        bool headerSeen = false;
                        int propertiesDepth = reader.Depth;
                        while (reader.Read() && !(reader.NodeKind == OfficeXmlNodeKind.EndElement && reader.Depth == propertiesDepth)) {
                            if (reader.NodeKind != OfficeXmlNodeKind.StartElement || reader.NamespaceId != W) continue;
                            // Revised row properties are a revision; the model read handles those.
                            if (reader.LocalName == "trPrChange") return context.Decline("row property w:" + reader.LocalName);
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
                    if (!ReadCell(reader, context, fields, recordNotes, out string? value)) return false;
                    cells.Add(value!);
                    break;
                case "bookmarkStart": case "bookmarkEnd": case "proofErr":
                    reader.SkipElement();
                    break;
                default:
                    return context.Decline("row element w:" + reader.LocalName);
            }
        }
        return context.Decline("unclosed row");
    }

    // A cell's text is its direct paragraphs; nested tables and paragraphs inside content controls are not read.
    private static bool ReadCell(OfficeXmlPullReader reader, Context context, ParagraphState fields, bool recordNotes, out string? value) {
        value = null;
        var parts = new List<string>();
        if (!reader.IsEmptyElement) {
            int depth = reader.Depth;
            bool closed = false;
            while (reader.Read()) {
                if (reader.NodeKind == OfficeXmlNodeKind.EndElement && reader.Depth == depth) {
                    closed = true;
                    break;
                }
                if (reader.NodeKind != OfficeXmlNodeKind.StartElement) continue;
                if (reader.NamespaceId != W) return context.Decline(Outside(reader));
                switch (reader.LocalName) {
                    case "tcPr": case "tbl": case "sdt": case "customXml":
                        reader.SkipElement();
                        break;
                    case "p":
                        if (!ReadParagraph(reader, context, fields, recordNotes, canDrop: true, out string? text, out SectionReferences? section)) return false;
                        if (section != null) return context.Decline("section properties in a table cell");
                        if (!string.IsNullOrWhiteSpace(text)) parts.Add(text!);
                        break;
                    case "bookmarkStart": case "bookmarkEnd": case "proofErr":
                        reader.SkipElement();
                        break;
                    default:
                        return context.Decline("cell element w:" + reader.LocalName);
                }
            }
            if (!closed) return context.Decline("unclosed cell");
        }
        value = string.Join(" ", parts);
        return true;
    }

    // Each referenced note, once, in reference order: its non-blank paragraphs joined by line breaks.
    private static bool WriteNotes(Context context) {
        if (context.NoteReferences.Count == 0) return true;
        Dictionary<string, string>? footnotes = null, endnotes = null;
        foreach ((bool endnote, string id) in context.NoteReferences.ToArray()) {
            Dictionary<string, string>? notes = endnote ? endnotes : footnotes;
            if (notes == null) {
                if (!ReadNotes(context, endnote, out notes)) return false;
                if (endnote) endnotes = notes; else footnotes = notes;
            }
            if (!notes!.TryGetValue(id, out string? text) || string.IsNullOrWhiteSpace(text)) continue;
            if (context.FirstNoteReference(endnote, id)) context.Lines.Add(text);
        }
        return true;
    }

    private static bool ReadNotes(Context context, bool endnotes, out Dictionary<string, string> notes) {
        var read = new Dictionary<string, string>(StringComparer.Ordinal);
        notes = read;
        string? partName = context.Package.GetRelatedPartName(context.MainPart, endnotes ? EndnotesRelationshipType : FootnotesRelationshipType);
        if (partName == null || !context.Package.PartExists(partName)) return true;
        string noteElement = endnotes ? "endnote" : "footnote";
        bool result = context.Read(partName, reader => {
            var paragraphs = new List<string>();
            while (reader.Read()) {
                if (reader.NodeKind != OfficeXmlNodeKind.StartElement || reader.Depth != 2) continue;
                if (reader.NamespaceId != W || reader.LocalName != noteElement) {
                    reader.SkipElement();
                    continue;
                }
                // Separator notes are not note text.
                if (reader.TryGetAttribute(W, "type", out string type) && type != "normal") {
                    reader.SkipElement();
                    continue;
                }
                string id = reader.TryGetAttribute(W, "id", out string value) ? NoteId(value) : string.Empty;
                paragraphs.Clear();
                // Each note is a story of its own.
                var fields = new ParagraphState();
                if (!reader.IsEmptyElement) {
                    int depth = reader.Depth;
                    while (reader.Read() && !(reader.NodeKind == OfficeXmlNodeKind.EndElement && reader.Depth == depth)) {
                        if (reader.NodeKind != OfficeXmlNodeKind.StartElement) continue;
                        if (reader.Depth == depth + 1 && reader.NamespaceId == W && reader.LocalName == "p") {
                            int references = context.NoteReferences.Count;
                            if (!ReadParagraph(reader, context, fields, recordNotes: true, canDrop: true, out string? text, out SectionReferences? section)) return false;
                            if (section != null || context.NoteReferences.Count != references) return context.Decline("note with a section or a note reference");
                            if (!string.IsNullOrWhiteSpace(text)) paragraphs.Add(text!);
                        } else {
                            reader.SkipElement();
                        }
                    }
                }
                if (id.Length > 0 && !read.ContainsKey(id)) read.Add(id, string.Join(Environment.NewLine, paragraphs));
            }
            return true;
        });
        return result;
    }

    private static string Outside(OfficeXmlPullReader reader) =>
        "element outside WordprocessingML: " + reader.Namespaces.GetUri(reader.NamespaceId) + " " + reader.LocalName;

    private static bool IsOn(OfficeXmlPullReader reader) =>
        !reader.TryGetAttribute(W, "val", out string value) || value is "1" or "true" or "on";
}
