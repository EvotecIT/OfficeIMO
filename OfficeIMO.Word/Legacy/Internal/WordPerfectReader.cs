using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;

namespace OfficeIMO.Word.Legacy;

/// <summary>Bounded, inert WordPerfect document decoding. Each text story has independent formatting state.</summary>
internal sealed partial class WordPerfectReader {
    private readonly byte[] _data;
    private readonly LegacyWordModel _model;
    private readonly Budget _budget;
    private readonly Dictionary<int, Packet> _packets;
    private readonly HashSet<int> _activePackets;
    private readonly int _version;
    private readonly int _depth;
    private readonly List<LegacyWordParagraph>? _story;
    private LegacyWordSection _section;
    private LegacyWordParagraph? _paragraph;
    private LegacyWordTable? _table;
    private LegacyWordTableRow? _row;
    private LegacyWordTableCell? _cell;
    private uint _attributes;
    private string? _font, _color;
    private double? _fontSize;
    private WordParagraphAlignment? _alignment;
    private double? _spacingAfter;
    private bool _pageBreak;
    private int _encasedNote = -1;

    private WordPerfectReader(byte[] data, int version, OfficeLegacyImportLimits limits, CancellationToken cancellation) {
        _data = data; _version = version;
        _model = new LegacyWordModel { Quality = OfficeLegacyImportQuality.Structured };
        _budget = new Budget(limits, cancellation);
        _packets = new Dictionary<int, Packet>();
        _activePackets = new HashSet<int>();
        _section = new LegacyWordSection();
        _model.Sections.Add(_section);
        _budget.Item();
    }

    private WordPerfectReader(WordPerfectReader parent, List<LegacyWordParagraph> story) {
        _data = parent._data; _version = parent._version; _model = parent._model;
        _budget = parent._budget; _packets = parent._packets; _activePackets = parent._activePackets;
        _depth = parent._depth + 1; _story = story; _section = new LegacyWordSection();
        _font = parent._font; _fontSize = parent._fontSize;
        if (_depth > 32) throw new InvalidDataException("WordPerfect text-story nesting exceeds the supported depth.");
    }

    internal static bool IsStructuredProfile(byte[] data) =>
        data.Length >= 16 && OfficeLegacyImportBuffer.StartsWith(data, 0xff, 0x57, 0x50, 0x43) &&
        data[9] == 10 && (data[10] == 0 || data[10] == 2);

    internal static LegacyWordModel Read(byte[] data, OfficeLegacyImportLimits limits, CancellationToken cancellation) {
        if (!IsStructuredProfile(data)) throw new InvalidDataException("The input is not a WordPerfect 5/6 document.");
        if (U16(data, 12) != 0) throw new InvalidDataException("Encrypted WordPerfect documents cannot be imported.");
        int offset = I32(data, 4);
        if (offset < 16 || offset > data.Length) throw new InvalidDataException("The WordPerfect document offset is outside the file.");
        var reader = new WordPerfectReader(data, data[10] == 0 ? 5 : 6, limits, cancellation);
        reader._model.Metadata["WordPerfectVersion"] = reader._version.ToString(CultureInfo.InvariantCulture);
        reader._model.Metadata["DocumentAreaOffset"] = offset.ToString(CultureInfo.InvariantCulture);
        if (reader._version == 6) reader.ReadPrefix6(offset);
        else reader.ReadPrefix5(offset);
        int end = data.Length;
        if (reader._version == 6 && U16(data, 14) >= 24) {
            int declared = I32(data, 20);
            if (declared >= offset && declared <= data.Length) end = declared;
            else reader.Loss("WORDPERFECT_FILE_SIZE", "Structure", "The stored file size is inconsistent; records were bounded by the actual input length.");
        }
        reader.ReadText(offset, end);
        reader.FinishParagraph();
        if (reader._table != null) throw new InvalidDataException("The WordPerfect document ends inside a table.");
        if (reader._encasedNote >= 0) throw new InvalidDataException("The WordPerfect document ends inside a note reference.");
        if (reader._invalidUndo.Count != 0) throw new InvalidDataException("The WordPerfect document ends inside deleted undo text.");
        return reader._model;
    }

    private void ReadText(int start, int end) {
        CheckRange(_data, start, end - start);
        for (int position = start; position < end;) {
            _budget.Record();
            byte code = _data[position];
            int asciiStart = _version == 6 ? 0x21 : 0x20;
            if (code >= asciiStart && code <= 0x7e) {
                int next = position + 1;
                while (next < end && _data[next] >= asciiStart && _data[next] <= 0x7e) {
                    if ((next & 4095) == 0) _budget.Cancellation.ThrowIfCancellationRequested();
                    next++;
                }
                if (!Suppressed) {
                    _budget.Text(next - position);
                    AppendRun(Encoding.ASCII.GetString(_data, position, next - position));
                }
                position = next;
            } else if (_version == 6) position = ReadCode6(position, end);
            else position = ReadCode5(position, end);
        }
    }

    private LegacyWordParagraph Paragraph() {
        if (_paragraph == null) {
            _budget.Item();
            _paragraph = new LegacyWordParagraph {
                Alignment = _alignment, SpacingAfterPoints = _spacingAfter,
                PageBreakBefore = _pageBreak
            };
            _pageBreak = false;
        }
        return _paragraph;
    }

    private void Append(string text) {
        _budget.Text(text.Length);
        AppendRun(text);
    }

    private void AppendRun(string text) {
        var run = new LegacyWordRun(text) {
            Bold = (_attributes & (1u << 12)) != 0,
            Italic = (_attributes & (1u << 8)) != 0,
            Strike = (_attributes & (1u << 13)) != 0,
            Underline = (_attributes & (1u << 11)) != 0 ? WordUnderlineStyle.Double :
                (_attributes & (1u << 14)) != 0 ? WordUnderlineStyle.Single : (WordUnderlineStyle?)null,
            VerticalPosition = (_attributes & (1u << 5)) != 0 ? WordVerticalTextPosition.Superscript :
                (_attributes & (1u << 6)) != 0 ? WordVerticalTextPosition.Subscript : (WordVerticalTextPosition?)null,
            FontFamily = _font, FontSizePoints = _fontSize, ColorHex = _color
        };
        LegacyWordParagraph paragraph = Paragraph();
        LegacyWordRun? last = paragraph.Runs.LastOrDefault();
        if (last != null && last.NoteIndex == null && last.Image == null && last.Bold == run.Bold && last.Italic == run.Italic &&
            last.Strike == run.Strike && last.Underline == run.Underline && last.VerticalPosition == run.VerticalPosition &&
            last.FontFamily == run.FontFamily && last.FontSizePoints == run.FontSizePoints && last.ColorHex == run.ColorHex) last.AppendText(text);
        else { _budget.Item(); paragraph.Runs.Add(run); }
    }

    private void FinishParagraph(bool force = false) {
        if (_paragraph == null && !force) return;
        LegacyWordParagraph paragraph = Paragraph();
        _paragraph = null;
        if (_story != null) _story.Add(paragraph);
        else if (_table != null) {
            if (_cell == null) throw new InvalidDataException("WordPerfect table text appears outside a cell.");
            _cell.Paragraphs.Add(paragraph);
        } else {
            _section.Blocks.Add(paragraph);
            _model.Paragraphs.Add(paragraph);
        }
    }

    private void BreakPage() {
        FinishParagraph();
        _pageBreak = true;
    }

    private bool ChangeSection() {
        if (_story != null) { Loss("WORDPERFECT_STORY_LAYOUT", "Layout", "Page geometry inside a running story or note was not applied to the document."); return false; }
        FinishParagraph();
        if (_table != null) { Loss("WORDPERFECT_TABLE_LAYOUT", "Layout", "A page-setting change inside a table was not applied."); return false; }
        if (_section.Blocks.Count == 0) return true;
        _budget.Item();
        _section = _section.CopySettings();
        _section.StartsNewPage = _pageBreak;
        _model.Sections.Add(_section);
        _pageBreak = false;
        return true;
    }

    private void Attribute(byte attribute, bool enabled) {
        if ((attribute & 0x80) != 0) return; // Nested same-attribute codes explicitly marked ignored by the producer.
        int value = attribute & 0x3f;
        if (value >= 32) { Loss("WORDPERFECT_ATTRIBUTE", "Formatting", "An unknown character attribute was omitted."); return; }
        if (enabled) _attributes |= 1u << value;
        else _attributes &= ~(1u << value);
        if (value != 5 && value != 6 && value != 8 && value != 11 && value != 12 && value != 13 && value != 14)
            Loss("WORDPERFECT_ATTRIBUTE_" + value, "Formatting", "Character attribute " + value + " is outside the recovered formatting profile.");
    }

    private void Justification(int value) {
        _alignment = value switch {
            0 => WordParagraphAlignment.Left, 1 => WordParagraphAlignment.Both,
            2 => WordParagraphAlignment.Center, 3 => WordParagraphAlignment.Right,
            _ => (WordParagraphAlignment?)null
        };
        if (!_alignment.HasValue) Loss("WORDPERFECT_JUSTIFICATION", "Formatting", "An unsupported justification mode was omitted.");
        if (_paragraph != null) _paragraph.Alignment = _alignment;
    }

    private void BeginTable() {
        if (_story != null) throw new InvalidDataException("Tables inside WordPerfect running stories are outside the qualified profile.");
        if (_table != null) throw new InvalidDataException("Nested WordPerfect tables are outside the qualified profile.");
        FinishParagraph(); _budget.Item();
        _table = new LegacyWordTable(); _section.Blocks.Add(_table);
        Loss("WORDPERFECT_TABLE_APPEARANCE", "Table", "Table text, grid and column widths are recovered; source border, fill and advanced cell formatting are not fully projected.");
    }

    private void BeginRow(bool createCell = true) {
        if (_table == null) throw new InvalidDataException("WordPerfect row appears outside a table.");
        FinishParagraph(); _budget.Item();
        _row = new LegacyWordTableRow(); _table.Rows.Add(_row); _cell = null;
        if (createCell) BeginCell();
    }

    private void BeginCell() {
        if (_row == null) throw new InvalidDataException("WordPerfect cell appears outside a row.");
        FinishParagraph(); _budget.Item();
        _cell = new LegacyWordTableCell(); _row.Cells.Add(_cell);
    }

    private void EndTable() {
        if (_table == null) throw new InvalidDataException("WordPerfect table end appears outside a table.");
        FinishParagraph();
        if (_table.Rows.Count == 0) throw new InvalidDataException("WordPerfect table has no rows.");
        int columns = _table.ColumnWidthsPoints.Count;
        if (columns == 0) columns = _table.Rows.Max(row => row.Cells.Sum(cell => cell.ColumnSpan));
        foreach (LegacyWordTableRow row in _table.Rows) {
            if (row.Cells.Sum(cell => cell.ColumnSpan) != columns)
                throw new InvalidDataException("WordPerfect table rows do not match their declared grid.");
        }
        _table = null; _row = null; _cell = null;
    }

    private void AddNote(LegacyWordNoteKind kind, List<LegacyWordParagraph> paragraphs) {
        if (_story != null) throw new InvalidDataException("Notes inside WordPerfect running stories or notes are outside the qualified profile.");
        string text = string.Join("\n", paragraphs.Select(paragraph => paragraph.Text));
        _budget.Item();
        var note = new LegacyWordNote(kind, text) { IsAnchored = true };
        note.Paragraphs.AddRange(paragraphs);
        int index = _model.Notes.Count; _model.Notes.Add(note);
        _budget.Item();
        Paragraph().Runs.Add(new LegacyWordRun(string.Empty) { NoteIndex = index });
    }

    private void FontSize(double size) {
        _fontSize = size > 0 ? size : (double?)null;
        if (!_fontSize.HasValue) Loss("WORDPERFECT_FONT_SIZE", "Formatting", "A source font has no positive point size.");
    }

    private void SetStory(int slot, int occurrenceBits, List<LegacyWordParagraph> paragraphs) {
        if (slot > 3) { Loss("WORDPERFECT_WATERMARK", "Layout", "A watermark or fancy-border story was not projected."); return; }
        if (!ChangeSection()) return;
        _section.HeadersAndFooters.RemoveAll(story => story.Slot == slot);
        if (occurrenceBits == 0) return;
        _budget.Item();
        var story = new LegacyWordHeaderFooter {
            Slot = slot, IsFooter = slot >= 2,
            Occurrence = occurrenceBits == 1 ? LegacyWordHeaderFooterOccurrence.Odd :
                occurrenceBits == 2 ? LegacyWordHeaderFooterOccurrence.Even : LegacyWordHeaderFooterOccurrence.All
        };
        story.Paragraphs.AddRange(paragraphs); _section.HeadersAndFooters.Add(story);
        _section.HeadersAndFooters.Sort((left, right) => left.Slot.CompareTo(right.Slot));
    }

    private List<LegacyWordParagraph> ReadStory(int start, int end) {
        var paragraphs = new List<LegacyWordParagraph>();
        var reader = new WordPerfectReader(this, paragraphs);
        reader.ReadText(start, end); reader.FinishParagraph();
        if (reader._encasedNote >= 0) throw new InvalidDataException("WordPerfect story ends inside a note reference.");
        if (reader._invalidUndo.Count != 0) throw new InvalidDataException("WordPerfect story ends inside deleted undo text.");
        return paragraphs;
    }

    private void Loss(string code, string category, string message) {
        if (_budget.Reported.Add(code)) _model.Findings.Add(LegacyWordAdapterBase.LossFinding(code, category, message));
    }

    private void Inert(string code, OfficeLegacyInertContentKind kind, string message) {
        _model.InertContent |= kind;
        if (_budget.Reported.Add(code)) _model.Findings.Add(LegacyWordAdapterBase.InertFinding(code, "Security", message));
    }

    private static void CheckRange(byte[] data, int start, int length) {
        if (start < 0 || length < 0 || start > data.Length - length) throw new InvalidDataException("A WordPerfect structure is truncated or outside its containing data.");
    }
    private static int U16(byte[] data, int at) { CheckRange(data, at, 2); return data[at] | data[at + 1] << 8; }
    private static int I32(byte[] data, int at) {
        CheckRange(data, at, 4);
        int value = data[at] | data[at + 1] << 8 | data[at + 2] << 16 | data[at + 3] << 24;
        if (value < 0) throw new InvalidDataException("A WordPerfect size or offset exceeds the supported range.");
        return value;
    }
    private static double Points(int units) => units * 72d / 1200d;

    private sealed class Packet {
        internal int Type, Offset, Length;
        internal int[] Children = Array.Empty<int>();
    }

    private sealed class Budget {
        internal readonly OfficeLegacyImportLimits Limits;
        internal readonly CancellationToken Cancellation;
        internal readonly HashSet<string> Reported = new(StringComparer.Ordinal);
        internal int RemainingResourceBytes => Limits.MaxResourceBytes - _resourceBytes;
        private int _records, _items, _characters, _resourceBytes;
        private long _imagePixels;
        internal Budget(OfficeLegacyImportLimits limits, CancellationToken cancellation) { Limits = limits; Cancellation = cancellation; }
        internal void Record() {
            Cancellation.ThrowIfCancellationRequested();
            if (_records >= Limits.MaxRecords) throw new InvalidDataException("WordPerfect exceeds the configured record limit.");
            _records++;
        }
        internal void Item() {
            if (_items >= Limits.MaxItems) throw new InvalidDataException("WordPerfect exceeds the configured item limit.");
            _items++;
        }
        internal void Text(int characters) {
            if (characters > Limits.MaxTextCharacters - _characters) throw new InvalidDataException("WordPerfect exceeds the configured text limit.");
            _characters += characters;
        }
        internal void Resource(int bytes) {
            if (bytes > Limits.MaxResourceBytes - _resourceBytes) throw new InvalidDataException("WordPerfect exceeds the configured resource-byte limit.");
            _resourceBytes += bytes;
        }
        internal void Image(long pixels) {
            if (pixels > Limits.MaxImagePixels - _imagePixels) throw new InvalidDataException("WordPerfect exceeds the configured image-pixel limit.");
            _imagePixels += pixels;
        }
    }
}
