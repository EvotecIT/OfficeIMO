using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;
using OfficeIMO.Word;

namespace OfficeIMO.Word.Legacy;

/// <summary>Bounded parser for the documented WordStar 3-7 character and symmetrical-sequence grammar.</summary>
internal sealed class WordStarStructuredParser {
    private readonly byte[] _data;
    private readonly OfficeLegacyImportLimits _limits;
    private readonly CancellationToken _cancellationToken;
    private readonly LegacyWordModel _model = new() { Quality = OfficeLegacyImportQuality.Structured };
    private readonly List<LegacyWordRun> _runs = new();
    private readonly StringBuilder _text = new();
    private readonly RunState _state = new();
    private readonly Dictionary<byte, int> _partialSequenceCounts = new();
    private readonly Dictionary<byte, int> _unsupportedSequenceCounts = new();
    private readonly Dictionary<byte, int> _unsupportedControlCounts = new();
    private int _characterCount;
    private int _recordCount;
    private int _sequenceCount;
    private int _itemCount;
    private int _unknownDotCommandCount;
    private int _graphicsReferenceCount;
    private int _inferredListCount;
    private int _headerFooterMetadataOnlyCount;
    private string? _paragraphStyleName;
    private bool _nextPageBreak;

    internal WordStarStructuredParser(byte[] data, OfficeLegacyImportLimits limits, CancellationToken cancellationToken) {
        _data = data;
        _limits = limits;
        _cancellationToken = cancellationToken;
    }

    internal LegacyWordModel Parse() {
        for (int index = 0; index < _data.Length;) {
            _cancellationToken.ThrowIfCancellationRequested();
            byte value = _data[index];
            if (value == 0x1A) {
                ValidateTrailingPadding(index + 1);
                break;
            }
            if (value == 0x1D) {
                index = ParseSymmetricalSequence(index);
                continue;
            }
            if (value == 0x8D && index + 1 < _data.Length && _data[index + 1] == 0x0A) {
                AppendSoftBreak();
                index += 2;
                continue;
            }
            if (value == 0x0D && index + 1 < _data.Length && _data[index + 1] == 0x0A) {
                FlushParagraph(explicitBreak: true);
                index += 2;
                continue;
            }
            switch (value) {
                case 0x02: Toggle(static state => state.Bold = !state.Bold); break;
                case 0x09: Append('\t'); break;
                case 0x0A: FlushParagraph(explicitBreak: true); break;
                case 0x0C: FlushParagraph(); _nextPageBreak = true; break;
                case 0x0D: FlushParagraph(explicitBreak: true); break;
                case 0x0F: Append('\u00A0'); break;
                case 0x13: Toggle(static state => state.Underline = !state.Underline); break;
                case 0x14: Toggle(static state => state.Superscript = !state.Superscript); break;
                case 0x16: Toggle(static state => state.Subscript = !state.Subscript); break;
                case 0x18: Toggle(static state => state.Strike = !state.Strike); break;
                case 0x19: Toggle(static state => state.Italic = !state.Italic); break;
                case 0x1E: RecordUnsupportedControl(value); break;
                case 0x1F: Append('-'); break;
                case 0x1B:
                    index = ParseExtendedCharacter(index);
                    continue;
                default:
                    byte character = value >= 0x80 ? (byte)(value & 0x7F) : value;
                    if (character >= 0x20 && character != 0x7F) Append((char)character);
                    else RecordUnsupportedControl(value);
                    break;
            }
            index++;
        }
        FlushParagraph(force: _model.Paragraphs.Count == 0);
        if (_sequenceCount > 0) _model.Metadata["WordStarSymmetricalSequenceCount"] = _sequenceCount.ToString(System.Globalization.CultureInfo.InvariantCulture);
        if (_unknownDotCommandCount > 0) _model.Metadata["WordStarUnknownDotCommandCount"] = _unknownDotCommandCount.ToString(System.Globalization.CultureInfo.InvariantCulture);
        if (_graphicsReferenceCount > 0) _model.Metadata["WordStarGraphicsReferenceCount"] = _graphicsReferenceCount.ToString(System.Globalization.CultureInfo.InvariantCulture);
        if (_inferredListCount > 0) _model.Metadata["WordStarInferredListCount"] = _inferredListCount.ToString(System.Globalization.CultureInfo.InvariantCulture);
        if (_headerFooterMetadataOnlyCount > 0) _model.Metadata["WordStarHeaderFooterMetadataOnlyCount"] = _headerFooterMetadataOnlyCount.ToString(System.Globalization.CultureInfo.InvariantCulture);
        AddSequenceCountMetadata("WordStarPartialSequence", _partialSequenceCounts);
        AddSequenceCountMetadata("WordStarUnsupportedSequence", _unsupportedSequenceCounts);
        AddSequenceCountMetadata("WordStarUnsupportedControl", _unsupportedControlCounts);
        return _model;
    }

    private int ParseSymmetricalSequence(int index) {
        ConsumeRecord("symmetrical sequence");
        _sequenceCount++;
        if (index + 4 > _data.Length) throw new InvalidDataException("Truncated WordStar symmetrical-sequence header.");
        int count = _data[index + 1] | (_data[index + 2] << 8);
        int totalLength = count + 3;
        if (totalLength < 7 || index > _data.Length - totalLength) throw new InvalidDataException("Invalid WordStar symmetrical-sequence length.");
        int suffix = index + totalLength - 3;
        if (_data[suffix] != _data[index + 1] || _data[suffix + 1] != _data[index + 2] || _data[suffix + 2] != 0x1D) {
            throw new InvalidDataException("WordStar symmetrical-sequence delimiter or repeated length is invalid.");
        }

        byte type = _data[index + 3];
        int payloadOffset = index + 4;
        int payloadLength = totalLength - 7;
        switch (type) {
            case 0x03: AddNote(LegacyWordNoteKind.Footnote, ExtractSequenceText(payloadOffset, payloadLength)); break;
            case 0x04: AddNote(LegacyWordNoteKind.Endnote, ExtractSequenceText(payloadOffset, payloadLength)); break;
            case 0x05: AddNote(LegacyWordNoteKind.Annotation, ExtractSequenceText(payloadOffset, payloadLength)); break;
            case 0x06: AddNote(LegacyWordNoteKind.Comment, ExtractSequenceText(payloadOffset, payloadLength)); break;
            case 0x10:
                string payloadText = ExtractSequenceText(payloadOffset, payloadLength);
                if (payloadText.Length > 0) {
                    ConsumeItem("resource reference");
                    ConsumeText(payloadText.Length);
                    _model.Resources.Add(new LegacyWordResource("Graphics", payloadText));
                    _model.InertContent |= OfficeLegacyInertContentKind.ExternalLinks;
                    if (++_graphicsReferenceCount == 1) {
                        _model.Findings.Add(LegacyWordAdapterBase.InertFinding("WORDSTAR_GRAPHICS_REFERENCE_INERT", "Images", "One or more WordStar graphics references were recorded but were not resolved or loaded; the total is available in metadata."));
                    }
                }
                break;
            case 0x11:
                payloadText = ExtractSequenceText(payloadOffset, payloadLength);
                FlushRun();
                if (payloadText.Length > 0) {
                    ConsumeText(payloadText.Length);
                    _paragraphStyleName = payloadText;
                    _model.Metadata["WordStarParagraphStyle"] = payloadText;
                }
                break;
            case 0x00:
            case 0x01:
            case 0x02:
            case 0x15:
                if (Increment(_partialSequenceCounts, type) == 1) {
                    _model.Findings.Add(LegacyWordAdapterBase.LossFinding("WORDSTAR_SEQUENCE_PARTIAL", "Formatting", $"WordStar sequence type 0x{type:X2} was validated but is not projected by this profile; its total is available in metadata."));
                }
                break;
            default:
                if (Increment(_unsupportedSequenceCounts, type) == 1) {
                    _model.Findings.Add(LegacyWordAdapterBase.LossFinding("WORDSTAR_SEQUENCE_UNSUPPORTED", "Structure", $"WordStar sequence type 0x{type:X2} was kept inert and omitted; its total is available in metadata."));
                }
                break;
        }
        return index + totalLength;
    }

    private int ParseExtendedCharacter(int index) {
        if (index + 1 >= _data.Length) throw new InvalidDataException("Truncated WordStar extended-character sequence.");
        byte value = (byte)(_data[index + 1] & 0x7F);
        if (value >= 0x20 && value != 0x7F) Append((char)value);
        else RecordUnsupportedControl(_data[index + 1]);
        return index + 2 < _data.Length && _data[index + 2] == 0x1C ? index + 3 : index + 2;
    }

    private string ExtractSequenceText(int offset, int length) {
        var result = new StringBuilder(Math.Min(length, 1024));
        int end = offset + length;
        int remaining = _limits.MaxTextCharacters - _characterCount;
        for (int index = offset; index < end; index++) {
            _cancellationToken.ThrowIfCancellationRequested();
            byte value = (byte)(_data[index] & 0x7F);
            char? recovered = value == 0x0D || value == 0x0A || value == 0x09
                ? ' '
                : value >= 0x20 && value != 0x7F ? (char)value : (char?)null;
            if (!recovered.HasValue) continue;
            if (result.Length >= remaining) throw new InvalidDataException("WordStar sequence text exceeds the configured text-character limit.");
            result.Append(recovered.Value);
        }
        return result.ToString().Trim('\0', ' ');
    }

    private void AddNote(LegacyWordNoteKind kind, string text) {
        if (string.IsNullOrWhiteSpace(text)) return;
        ConsumeItem("note");
        ConsumeText(text.Length);
        _model.Notes.Add(new LegacyWordNote(kind, text));
        if (_model.Notes.Count == 1) _model.Findings.Add(LegacyWordAdapterBase.LossFinding("WORDSTAR_NOTE_ANCHOR_APPROXIMATED", "Notes", "WordStar note text was recovered, but its exact source anchor was unavailable; conversion appends a labeled note paragraph."));
    }

    private void AppendSoftBreak() {
        if (_text.Length > 0 && !char.IsWhiteSpace(_text[_text.Length - 1]) && _text[_text.Length - 1] != '-') Append(' ');
    }

    private void Append(char value) {
        if (++_characterCount > _limits.MaxTextCharacters) throw new InvalidDataException("WordStar source exceeds the configured text-character limit.");
        _text.Append(value);
    }

    private void Toggle(Action<RunState> update) {
        FlushRun();
        update(_state);
        if (_state.Superscript && _state.Subscript) _state.Subscript = false;
    }

    private void FlushRun() {
        if (_text.Length == 0) return;
        ConsumeItem("formatted run");
        _runs.Add(_state.CreateRun(_text.ToString()));
        _text.Clear();
    }

    private void FlushParagraph(bool force = false, bool explicitBreak = false) {
        FlushRun();
        if (_runs.Count == 0 && !force && !explicitBreak) return;
        ConsumeRecord("paragraph");
        string text = JoinText(_runs);
        if (TryHandleDotCommand(text)) {
            _runs.Clear();
            return;
        }
        ConsumeItem("paragraph");
        var paragraph = new LegacyWordParagraph(_runs) {
            PageBreakBefore = _nextPageBreak,
            StyleName = _paragraphStyleName
        };
        _nextPageBreak = false;
        if (text.StartsWith("- ", StringComparison.Ordinal) || text.StartsWith("* ", StringComparison.Ordinal)) {
            paragraph.IsList = true;
            RemoveRunPrefix(paragraph.Runs, 2);
            if (++_inferredListCount == 1) {
                _model.Findings.Add(LegacyWordAdapterBase.LossFinding("WORDSTAR_LIST_INFERRED", "Lists", "One or more bullet-list items were inferred from leading text markers because WordStar stores them as document text; the total is available in metadata."));
            }
        }
        _model.Paragraphs.Add(paragraph);
        _runs.Clear();
    }

    private void ConsumeItem(string kind) {
        if (++_itemCount > _limits.MaxItems) throw new InvalidDataException($"WordStar source exceeds the configured item limit while recovering a {kind}.");
    }

    private void ConsumeRecord(string kind) {
        if (++_recordCount > _limits.MaxRecords) throw new InvalidDataException($"WordStar source exceeds the configured record limit while inspecting a {kind}.");
    }

    private void ConsumeText(int count) {
        if (count > _limits.MaxTextCharacters - _characterCount) throw new InvalidDataException("WordStar source exceeds the configured text-character limit.");
        _characterCount += count;
    }

    private bool TryHandleDotCommand(string text) {
        if (text.Length < 3 || text[0] != '.') return false;
        string command = text.Substring(1, 2).ToUpperInvariant();
        string argument = text.Length > 3 ? text.Substring(3).Trim() : string.Empty;
        switch (command) {
            case "PA": _nextPageBreak = true; return true;
            case "HE":
                _model.Metadata["Header"] = argument;
                RecordHeaderFooterMetadataOnly();
                return true;
            case "FO":
                _model.Metadata["Footer"] = argument;
                RecordHeaderFooterMetadataOnly();
                return true;
            default:
                _unknownDotCommandCount++;
                if (_unknownDotCommandCount == 1) {
                    _model.Findings.Add(LegacyWordAdapterBase.LossFinding("WORDSTAR_DOT_COMMAND", "Layout", "One or more unrecognized WordStar dot commands were kept inert and omitted; the total is available in metadata."));
                }
                return true;
        }
    }

    private void RecordUnsupportedControl(byte value) {
        if (Increment(_unsupportedControlCounts, value) == 1 && _unsupportedControlCounts.Count == 1) {
            _model.Findings.Add(LegacyWordAdapterBase.LossFinding("WORDSTAR_CONTROL_UNSUPPORTED", "Structure", "One or more unsupported WordStar standalone control bytes were omitted; per-control totals are available in metadata."));
        }
    }

    private void RecordHeaderFooterMetadataOnly() {
        if (++_headerFooterMetadataOnlyCount == 1) {
            _model.Findings.Add(LegacyWordAdapterBase.LossFinding("WORDSTAR_HEADER_FOOTER_METADATA_ONLY", "Layout", "WordStar header and footer dot commands were retained as metadata but are not projected into document header/footer parts; the total is available in metadata."));
        }
    }

    private void ValidateTrailingPadding(int offset) {
        for (int index = offset; index < _data.Length; index++) {
            byte value = _data[index];
            if (value != 0x00 && value != 0x1A) throw new InvalidDataException("WordStar source contains non-padding data after its EOF marker.");
        }
    }

    private static string JoinText(List<LegacyWordRun> runs) {
        var text = new StringBuilder();
        foreach (LegacyWordRun run in runs) text.Append(run.Text);
        return text.ToString();
    }

    private void AddSequenceCountMetadata(string prefix, Dictionary<byte, int> counts) {
        foreach (KeyValuePair<byte, int> entry in counts) {
            _model.Metadata[$"{prefix}.0x{entry.Key:X2}Count"] = entry.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
    }

    private static int Increment(Dictionary<byte, int> counts, byte type) {
        counts.TryGetValue(type, out int count);
        count++;
        counts[type] = count;
        return count;
    }

    private static void RemoveRunPrefix(List<LegacyWordRun> runs, int count) {
        while (count > 0 && runs.Count > 0) {
            LegacyWordRun first = runs[0];
            if (first.Text.Length <= count) {
                count -= first.Text.Length;
                runs.RemoveAt(0);
            } else {
                LegacyWordRun replacement = CopyRun(first, first.Text.Substring(count));
                runs[0] = replacement;
                count = 0;
            }
        }
    }

    private static LegacyWordRun CopyRun(LegacyWordRun source, string text) => new(text) {
        Bold = source.Bold, Italic = source.Italic, Strike = source.Strike, Underline = source.Underline,
        VerticalPosition = source.VerticalPosition, FontSizePoints = source.FontSizePoints,
        FontFamily = source.FontFamily, ColorHex = source.ColorHex
    };

    private sealed class RunState {
        internal bool Bold;
        internal bool Italic;
        internal bool Underline;
        internal bool Strike;
        internal bool Superscript;
        internal bool Subscript;

        internal LegacyWordRun CreateRun(string text) => new(text) {
            Bold = Bold,
            Italic = Italic,
            Underline = Underline ? WordUnderlineStyle.Single : (WordUnderlineStyle?)null,
            Strike = Strike,
            VerticalPosition = Superscript ? WordVerticalTextPosition.Superscript : Subscript ? WordVerticalTextPosition.Subscript : (WordVerticalTextPosition?)null
        };
    }
}
