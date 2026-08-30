using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading;
using OfficeIMO.Word;

namespace OfficeIMO.Word.Legacy;

/// <summary>Bounded semantic parser for the documented Ami Pro SAM version 4 tagged-text profile.</summary>
internal sealed class AmiProSamParser {
    private readonly string[] _lines;
    private readonly OfficeLegacyImportLimits _limits;
    private readonly CancellationToken _cancellationToken;
    private readonly Dictionary<string, AmiStyle> _styles = new(StringComparer.OrdinalIgnoreCase);
    private readonly HashSet<string> _unknownTags = new(StringComparer.Ordinal);
    private readonly HashSet<string> _unsupportedSections = new(StringComparer.OrdinalIgnoreCase);
    private readonly HashSet<string> _missingStyles = new(StringComparer.OrdinalIgnoreCase);
    private readonly LegacyWordModel _model = new() { Quality = OfficeLegacyImportQuality.Structured };
    private int _characterCount;
    private int _itemCount;
    private int _inferredListCount;
    private int _embeddedSectionCount;
    private int _unsupportedSectionCount;
    private int _malformedStyleBlockCount;
    private int _duplicateStyleCount;
    private int _malformedInlineTagCount;

    internal AmiProSamParser(byte[] data, OfficeLegacyImportLimits limits, CancellationToken cancellationToken) {
        _limits = limits;
        _cancellationToken = cancellationToken;
        ValidateRecordCountAndEncoding(data);
        _cancellationToken.ThrowIfCancellationRequested();
        _lines = Encoding.ASCII.GetString(data).Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
    }

    private void ValidateRecordCountAndEncoding(byte[] data) {
        int records = 1;
        for (int index = 0; index < data.Length; index++) {
            if ((index & 0x0FFF) == 0) _cancellationToken.ThrowIfCancellationRequested();
            if (data[index] > 0x7F) throw new InvalidDataException("Ami Pro SAM contains an extended character byte outside the structured ASCII profile.");
            if ((data[index] == (byte)'\n' || data[index] == (byte)'\r') &&
                (index == 0 || data[index] != (byte)'\n' || data[index - 1] != (byte)'\r') &&
                ++records > _limits.MaxRecords) {
                throw new InvalidDataException("Ami Pro source exceeds the configured record limit.");
            }
        }
    }

    internal LegacyWordModel Parse() {
        if (_lines.Length > _limits.MaxRecords) throw new InvalidDataException("Ami Pro source exceeds the configured record limit.");
        int versionIndex = FindSection("ver");
        if (versionIndex < 0 || versionIndex + 1 >= _lines.Length || !int.TryParse(_lines[versionIndex + 1].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out int version)) {
            throw new InvalidDataException("Ami Pro SAM source has no valid [ver] value.");
        }
        if (version != 4) throw new InvalidDataException($"Ami Pro SAM version {version} is outside the structured sam-v4 profile.");
        _model.Metadata["AmiProVersion"] = version.ToString(CultureInfo.InvariantCulture);
        InventoryUnsupportedSections();
        ParseStyles();
        ParseDocument();
        if (_model.Paragraphs.Count == 0) throw new InvalidDataException("Ami Pro SAM [edoc] section contains no recoverable paragraphs.");
        if (_unknownTags.Count > 0) {
            _model.Findings.Add(LegacyWordAdapterBase.LossFinding("AMIPRO_INLINE_TAG_UNSUPPORTED", "Formatting", $"{_unknownTags.Count} distinct Ami Pro inline tag kinds were omitted."));
        }
        if (_missingStyles.Count > 0) {
            _model.Metadata["AmiProMissingStyleCount"] = _missingStyles.Count.ToString(CultureInfo.InvariantCulture);
            _model.Findings.Add(LegacyWordAdapterBase.LossFinding("AMIPRO_STYLE_MISSING", "Styles", "One or more referenced Ami Pro styles were not defined in parsed [tag] blocks; the distinct total is available in metadata."));
        }
        if (_inferredListCount > 0) _model.Metadata["AmiProInferredListCount"] = _inferredListCount.ToString(CultureInfo.InvariantCulture);
        if (_embeddedSectionCount > 0) _model.Metadata["AmiProEmbeddedSectionCount"] = _embeddedSectionCount.ToString(CultureInfo.InvariantCulture);
        if (_unsupportedSectionCount > 0) _model.Metadata["AmiProUnsupportedSectionCount"] = _unsupportedSectionCount.ToString(CultureInfo.InvariantCulture);
        if (_malformedStyleBlockCount > 0) _model.Metadata["AmiProMalformedStyleBlockCount"] = _malformedStyleBlockCount.ToString(CultureInfo.InvariantCulture);
        if (_duplicateStyleCount > 0) _model.Metadata["AmiProDuplicateStyleCount"] = _duplicateStyleCount.ToString(CultureInfo.InvariantCulture);
        if (_malformedInlineTagCount > 0) _model.Metadata["AmiProMalformedInlineTagCount"] = _malformedInlineTagCount.ToString(CultureInfo.InvariantCulture);
        return _model;
    }

    private void ParseStyles() {
        for (int index = 0; index < _lines.Length; index++) {
            _cancellationToken.ThrowIfCancellationRequested();
            if (!IsSection(_lines[index], "tag")) continue;
            int end = FindNextTopLevelSection(index + 1);
            var block = new List<string>();
            for (int line = index + 1; line < end; line++) block.Add(_lines[line].Trim());
            ConsumeItem("style block");
            AmiStyle? style = ParseStyle(block);
            index = end - 1;
            if (style == null || string.IsNullOrWhiteSpace(style.Name)) {
                if (++_malformedStyleBlockCount == 1) {
                    _model.Findings.Add(LegacyWordAdapterBase.LossFinding("AMIPRO_STYLE_BLOCK_MALFORMED", "Styles", "One or more Ami Pro [tag] style blocks were malformed and could not be projected; the total is available in metadata."));
                }
                continue;
            }
            ConsumeText(style.Name.Length, "style name");
            if (style.FontFamily is { Length: > 0 } fontFamily) ConsumeText(fontFamily.Length, "style font family");
            if (_styles.ContainsKey(style.Name) && ++_duplicateStyleCount == 1) {
                _model.Findings.Add(LegacyWordAdapterBase.LossFinding("AMIPRO_STYLE_DUPLICATE", "Styles", "One or more duplicate Ami Pro style definitions replaced an earlier definition; the total is available in metadata."));
            }
            _styles[style.Name] = style;
            _model.Metadata[$"Style.{_styles.Count}.Name"] = style.Name;
        }
        _model.Metadata["StyleCount"] = _styles.Count.ToString(CultureInfo.InvariantCulture);
    }

    private AmiStyle? ParseStyle(List<string> lines) {
        if (lines.Count < 19 || !string.Equals(lines[2], "[fnt]", StringComparison.OrdinalIgnoreCase)) return null;
        if (!int.TryParse(lines[4], NumberStyles.Integer, CultureInfo.InvariantCulture, out int fontTwips)) return null;
        if (!uint.TryParse(lines[5], NumberStyles.Integer, CultureInfo.InvariantCulture, out uint color)) color = 0;
        if (!uint.TryParse(lines[6], NumberStyles.Integer, CultureInfo.InvariantCulture, out uint formatting)) formatting = 0;
        var style = new AmiStyle {
            Name = UnescapePlain(lines[0]), FontFamily = lines[3], FontSizePoints = Math.Max(1d, fontTwips / 20d),
            ColorHex = $"{color & 0xFF:X2}{(color >> 8) & 0xFF:X2}{(color >> 16) & 0xFF:X2}",
            Bold = (formatting & 1) != 0, Italic = (formatting & 2) != 0,
            Underline = (formatting & 64) != 0 ? WordUnderlineStyle.Double : (formatting & 8) != 0 ? WordUnderlineStyle.Words : (formatting & 4) != 0 ? WordUnderlineStyle.Single : (WordUnderlineStyle?)null
        };
        int alignIndex = lines.FindIndex(static value => string.Equals(value, "[algn]", StringComparison.OrdinalIgnoreCase));
        if (alignIndex >= 0 && alignIndex + 1 < lines.Count && uint.TryParse(lines[alignIndex + 1], out uint alignment)) style.Alignment = DecodeAlignment(alignment);
        int spacingIndex = lines.FindIndex(static value => string.Equals(value, "[spc]", StringComparison.OrdinalIgnoreCase));
        if (spacingIndex >= 0 && spacingIndex + 5 < lines.Count) {
            if (uint.TryParse(lines[spacingIndex + 1], out uint spacing)) {
                if ((spacing & 1) != 0) style.LineSpacingPoints = 12;
                else if ((spacing & 2) != 0) style.LineSpacingPoints = 18;
                else if ((spacing & 4) != 0) style.LineSpacingPoints = 24;
                else if ((spacing & 8) != 0 && int.TryParse(lines[spacingIndex + 2], out int custom)) style.LineSpacingPoints = custom / 20d;
            }
            if (int.TryParse(lines[spacingIndex + 4], out int before)) style.SpacingBeforePoints = before / 20d;
            if (int.TryParse(lines[spacingIndex + 5], out int after)) style.SpacingAfterPoints = after / 20d;
        }
        int breakIndex = lines.FindIndex(static value => string.Equals(value, "[brk]", StringComparison.OrdinalIgnoreCase));
        if (breakIndex >= 0 && breakIndex + 1 < lines.Count && uint.TryParse(lines[breakIndex + 1], out uint breaks)) {
            style.PageBreakBefore = (breaks & 1) != 0;
            style.KeepWithNext = (breaks & 16) != 0;
            style.KeepLinesTogether = (breaks & 4) == 0;
        }
        return style;
    }

    private void ParseDocument() {
        int start = FindSection("edoc");
        if (start < 0) throw new InvalidDataException("Ami Pro SAM version 4 source has no [edoc] section.");
        var paragraphLines = new List<string>();
        for (int index = start + 1; index < _lines.Length; index++) {
            _cancellationToken.ThrowIfCancellationRequested();
            string line = _lines[index];
            if (line.StartsWith("[", StringComparison.Ordinal)) {
                FlushParagraph(paragraphLines);
                break;
            }
            if (line.Length == 0) {
                FlushParagraph(paragraphLines);
                continue;
            }
            if (line.StartsWith(">", StringComparison.Ordinal)) continue;
            paragraphLines.Add(line);
        }
        FlushParagraph(paragraphLines);
    }

    private void FlushParagraph(List<string> lines) {
        if (lines.Count == 0) return;
        string source = string.Join("\n", lines);
        lines.Clear();
        LegacyWordParagraph paragraph = ParseInline(source);
        if (paragraph.Text.Length == 0) return;
        ConsumeItem("paragraph");
        if (paragraph.Text.StartsWith("- ", StringComparison.Ordinal) || paragraph.Text.StartsWith("* ", StringComparison.Ordinal)) {
            paragraph.IsList = true;
            TrimPrefix(paragraph.Runs, 2);
            if (++_inferredListCount == 1) {
                _model.Findings.Add(LegacyWordAdapterBase.LossFinding("AMIPRO_LIST_INFERRED", "Lists", "One or more bullet-list items were inferred from leading source text markers; the total is available in metadata."));
            }
        }
        _model.Paragraphs.Add(paragraph);
    }

    private LegacyWordParagraph ParseInline(string source) {
        var paragraph = new LegacyWordParagraph();
        List<LegacyWordRun> runs = paragraph.Runs;
        var text = new StringBuilder();
        var state = new AmiRunState();
        if (_styles.TryGetValue("Body Text", out AmiStyle? bodyStyle)) {
            paragraph.StyleName = bodyStyle.Name;
            ApplyStyle(bodyStyle, state, paragraph);
        }
        for (int index = 0; index < source.Length;) {
            char value = source[index];
            if (value == '<') {
                if (index + 1 < source.Length && source[index + 1] == '<') { Append(text, '<'); index += 2; continue; }
                int end = source.IndexOf('>', index + 1);
                if (end < 0) { Append(text, '<'); index++; continue; }
                FlushRun(runs, text, state);
                HandleTag(source.Substring(index + 1, end - index - 1), state, paragraph, text);
                index = end + 1;
                continue;
            }
            if (value == '@') {
                if (index + 1 < source.Length && source[index + 1] == '@') { Append(text, '@'); index += 2; continue; }
                int end = source.IndexOf('@', index + 1);
                if (end > index + 1) {
                    FlushRun(runs, text, state);
                    string styleName = source.Substring(index + 1, end - index - 1);
                    ConsumeText(styleName.Length, "style reference");
                    paragraph.StyleName = styleName;
                    if (_styles.TryGetValue(styleName, out AmiStyle? style)) ApplyStyle(style, state, paragraph);
                    else if (_missingStyles.Add(styleName)) ConsumeItem("missing style reference");
                    index = end + 1;
                    continue;
                }
            }
            Append(text, value);
            index++;
        }
        FlushRun(runs, text, state);
        return paragraph;
    }

    private void HandleTag(string tag, AmiRunState state, LegacyWordParagraph paragraph, StringBuilder text) {
        switch (tag) {
            case ";": Append(text, '>'); return;
            case "[": Append(text, '['); return;
            case "+!": state.Bold = true; return;
            case "-!": state.Bold = false; return;
            case "+\"": state.Italic = true; return;
            case "-\"": state.Italic = false; return;
            case "+#": state.Underline = WordUnderlineStyle.Single; return;
            case "-#": state.Underline = null; return;
            case "+)": state.Underline = WordUnderlineStyle.Double; return;
            case "-)": state.Underline = null; return;
            case "+$": state.Underline = WordUnderlineStyle.Words; return;
            case "-$": state.Underline = null; return;
            case "+&": state.VerticalPosition = WordVerticalTextPosition.Superscript; return;
            case "-&": state.VerticalPosition = null; return;
            case "+'": state.VerticalPosition = WordVerticalTextPosition.Subscript; return;
            case "-'": state.VerticalPosition = null; return;
            case "+%": state.Strike = true; return;
            case "-%": state.Strike = false; return;
            case "+@": paragraph.Alignment = WordParagraphAlignment.Left; return;
            case "+A": paragraph.Alignment = WordParagraphAlignment.Right; return;
            case "+B": paragraph.Alignment = WordParagraphAlignment.Center; return;
            case "+C": paragraph.Alignment = WordParagraphAlignment.Both; return;
            case "/R": Append(text, '\''); return;
        }
        if (tag.StartsWith(":S+", StringComparison.Ordinal)) {
            string spacing = tag.Substring(3);
            if (spacing == "-1") paragraph.LineSpacingPoints = 12;
            else if (spacing == "-2") paragraph.LineSpacingPoints = 18;
            else if (spacing == "-3") paragraph.LineSpacingPoints = 24;
            else if (int.TryParse(spacing, out int twips)) paragraph.LineSpacingPoints = twips / 20d;
            else RecordMalformedInlineTag();
            return;
        }
        if (tag == ":f") { state.ResetFont(); return; }
        if (tag.StartsWith(":f", StringComparison.Ordinal)) {
            if (!ParseFont(tag.Substring(2), state)) RecordMalformedInlineTag();
            return;
        }
        if (tag.Length == 2 && tag[0] == '/') {
            Append(text, tag[1] == 'R' ? '\'' : (char)(tag[1] + 0x40));
            return;
        }
        if (tag.Length == 2 && tag[0] == '\\') {
            Append(text, (char)(tag[1] | 0x80));
            return;
        }
        string boundedTag = tag.Length > 16 ? tag.Substring(0, 16) : tag;
        if (_unknownTags.Add(boundedTag)) ConsumeItem("inline-tag kind");
    }

    private bool ParseFont(string value, AmiRunState state) {
        string[] parts = value.Split(',');
        if (parts.Length < 5 || !int.TryParse(parts[0], out int size)) return false;
        string family = parts[1].Length > 1 && char.IsDigit(parts[1][0]) ? parts[1].Substring(1) : parts[1];
        ConsumeText(family.Length, "inline font family");
        state.FontSizePoints = Math.Max(1d, size / 20d);
        state.FontFamily = family;
        if (byte.TryParse(parts[2], out byte red) && byte.TryParse(parts[3], out byte green) && byte.TryParse(parts[4], out byte blue)) state.ColorHex = $"{red:X2}{green:X2}{blue:X2}";
        else return false;
        return true;
    }

    private static void ApplyStyle(AmiStyle style, AmiRunState state, LegacyWordParagraph paragraph) {
        state.Bold = style.Bold;
        state.Italic = style.Italic;
        state.Underline = style.Underline;
        state.FontFamily = style.FontFamily;
        state.FontSizePoints = style.FontSizePoints;
        state.ColorHex = style.ColorHex;
        paragraph.Alignment = style.Alignment;
        paragraph.LineSpacingPoints = style.LineSpacingPoints;
        paragraph.SpacingBeforePoints = style.SpacingBeforePoints;
        paragraph.SpacingAfterPoints = style.SpacingAfterPoints;
        paragraph.PageBreakBefore = style.PageBreakBefore;
        paragraph.KeepWithNext = style.KeepWithNext;
        paragraph.KeepLinesTogether = style.KeepLinesTogether;
    }

    private void FlushRun(List<LegacyWordRun> runs, StringBuilder text, AmiRunState state) {
        if (text.Length == 0) return;
        ConsumeItem("formatted run");
        runs.Add(state.CreateRun(text.ToString()));
        text.Clear();
    }

    private void Append(StringBuilder text, char value) {
        ConsumeText(1, "document text");
        text.Append(value);
    }

    private void ConsumeText(int count, string kind) {
        if (count > _limits.MaxTextCharacters - _characterCount) throw new InvalidDataException($"Ami Pro source exceeds the configured text-character limit while recovering {kind}.");
        _characterCount += count;
    }

    private void RecordMalformedInlineTag() {
        if (++_malformedInlineTagCount == 1) {
            _model.Findings.Add(LegacyWordAdapterBase.LossFinding("AMIPRO_INLINE_TAG_MALFORMED", "Formatting", "One or more recognized Ami Pro inline formatting tags had malformed values and were not fully projected; the total is available in metadata."));
        }
    }

    private int FindSection(string name) {
        for (int index = 0; index < _lines.Length; index++) if (IsSection(_lines[index], name)) return index;
        return -1;
    }

    private int FindNextTopLevelSection(int start) {
        for (int index = start; index < _lines.Length; index++) {
            string value = _lines[index];
            if (IsSection(value, "tag") || IsSection(value, "edoc") || IsSection(value, "lay") || IsSection(value, "frm") || IsSection(value, "ver")) return index;
        }
        return _lines.Length;
    }

    private void RecordUnsupportedSection(string section) {
        if (!_unsupportedSections.Add(section)) return;
        ConsumeItem("section kind");
        if (section.Equals("frm", StringComparison.OrdinalIgnoreCase) || section.IndexOf("obj", StringComparison.OrdinalIgnoreCase) >= 0) {
            _model.InertContent |= OfficeLegacyInertContentKind.EmbeddedObjects;
            if (++_embeddedSectionCount == 1) {
                _model.Findings.Add(LegacyWordAdapterBase.InertFinding("AMIPRO_EMBEDDED_OBJECT_INERT", "EmbeddedObjects", "One or more Ami Pro embedded-object sections were kept inert; the distinct total is available in metadata."));
            }
        } else {
            if (++_unsupportedSectionCount == 1) {
                _model.Findings.Add(LegacyWordAdapterBase.LossFinding("AMIPRO_SECTION_UNSUPPORTED", "Structure", "One or more Ami Pro sections were inventoried but not projected; the distinct total is available in metadata."));
            }
        }
    }

    private void InventoryUnsupportedSections() {
        for (int index = 0; index < _lines.Length; index++) {
            _cancellationToken.ThrowIfCancellationRequested();
            string value = _lines[index];
            int first = 0;
            int last = value.Length - 1;
            while (first <= last && char.IsWhiteSpace(value[first])) first++;
            while (last >= first && char.IsWhiteSpace(value[last])) last--;
            if (last - first < 2 || value[first] != '[' || value[last] != ']') continue;
            int sectionStart = first + 1;
            int sectionLength = last - sectionStart;
            while (sectionLength > 0 && char.IsWhiteSpace(value[sectionStart])) { sectionStart++; sectionLength--; }
            while (sectionLength > 0 && char.IsWhiteSpace(value[sectionStart + sectionLength - 1])) sectionLength--;
            if (sectionLength == 0) continue;
            bool validIdentifier = sectionLength <= 64;
            for (int character = 0; validIdentifier && character < sectionLength; character++) {
                char current = value[sectionStart + character];
                validIdentifier = char.IsLetterOrDigit(current) || current == '_' || current == '-';
            }
            string section = validIdentifier ? value.Substring(sectionStart, sectionLength) : "invalid-or-overlong";
            if (section.Equals("ver", StringComparison.OrdinalIgnoreCase) || section.Equals("tag", StringComparison.OrdinalIgnoreCase) ||
                section.Equals("fnt", StringComparison.OrdinalIgnoreCase) || section.Equals("algn", StringComparison.OrdinalIgnoreCase) ||
                section.Equals("spc", StringComparison.OrdinalIgnoreCase) || section.Equals("brk", StringComparison.OrdinalIgnoreCase) ||
                section.Equals("edoc", StringComparison.OrdinalIgnoreCase)) continue;
            RecordUnsupportedSection(section);
        }
    }

    private void ConsumeItem(string kind) {
        if (++_itemCount > _limits.MaxItems) throw new InvalidDataException($"Ami Pro source exceeds the configured item limit while recovering a {kind}.");
    }

    private static bool IsSection(string value, string name) {
        int first = 0;
        int last = value.Length - 1;
        while (first <= last && char.IsWhiteSpace(value[first])) first++;
        while (last >= first && char.IsWhiteSpace(value[last])) last--;
        if (last - first != name.Length + 1 || value[first] != '[' || value[last] != ']') return false;
        return string.Compare(value, first + 1, name, 0, name.Length, StringComparison.OrdinalIgnoreCase) == 0;
    }
    private static string UnescapePlain(string value) => value.Replace("@@", "@").Replace("<<", "<").Replace("<[>", "[").Replace("<;>", ">");
    private static WordParagraphAlignment? DecodeAlignment(uint value) => (value & 8) != 0 ? WordParagraphAlignment.Both : (value & 4) != 0 ? WordParagraphAlignment.Center : (value & 2) != 0 ? WordParagraphAlignment.Right : (value & 1) != 0 ? WordParagraphAlignment.Left : (WordParagraphAlignment?)null;

    private static void TrimPrefix(List<LegacyWordRun> runs, int count) {
        while (count > 0 && runs.Count > 0) {
            LegacyWordRun first = runs[0];
            if (first.Text.Length <= count) { count -= first.Text.Length; runs.RemoveAt(0); continue; }
            runs[0] = new LegacyWordRun(first.Text.Substring(count)) { Bold = first.Bold, Italic = first.Italic, Strike = first.Strike, Underline = first.Underline, VerticalPosition = first.VerticalPosition, FontSizePoints = first.FontSizePoints, FontFamily = first.FontFamily, ColorHex = first.ColorHex };
            count = 0;
        }
    }

    private sealed class AmiRunState {
        internal bool Bold;
        internal bool Italic;
        internal bool Strike;
        internal WordUnderlineStyle? Underline;
        internal WordVerticalTextPosition? VerticalPosition;
        internal double? FontSizePoints;
        internal string? FontFamily;
        internal string? ColorHex;
        internal void ResetFont() { FontSizePoints = null; FontFamily = null; ColorHex = null; }
        internal LegacyWordRun CreateRun(string text) => new(text) { Bold = Bold, Italic = Italic, Strike = Strike, Underline = Underline, VerticalPosition = VerticalPosition, FontSizePoints = FontSizePoints, FontFamily = FontFamily, ColorHex = ColorHex };
    }

    private sealed class AmiStyle {
        internal string Name = string.Empty;
        internal bool Bold;
        internal bool Italic;
        internal WordUnderlineStyle? Underline;
        internal double? FontSizePoints;
        internal string? FontFamily;
        internal string? ColorHex;
        internal WordParagraphAlignment? Alignment;
        internal double? LineSpacingPoints;
        internal double? SpacingBeforePoints;
        internal double? SpacingAfterPoints;
        internal bool PageBreakBefore;
        internal bool KeepWithNext;
        internal bool KeepLinesTogether;
    }
}
