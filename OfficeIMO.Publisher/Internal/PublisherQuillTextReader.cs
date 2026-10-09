using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

internal sealed class PublisherQuillTextReader {
    private readonly PublisherBinaryData _data;
    private readonly PublisherParseContext _context;
    internal PublisherQuillTextReader(PublisherBinaryData data, PublisherParseContext context) { _data = data; _context = context; }
    internal PublisherQuillText Read(IReadOnlyList<OfficeColor?> palette) {
        IReadOnlyList<PublisherQuillChunk> chunks = PublisherQuillDirectory.Read(_data, _context);
        PublisherQuillChunk? text = PublisherQuillDirectory.Single(chunks, "TEXT");
        var result = new PublisherQuillText();
        if (text == null || text.Length == 0) return result;
        if ((text.Length & 1) != 0 || text.Length / 2 > _context.Options.Limits.MaxTextCharacters)
            throw new InvalidDataException("Publisher text exceeds the configured character limit or has odd UTF-16 length.");
        var idsChunk = PublisherQuillDirectory.Single(chunks, "SYID") ?? throw new InvalidDataException("Publisher text story identifiers are missing.");
        var lengthsChunk = PublisherQuillDirectory.Single(chunks, "STRS") ?? throw new InvalidDataException("Publisher text story lengths are missing.");
        _data.Range(idsChunk.Offset, 8, idsChunk.End);
        _data.Range(lengthsChunk.Offset, 8, lengthsChunk.End);
        int count = _data.Offset(_data.U32(idsChunk.Offset + 4));
        if (count != _data.U32(lengthsChunk.Offset) || count > _context.Options.Limits.MaxItems)
            throw new InvalidDataException("Publisher story lengths and identifiers disagree or exceed the object limit.");
        int lengths = checked(lengthsChunk.Offset + 4 + _data.Offset(_data.U32(lengthsChunk.Offset + 4)));
        _data.Range(idsChunk.Offset + 8, checked(count * 4), idsChunk.End);
        _data.Range(lengths, checked(count * 4), lengthsChunk.End);
        var styles = new PublisherQuillStyleReader(_data, _context, chunks, palette, text).Read();
        int position = 0;
        var rawStories = new List<(uint Id, string Text, int Offset)>();
        for (int i = 0; i < count; i++) {
            _context.Record();
            uint id = _data.U32(idsChunk.Offset + 8 + i * 4);
            int characters = _data.Offset(_data.U32(lengths + i * 4));
            if (characters > 100000) throw new InvalidDataException("Publisher story exceeds the shared drawing text limit of 100,000 characters.");
            if (characters > text.Length / 2 - position) throw new InvalidDataException("Publisher story extends past its text section.");
            string value = _data.Utf16(text.Offset + position * 2, characters * 2);
            IReadOnlyList<OfficeRichTextParagraph> paragraphs = ReadParagraphs(value, position, styles);
            if (result.Stories.ContainsKey(id)) throw new InvalidDataException("Duplicate Publisher text story identifier.");
            result.Stories.Add(id, new PublisherTextStory(id, Normalize(value), paragraphs));
            rawStories.Add((id, value, position));
            position += characters;
        }
        if (position != text.Length / 2) {
            _context.Add("PUB_UNASSIGNED_TEXT", "Text outside the declared native stories was not assigned to a publication object.",
                OfficeConversionLossKind.Omission, "Quill/TEXT");
        }
        foreach (PublisherQuillChunk chunk in chunks.Where(item => item.Tag == "TCD ")) {
            _data.Range(chunk.Offset, 12, chunk.End);
            int cells = checked(_data.Offset(_data.U32(chunk.Offset)) + 1);
            _data.Range(chunk.Offset + 12, checked(cells * 4), chunk.End);
            if (chunk.Id >= rawStories.Count || cells > _context.Options.Limits.MaxItems)
                throw new InvalidDataException("Publisher table text refers to an invalid story or exceeds its cell limit.");
            var story = rawStories[chunk.Id];
            var cellParagraphs = new List<IReadOnlyList<OfficeRichTextParagraph>>();
            int start = 0;
            for (int i = 0; i < cells; i++) {
                _context.Record();
                // Non-final entries point to the terminating UTF-16 character;
                // the final entry is the exclusive story endpoint.
                int end = checked(_data.Offset(_data.U32(chunk.Offset + 12 + i * 4)) + (i < cells - 1 ? 1 : 0));
                if (end < start || end > story.Text.Length) throw new InvalidDataException("Publisher cell text endpoints are outside their story.");
                cellParagraphs.Add(ReadParagraphs(story.Text.Substring(start, end - start), story.Offset + start, styles));
                start = end;
            }
            if (start != story.Text.Length) throw new InvalidDataException("Publisher cell text does not cover its complete story.");
            if (!result.CellParagraphs.ContainsKey(story.Id)) result.CellParagraphs.Add(story.Id, cellParagraphs);
            else throw new InvalidDataException("Duplicate Publisher table text definition.");
        }
        return result;
    }
    private IReadOnlyList<OfficeRichTextParagraph> ReadParagraphs(string text, int storyOffset, PublisherQuillStyles styles) {
        var paragraphs = new List<OfficeRichTextParagraph>();
        int start = 0, runCount = 0, projectedCharacters = 0;
        // A trailing native paragraph terminator belongs to its paragraph, not an additional empty paragraph.
        while (start < text.Length) {
            _context.Record();
            int newline = text.IndexOf('\r', start);
            int end = newline < 0 ? text.Length : newline;
            PublisherParagraphStyle paragraph = styles.ParagraphAt(storyOffset + start);
            if ((paragraph.Left ?? 0) + Math.Min(0, paragraph.FirstLine ?? 0) < 0 || paragraph.Right < 0)
                _context.Add("PUB_NEGATIVE_PARAGRAPH_MARGIN_APPROXIMATED", "A paragraph extending outside its text frame was clamped to that frame.",
                    OfficeConversionLossKind.Approximation, "Quill/FDPP");
            int cursor = start;
            var runs = new List<OfficeRichTextRun>();
            while (cursor < end) {
                PublisherCharacterRange? range = styles.CharacterAt(storyOffset + cursor);
                int next = range == null ? end : Math.Min(end, range.End - storyOffset);
                if (next <= cursor) throw new InvalidDataException("Publisher character formatting does not advance.");
                runs.Add(styles.Run(Normalize(text.Substring(cursor, next - cursor)), range?.Style, paragraph));
                if (++runCount > 4096) throw new InvalidDataException("Publisher story exceeds the shared rich text run limit.");
                cursor = next;
            }
            if (runs.Count == 0) { runs.Add(styles.Run(string.Empty, null, paragraph)); runCount++; }
            OfficeTextParagraphLabel? label = ListLabel(paragraph, runs[0], styles);
            projectedCharacters = checked(projectedCharacters + runs.Sum(run => run.Text.Length)
                + (label?.Run.Text.Length ?? 0) + (paragraphs.Count == 0 ? 0 : 1));
            if (projectedCharacters > 100000) throw new InvalidDataException("Publisher story exceeds the shared drawing text limit, including list labels.");
            OfficeRichTextParagraph projected = label == null
                ? new OfficeRichTextParagraph(runs, paragraph.Alignment ?? OfficeTextAlignment.Left, paragraph.LineHeight,
                    paragraph.Margins, paragraph.Indent, paragraph.LineHeightFactor)
                : new OfficeRichTextParagraph(runs, label, paragraph.Alignment ?? OfficeTextAlignment.Left, paragraph.LineHeight,
                    paragraph.Margins, paragraph.Indent, paragraph.LineHeightFactor);
            if (label != null && ++runCount > 4096) throw new InvalidDataException("Publisher story exceeds the shared rich text run limit, including list labels.");
            paragraphs.Add(projected.WithTabStops(paragraph.TabStops));
            if (runCount + paragraphs.Count - 1 > 4096) throw new InvalidDataException("Publisher story exceeds the shared rich text run limit, including paragraph separators.");
            if (paragraphs.Count > 4096) throw new InvalidDataException("Publisher paragraph limit exceeded.");
            start = newline < 0 ? text.Length : end + 1;
        }
        return paragraphs;
    }

    private OfficeTextParagraphLabel? ListLabel(PublisherParagraphStyle paragraph, OfficeRichTextRun text, PublisherQuillStyles styles) {
        uint bullet = paragraph.List?.Bullet ?? 0;
        if (bullet == 0) return null;
        if (bullet > 0x10FFFF || bullet is >= 0xD800 and <= 0xDFFF || char.IsControl(char.ConvertFromUtf32((int)bullet), 0))
            throw new InvalidDataException("Invalid Publisher list-label Unicode scalar.");
        string native = char.ConvertFromUtf32((int)bullet);
        OfficeRichTextRun label = styles.Run(native, new PublisherCharacterStyle {
            Font = paragraph.LabelFont, Size = paragraph.LabelSize > 0 ? paragraph.LabelSize : text.FontSize, Color = null
        }, paragraph);
        (string marker, bool useTextFont) = OfficeTextListMarkerNormalizer.Normalize(native, label.FontFamily);
        label = new OfficeRichTextRun(marker, label.FontSize, text.Color, fontFamily: useTextFont ? text.FontFamily : label.FontFamily);
        double first = paragraph.Margins.Left + paragraph.Indent.FirstLineOffset;
        return OfficeTextParagraphLabel.AtPosition(label, first, textPosition: paragraph.LabelTextPosition
            ?? paragraph.Margins.Left + paragraph.Indent.ContinuationLineOffset);
    }
    internal static string Normalize(string text) => text.Replace('\r', '\n').Replace("\0", string.Empty);
}

internal sealed class PublisherQuillText {
    internal Dictionary<uint, PublisherTextStory> Stories { get; } = new();
    internal Dictionary<uint, List<IReadOnlyList<OfficeRichTextParagraph>>> CellParagraphs { get; } = new();
}
