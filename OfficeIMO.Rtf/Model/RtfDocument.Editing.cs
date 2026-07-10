namespace OfficeIMO.Rtf;

/// <content>Provides semantic block, text, bookmark, and clone operations.</content>
public sealed partial class RtfDocument {
    /// <summary>Creates an independent semantic clone through the deterministic RTF representation.</summary>
    public RtfDocument Clone() => Read(ToRtf(new RtfWriteOptions { IncludeGenerator = false })).Document;

    /// <summary>Adds an existing semantic block at the end of the document.</summary>
    public void AddBlock(IRtfBlock block) => InsertBlock(_blocks.Count, block);

    /// <summary>Inserts an existing semantic block at the specified document block index.</summary>
    public void InsertBlock(int index, IRtfBlock block) {
        if (block == null) throw new ArgumentNullException(nameof(block));
        if (index < 0 || index > _blocks.Count) throw new ArgumentOutOfRangeException(nameof(index));
        ValidateEditableBlock(block);
        if (_blocks.Contains(block)) throw new InvalidOperationException("The block already belongs to this document.");

        InsertIntoSection(index, block);
        _blocks.Insert(index, block);
        RebuildParagraphIndex();
    }

    /// <summary>Creates and inserts a paragraph at the specified document block index.</summary>
    public RtfParagraph InsertParagraph(int index, string? text = null) {
        var paragraph = new RtfParagraph();
        if (!string.IsNullOrEmpty(text)) paragraph.AddText(text!);
        InsertBlock(index, paragraph);
        return paragraph;
    }

    /// <summary>Creates and inserts a table at the specified document block index.</summary>
    public RtfTable InsertTable(int index, int rows, int columns) {
        if (rows < 0) throw new ArgumentOutOfRangeException(nameof(rows));
        if (columns <= 0) throw new ArgumentOutOfRangeException(nameof(columns));
        var table = new RtfTable();
        for (int rowIndex = 0; rowIndex < rows; rowIndex++) {
            RtfTableRow row = table.AddRow();
            for (int columnIndex = 0; columnIndex < columns; columnIndex++) row.AddCell((columnIndex + 1) * 2400);
        }

        InsertBlock(index, table);
        return table;
    }

    /// <summary>Creates and inserts an image at the specified document block index.</summary>
    public RtfImage InsertImage(int index, RtfImageFormat format, byte[] data) {
        var image = new RtfImage(format, data);
        InsertBlock(index, image);
        return image;
    }

    /// <summary>Removes and returns the block at the specified index.</summary>
    public IRtfBlock RemoveBlockAt(int index) {
        if (index < 0 || index >= _blocks.Count) throw new ArgumentOutOfRangeException(nameof(index));
        IRtfBlock block = _blocks[index];
        _blocks.RemoveAt(index);
        foreach (RtfSection section in _sections) section.RemoveBlock(block);
        RebuildParagraphIndex();
        return block;
    }

    /// <summary>Removes a block when it belongs to this document.</summary>
    public bool RemoveBlock(IRtfBlock block) {
        if (block == null) throw new ArgumentNullException(nameof(block));
        int index = _blocks.IndexOf(block);
        if (index < 0) return false;
        RemoveBlockAt(index);
        return true;
    }

    /// <summary>Moves a document block to its final zero-based index.</summary>
    public void MoveBlock(int fromIndex, int toIndex) {
        if (fromIndex < 0 || fromIndex >= _blocks.Count) throw new ArgumentOutOfRangeException(nameof(fromIndex));
        if (toIndex < 0 || toIndex >= _blocks.Count) throw new ArgumentOutOfRangeException(nameof(toIndex));
        if (fromIndex == toIndex) return;

        IRtfBlock block = _blocks[fromIndex];
        _blocks.RemoveAt(fromIndex);
        foreach (RtfSection section in _sections) section.RemoveBlock(block);
        InsertIntoSection(toIndex, block);
        _blocks.Insert(toIndex, block);
        RebuildParagraphIndex();
    }

    /// <summary>Replaces text across adjacent runs throughout the document body, headers, footers, and notes.</summary>
    public int ReplaceText(string oldText, string newText, StringComparison comparison = StringComparison.Ordinal) {
        var paragraphs = new HashSet<RtfParagraph>();
        CollectParagraphs(_blocks, paragraphs);
        foreach (RtfHeaderFooter headerFooter in _headerFooters) foreach (RtfParagraph paragraph in headerFooter.Paragraphs) paragraphs.Add(paragraph);
        foreach (RtfNote note in _notes) foreach (RtfParagraph paragraph in note.Paragraphs) paragraphs.Add(paragraph);
        return paragraphs.Sum(paragraph => paragraph.ReplaceText(oldText, newText, comparison));
    }

    /// <summary>Replaces all inline content inside the first matching named bookmark range.</summary>
    public bool ReplaceBookmarkText(string name, string replacement) {
        if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Bookmark name cannot be empty.", nameof(name));
        if (replacement == null) throw new ArgumentNullException(nameof(replacement));
        var paragraphs = new List<RtfParagraph>();
        CollectParagraphsInOrder(_blocks, paragraphs);

        int startParagraph = -1;
        int startInline = -1;
        for (int paragraphIndex = 0; paragraphIndex < paragraphs.Count; paragraphIndex++) {
            int marker = paragraphs[paragraphIndex].FindBookmark(RtfBookmarkMarkerKind.Start, name);
            if (marker >= 0) {
                startParagraph = paragraphIndex;
                startInline = marker;
                break;
            }
        }

        if (startParagraph < 0) return false;
        for (int paragraphIndex = startParagraph; paragraphIndex < paragraphs.Count; paragraphIndex++) {
            int endInline = paragraphs[paragraphIndex].FindBookmark(RtfBookmarkMarkerKind.End, name);
            if (endInline < 0 || (paragraphIndex == startParagraph && endInline <= startInline)) continue;

            if (paragraphIndex == startParagraph) {
                paragraphs[paragraphIndex].ReplaceInlineRange(startInline + 1, endInline - startInline - 1, replacement);
            } else {
                RtfParagraph start = paragraphs[startParagraph];
                start.ReplaceInlineRange(startInline + 1, start.Inlines.Count - startInline - 1, replacement);
                for (int clearIndex = startParagraph + 1; clearIndex < paragraphIndex; clearIndex++) {
                    paragraphs[clearIndex].ReplaceInlineRange(0, paragraphs[clearIndex].Inlines.Count, null);
                }
                paragraphs[paragraphIndex].ReplaceInlineRange(0, endInline, null);
            }

            return true;
        }

        return false;
    }

    private void InsertIntoSection(int documentIndex, IRtfBlock block) {
        if (_sections.Count == 0) return;
        if (documentIndex < _blocks.Count) {
            IRtfBlock anchor = _blocks[documentIndex];
            foreach (RtfSection section in _sections) {
                int localIndex = section.IndexOfBlock(anchor);
                if (localIndex >= 0) {
                    section.InsertBlock(localIndex, block);
                    return;
                }
            }
        }

        _sections[_sections.Count - 1].InsertBlock(_sections[_sections.Count - 1].Blocks.Count, block);
    }

    private void RebuildParagraphIndex() {
        _paragraphs.Clear();
        foreach (IRtfBlock block in _blocks) if (block is RtfParagraph paragraph) _paragraphs.Add(paragraph);
    }

    private static void ValidateEditableBlock(IRtfBlock block) {
        if (!(block is RtfParagraph) && !(block is RtfTable) && !(block is RtfImage) && !(block is RtfObject) && !(block is RtfShape)) {
            throw new ArgumentException("Unsupported RTF block type.", nameof(block));
        }
    }

    private static void CollectParagraphs(IEnumerable<IRtfBlock> blocks, ISet<RtfParagraph> paragraphs) {
        foreach (IRtfBlock block in blocks) {
            if (block is RtfParagraph paragraph) {
                paragraphs.Add(paragraph);
            } else if (block is RtfTable table) {
                foreach (RtfTableRow row in table.Rows) foreach (RtfTableCell cell in row.Cells) CollectParagraphs(cell.Blocks, paragraphs);
            } else if (block is RtfObject rtfObject) {
                paragraphs.Add(rtfObject.Result);
            } else if (block is RtfShape shape) {
                foreach (RtfParagraph shapeParagraph in shape.TextBoxParagraphs) paragraphs.Add(shapeParagraph);
            }
        }
    }

    private static void CollectParagraphsInOrder(IEnumerable<IRtfBlock> blocks, ICollection<RtfParagraph> paragraphs) {
        foreach (IRtfBlock block in blocks) {
            if (block is RtfParagraph paragraph) {
                paragraphs.Add(paragraph);
            } else if (block is RtfTable table) {
                foreach (RtfTableRow row in table.Rows) foreach (RtfTableCell cell in row.Cells) CollectParagraphsInOrder(cell.Blocks, paragraphs);
            } else if (block is RtfObject rtfObject) {
                paragraphs.Add(rtfObject.Result);
            } else if (block is RtfShape shape) {
                foreach (RtfParagraph shapeParagraph in shape.TextBoxParagraphs) paragraphs.Add(shapeParagraph);
            }
        }
    }
}
