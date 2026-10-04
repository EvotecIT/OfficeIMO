namespace OfficeIMO.AsciiDoc;

public sealed partial class AsciiDocDocument {
    /// <summary>Creates an empty editable AsciiDoc document.</summary>
    public static AsciiDocDocument Create() => Parse(string.Empty);

    /// <summary>Appends parsed AsciiDoc blocks, including their metadata and source trivia.</summary>
    public AsciiDocDocument Add(string source, AsciiDocParseOptions? options = null) {
        Insert(Blocks.Count, source, options);
        return this;
    }

    /// <summary>Inserts parsed AsciiDoc blocks at a zero-based source-block index.</summary>
    public IReadOnlyList<AsciiDocBlock> Insert(int index, string source, AsciiDocParseOptions? options = null) {
        if (index < 0 || index > Blocks.Count) throw new ArgumentOutOfRangeException(nameof(index));
        EnsureInsertionBoundary(index);
        AsciiDocParseResult parsed = ParseResult(source, options);
        if (parsed.HasErrors) throw new ArgumentException("Inserted AsciiDoc must parse without recovery errors. Inspect ParseResult before insertion.", nameof(source));
        return InsertParsed(index, parsed.Document);
    }

    private IReadOnlyList<AsciiDocBlock> InsertParsed(int index, AsciiDocDocument fragment) {
        AsciiDocBlock[] added = fragment.Blocks.ToArray();
        if (added.Length == 0) return added;
        _editableBlocks.InsertRange(index, added);
        _structureWasModified = true;
        return Array.AsReadOnly(added);
    }

    /// <summary>Appends a paragraph with native inline syntax while protecting block-start syntax.</summary>
    /// <exception cref="ArgumentException">The source is empty or contains more than one paragraph.</exception>
    public AsciiDocDocument AddParagraph(string inlineSource) {
        if (inlineSource == null) throw new ArgumentNullException(nameof(inlineSource));
        AsciiDocParseResult parsed = ParseResult(AsciiDocLiteralText.EscapeBlockStarts(inlineSource));
        AsciiDocBlock[] semantic = parsed.Document.Blocks.Where(block => !(block is AsciiDocBlankLine)).ToArray();
        if (parsed.HasErrors || semantic.Length != 1 || !(semantic[0] is AsciiDocParagraph))
            throw new ArgumentException("The source must contain one nonempty paragraph. Use Add for multiple blocks.", nameof(inlineSource));
        InsertParsed(Blocks.Count, parsed.Document);
        return this;
    }

    /// <summary>Appends a section heading. Section levels 1 through 5 use two through six equals signs.</summary>
    public AsciiDocDocument AddHeading(int sectionLevel, string inlineSource) {
        if (sectionLevel < 1 || sectionLevel > 5) throw new ArgumentOutOfRangeException(nameof(sectionLevel));
        if (inlineSource == null) throw new ArgumentNullException(nameof(inlineSource));
        AsciiDocText.EnsureSingleLine(inlineSource, nameof(inlineSource));
        return Add(new string('=', sectionLevel + 1) + " " + inlineSource);
    }

    /// <summary>Removes a semantic block together with its bound metadata and list attachments.</summary>
    public bool Remove(AsciiDocBlock block) {
        if (block == null) throw new ArgumentNullException(nameof(block));
        if (!_editableBlocks.Contains(block)) return false;
        EnsureSemanticUnit(block);
        HashSet<AsciiDocBlock> unit = FindUnit(block);
        DetachOutsideAttachments(unit);
        _editableBlocks.RemoveAll(unit.Contains);
        _structureWasModified = true;
        return true;
    }

    /// <summary>Moves a semantic block and its bound metadata and attachments before the original zero-based index.</summary>
    public AsciiDocDocument Move(AsciiDocBlock block, int index) {
        if (block == null) throw new ArgumentNullException(nameof(block));
        if (index < 0 || index > Blocks.Count) throw new ArgumentOutOfRangeException(nameof(index));
        if (!_editableBlocks.Contains(block)) throw new ArgumentException("The block does not belong to this document.", nameof(block));
        EnsureSemanticUnit(block);
        HashSet<AsciiDocBlock> unit = FindUnit(block);
        AsciiDocBlock[] moving = _editableBlocks.Where(unit.Contains).ToArray();
        int first = _editableBlocks.FindIndex(unit.Contains);
        int last = _editableBlocks.FindLastIndex(unit.Contains);
        if (index >= first && index <= last + 1) return this;
        EnsureInsertionBoundary(index, unit);
        int destination = index - _editableBlocks.Take(index).Count(unit.Contains);
        DetachOutsideAttachments(unit);
        _editableBlocks.RemoveAll(unit.Contains);
        _editableBlocks.InsertRange(destination, moving);
        _structureWasModified = true;
        return this;
    }

    private static void EnsureSemanticUnit(AsciiDocBlock block) {
        if (block is IAsciiDocBlockMetadata metadata && metadata.Target != null || block is AsciiDocListContinuation)
            throw new ArgumentException("Move or remove the associated semantic block so its bound metadata and attachments remain together.", nameof(block));
    }

    private HashSet<AsciiDocBlock> FindUnit(AsciiDocBlock block) {
        var unit = new HashSet<AsciiDocBlock> { block };
        bool changed;
        do {
            changed = false;
            foreach (AsciiDocBlock candidate in Blocks) {
                if (candidate is IAsciiDocBlockMetadata metadata && metadata.Target != null && unit.Contains(metadata.Target)) changed |= unit.Add(candidate);
                if (candidate is AsciiDocListContinuation continuation && continuation.AttachedBlock != null && unit.Contains(continuation.AttachedBlock)) changed |= unit.Add(candidate);
                if (candidate is AsciiDocListBlock list && unit.Contains(list))
                    foreach (AsciiDocBlock attached in list.Items.SelectMany(static item => item.AttachedBlocks)) changed |= unit.Add(attached);
            }
        } while (changed);
        return unit;
    }

    private void EnsureInsertionBoundary(int index, HashSet<AsciiDocBlock>? ignored = null) {
        foreach (AsciiDocBlock block in Blocks) {
            if (ignored?.Contains(block) == true) continue;
            if (block is IAsciiDocBlockMetadata metadata && metadata.Target != null) {
                int start = _editableBlocks.IndexOf(block);
                int end = _editableBlocks.IndexOf(metadata.Target);
                if (start < index && index <= end) throw new ArgumentException("Insertion cannot split a block from its bound metadata.", nameof(index));
            }
            if (block is AsciiDocListBlock list && list.Items.Any(static item => item.AttachedBlocks.Count > 0)) {
                HashSet<AsciiDocBlock> unit = FindUnit(block);
                int first = _editableBlocks.FindIndex(unit.Contains);
                int last = _editableBlocks.FindLastIndex(unit.Contains);
                if (first < index && index <= last) throw new ArgumentException("Insertion cannot split a list from its continuation attachments.", nameof(index));
            }
        }
    }

    private void DetachOutsideAttachments(HashSet<AsciiDocBlock> unit) {
        foreach (AsciiDocListBlock list in BlocksOfType<AsciiDocListBlock>().Where(list => !unit.Contains(list)))
            foreach (AsciiDocListItem item in list.Items)
                foreach (AsciiDocBlock attached in item.AttachedBlocks.Where(unit.Contains).ToArray()) item.RemoveAttachedBlock(attached);
    }
}
