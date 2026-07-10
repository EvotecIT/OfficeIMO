namespace OfficeIMO.AsciiDoc;

/// <summary>Dependency-free, lossless AsciiDoc parser.</summary>
public static class AsciiDocParser {
    /// <summary>Parses source into a lossless syntax tree and typed semantic blocks.</summary>
    public static AsciiDocParseResult Parse(string source, AsciiDocParseOptions? options = null) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        options ??= new AsciiDocParseOptions();
        ValidateOptions(source, options);

        var sourceText = new AsciiDocSourceText(source);
        IReadOnlyList<AsciiDocSourceLine> lines = AsciiDocLineReader.Read(source);
        var factory = new AsciiDocSyntaxFactory(sourceText);
        var inlineParser = new AsciiDocInlineParser(factory, options);
        var blocks = new List<AsciiDocBlock>();
        var syntaxNodes = new List<AsciiDocSyntaxNode>();
        var diagnostics = new List<AsciiDocDiagnostic>();
        bool hasStructuralContent = false;

        int lineIndex = 0;
        while (lineIndex < lines.Count) {
            EnforceBlockLimit(blocks.Count, options);
            AsciiDocSourceLine line = lines[lineIndex];
            string content = line.Content;

            if (AsciiDocLineClassifier.IsBlank(content)) {
                AddBlank(line, factory, blocks, syntaxNodes);
                lineIndex++;
                continue;
            }

            if (AsciiDocLineClassifier.TryGetDelimiter(content, out AsciiDocDelimitedBlockKind delimiterKind)) {
                AsciiDocTableConfiguration? tableConfiguration = delimiterKind == AsciiDocDelimitedBlockKind.Table
                    ? AsciiDocTableConfiguration.Create(content, GetPendingAttributeLists(blocks))
                    : null;
                lineIndex = AddDelimitedBlock(lines, lineIndex, delimiterKind, sourceText, factory, tableConfiguration, blocks, syntaxNodes, diagnostics);
                hasStructuralContent = true;
                continue;
            }

            if (AsciiDocLineClassifier.IsLineComment(content)) {
                AddLineComment(line, factory, blocks, syntaxNodes);
                lineIndex++;
                continue;
            }

            if (AsciiDocLineClassifier.TryParseAttribute(content, out AsciiDocLineClassifier.AttributeParts attribute)) {
                AddAttribute(line, attribute, factory, blocks, syntaxNodes);
                lineIndex++;
                continue;
            }

            if (AsciiDocLineClassifier.TryParseBlockAttributeList(content, out string attributeList)) {
                AddBlockAttributeList(line, attributeList, factory, blocks, syntaxNodes);
                lineIndex++;
                continue;
            }

            if (AsciiDocLineClassifier.TryParseBlockTitle(content, out string blockTitle)) {
                AddBlockTitle(line, blockTitle, factory, inlineParser, blocks, syntaxNodes);
                lineIndex++;
                continue;
            }

            if (AsciiDocLineClassifier.TryParseBlockAnchor(content, out string anchorId, out string? referenceText)) {
                AddBlockAnchor(line, anchorId, referenceText, factory, blocks, syntaxNodes);
                lineIndex++;
                continue;
            }

            if (AsciiDocLineClassifier.TryParseHeading(content, out int markerLength, out int titleStart)) {
                bool isDocumentTitle = markerLength == 1 && !hasStructuralContent;
                AddHeading(line, markerLength, titleStart, isDocumentTitle, factory, inlineParser, blocks, syntaxNodes);
                hasStructuralContent = true;
                lineIndex++;
                continue;
            }

            if (AsciiDocLineClassifier.TryParseAdmonition(content, out AsciiDocLineClassifier.AdmonitionParts admonition)) {
                AddAdmonition(line, admonition, factory, inlineParser, blocks, syntaxNodes);
                hasStructuralContent = true;
                lineIndex++;
                continue;
            }

            if (AsciiDocLineClassifier.TryParseDescriptionListItem(content, out _)) {
                lineIndex = AddDescriptionList(lines, lineIndex, factory, inlineParser, blocks, syntaxNodes);
                hasStructuralContent = true;
                continue;
            }

            if (AsciiDocLineClassifier.TryParseListItem(content, out AsciiDocLineClassifier.ListItemParts listItem)) {
                lineIndex = AddList(lines, lineIndex, listItem.Kind, factory, inlineParser, blocks, syntaxNodes);
                hasStructuralContent = true;
                continue;
            }

            if (AsciiDocLineClassifier.IsListContinuation(content)) {
                AddListContinuation(line, factory, blocks, syntaxNodes);
                lineIndex++;
                continue;
            }

            if (AsciiDocLineClassifier.TryParseBlockMacro(content, out AsciiDocLineClassifier.BlockMacroParts macro)) {
                AddBlockMacro(line, macro, factory, blocks, syntaxNodes);
                hasStructuralContent = true;
                lineIndex++;
                continue;
            }

            lineIndex = AddParagraph(lines, lineIndex, factory, inlineParser, blocks, syntaxNodes);
            hasStructuralContent = true;
        }

        AsciiDocSyntaxNode root = factory.Node(AsciiDocSyntaxKind.Document, 0, source.Length, syntaxNodes);
        var syntaxTree = new AsciiDocSyntaxTree(sourceText, root);
        if (!syntaxTree.IsLossless) {
            diagnostics.Add(new AsciiDocDiagnostic(
                "ADOC900",
                AsciiDocDiagnosticSeverity.Error,
                "The parser did not retain a contiguous representation of the complete source.",
                root.Span));
        }

        BindBlockMetadata(blocks);
        BindListContinuations(blocks);
        var document = new AsciiDocDocument(sourceText, syntaxTree, blocks, diagnostics);
        return new AsciiDocParseResult(document, diagnostics);
    }

    private static void AddBlank(
        AsciiDocSourceLine line,
        AsciiDocSyntaxFactory factory,
        List<AsciiDocBlock> blocks,
        List<AsciiDocSyntaxNode> syntaxNodes) {
        AsciiDocSyntaxNode syntax = factory.Node(AsciiDocSyntaxKind.BlankLine, line.Start, line.End);
        syntaxNodes.Add(syntax);
        blocks.Add(new AsciiDocBlankLine(syntax, line.LineEnding));
    }

    private static void AddLineComment(
        AsciiDocSourceLine line,
        AsciiDocSyntaxFactory factory,
        List<AsciiDocBlock> blocks,
        List<AsciiDocSyntaxNode> syntaxNodes) {
        int textStart = line.Start + 2;
        if (textStart < line.ContentEnd && factory.Source.Text[textStart] == ' ') textStart++;
        var children = new List<AsciiDocSyntaxNode> {
            factory.Node(AsciiDocSyntaxKind.Text, line.Start, line.Start + 2)
        };
        if (textStart < line.ContentEnd) children.Add(factory.Node(AsciiDocSyntaxKind.Text, textStart, line.ContentEnd));
        factory.AddLineEnding(children, line);
        AsciiDocSyntaxNode syntax = factory.Node(AsciiDocSyntaxKind.CommentLine, line.Start, line.End, children);
        syntaxNodes.Add(syntax);
        string text = textStart < line.ContentEnd ? factory.Source.Text.Substring(textStart, line.ContentEnd - textStart) : string.Empty;
        blocks.Add(new AsciiDocLineComment(syntax, text, line.LineEnding));
    }

    private static void AddAttribute(
        AsciiDocSourceLine line,
        AsciiDocLineClassifier.AttributeParts parts,
        AsciiDocSyntaxFactory factory,
        List<AsciiDocBlock> blocks,
        List<AsciiDocSyntaxNode> syntaxNodes) {
        var children = new List<AsciiDocSyntaxNode> {
            factory.Node(AsciiDocSyntaxKind.AttributeMarker, line.Start, line.Start + 1),
            factory.Node(AsciiDocSyntaxKind.AttributeName, line.Start + parts.NameStart, line.Start + parts.NameStart + parts.NameLength),
            factory.Node(AsciiDocSyntaxKind.AttributeMarker, line.Start + parts.Separator, line.Start + parts.Separator + 1)
        };
        if (parts.ValueStart < line.ContentLength) {
            children.Add(factory.Node(AsciiDocSyntaxKind.AttributeValue, line.Start + parts.ValueStart, line.ContentEnd));
        }
        factory.AddLineEnding(children, line);
        AsciiDocSyntaxNode syntax = factory.Node(AsciiDocSyntaxKind.AttributeEntry, line.Start, line.End, children);
        syntaxNodes.Add(syntax);
        blocks.Add(new AsciiDocAttributeEntry(syntax, parts.Name, parts.Value, parts.IsUnset, line.LineEnding));
    }

    private static void AddHeading(
        AsciiDocSourceLine line,
        int markerLength,
        int titleStart,
        bool isDocumentTitle,
        AsciiDocSyntaxFactory factory,
        AsciiDocInlineParser inlineParser,
        List<AsciiDocBlock> blocks,
        List<AsciiDocSyntaxNode> syntaxNodes) {
        var children = new List<AsciiDocSyntaxNode> {
            factory.Node(AsciiDocSyntaxKind.HeadingMarker, line.Start, line.Start + markerLength)
        };
        AsciiDocInlineSequence inlines = inlineParser.Parse(line.Start + titleStart, line.ContentEnd);
        if (titleStart < line.ContentLength) children.Add(inlines.Syntax);
        factory.AddLineEnding(children, line);
        AsciiDocSyntaxNode syntax = factory.Node(AsciiDocSyntaxKind.Heading, line.Start, line.End, children);
        syntaxNodes.Add(syntax);
        string marker = line.Content.Substring(0, markerLength);
        string title = line.Content.Substring(titleStart);
        blocks.Add(new AsciiDocHeading(syntax, marker, title, isDocumentTitle, inlines, line.LineEnding));
    }

    private static int AddList(
        IReadOnlyList<AsciiDocSourceLine> lines,
        int startIndex,
        AsciiDocListKind kind,
        AsciiDocSyntaxFactory factory,
        AsciiDocInlineParser inlineParser,
        List<AsciiDocBlock> blocks,
        List<AsciiDocSyntaxNode> syntaxNodes) {
        var items = new List<AsciiDocListItem>();
        var itemSyntax = new List<AsciiDocSyntaxNode>();
        int index = startIndex;
        while (index < lines.Count &&
               AsciiDocLineClassifier.TryParseListItem(lines[index].Content, out AsciiDocLineClassifier.ListItemParts parts) &&
               parts.Kind == kind) {
            AsciiDocSourceLine line = lines[index];
            var children = new List<AsciiDocSyntaxNode> {
                factory.Node(AsciiDocSyntaxKind.ListMarker, line.Start, line.Start + parts.MarkerLength)
            };
            int textStart = line.Start + parts.MarkerLength + 1;
            AsciiDocInlineSequence inlines = inlineParser.Parse(textStart, line.ContentEnd);
            if (textStart < line.ContentEnd) children.Add(inlines.Syntax);
            factory.AddLineEnding(children, line);
            AsciiDocSyntaxNode syntax = factory.Node(AsciiDocSyntaxKind.ListItem, line.Start, line.End, children);
            itemSyntax.Add(syntax);
            items.Add(new AsciiDocListItem(syntax, kind, parts.Marker, parts.MarkerLength, parts.Text, inlines, line.LineEnding));
            index++;
        }

        AsciiDocSourceLine first = lines[startIndex];
        AsciiDocSourceLine last = lines[index - 1];
        AsciiDocSyntaxKind syntaxKind = kind == AsciiDocListKind.Ordered ? AsciiDocSyntaxKind.OrderedList : AsciiDocSyntaxKind.UnorderedList;
        AsciiDocSyntaxNode blockSyntax = factory.Node(syntaxKind, first.Start, last.End, itemSyntax);
        syntaxNodes.Add(blockSyntax);
        blocks.Add(new AsciiDocListBlock(blockSyntax, kind, items, last.LineEnding));
        return index;
    }

    private static int AddDelimitedBlock(
        IReadOnlyList<AsciiDocSourceLine> lines,
        int startIndex,
        AsciiDocDelimitedBlockKind kind,
        AsciiDocSourceText source,
        AsciiDocSyntaxFactory factory,
        AsciiDocTableConfiguration? tableConfiguration,
        List<AsciiDocBlock> blocks,
        List<AsciiDocSyntaxNode> syntaxNodes,
        List<AsciiDocDiagnostic> diagnostics) {
        AsciiDocSourceLine opening = lines[startIndex];
        string delimiter = opening.Content;
        int closingIndex = -1;
        for (int index = startIndex + 1; index < lines.Count; index++) {
            if (string.Equals(lines[index].Content, delimiter, StringComparison.Ordinal)) {
                closingIndex = index;
                break;
            }
        }

        bool isTerminated = closingIndex >= 0;
        int endIndex = isTerminated ? closingIndex : lines.Count - 1;
        AsciiDocSourceLine ending = lines[endIndex];
        int contentStart = opening.End;
        int contentEnd = isTerminated ? lines[closingIndex].Start : ending.End;
        var children = new List<AsciiDocSyntaxNode> {
            factory.Node(AsciiDocSyntaxKind.BlockDelimiter, opening.Start, opening.ContentEnd)
        };
        factory.AddLineEnding(children, opening);
        AsciiDocTableParseResult? table = null;
        if (kind == AsciiDocDelimitedBlockKind.Table) {
            table = AsciiDocTableParser.Parse(
                factory,
                contentStart,
                contentEnd,
                tableConfiguration ?? new AsciiDocTableConfiguration(AsciiDocTableFormat.Psv, "|", null, false));
            children.Add(table.Syntax);
        } else if (contentEnd > contentStart) {
            children.Add(factory.Node(AsciiDocSyntaxKind.BlockContent, contentStart, contentEnd));
        }
        if (isTerminated) {
            AsciiDocSourceLine closing = lines[closingIndex];
            children.Add(factory.Node(AsciiDocSyntaxKind.BlockDelimiter, closing.Start, closing.ContentEnd));
            factory.AddLineEnding(children, closing);
        }

        AsciiDocSyntaxKind syntaxKind = kind == AsciiDocDelimitedBlockKind.Comment
            ? AsciiDocSyntaxKind.CommentBlock
            : kind == AsciiDocDelimitedBlockKind.Table
                ? AsciiDocSyntaxKind.Table
                : AsciiDocSyntaxKind.DelimitedBlock;
        AsciiDocSyntaxNode syntax = factory.Node(syntaxKind, opening.Start, ending.End, children);
        syntaxNodes.Add(syntax);

        string content = contentEnd > contentStart ? source.Text.Substring(contentStart, contentEnd - contentStart) : string.Empty;
        string closingText = isTerminated ? lines[closingIndex].FullText : string.Empty;
        if (kind == AsciiDocDelimitedBlockKind.Table && table != null) {
            blocks.Add(new AsciiDocTableBlock(
                syntax,
                delimiter,
                opening.FullText,
                content,
                closingText,
                isTerminated,
                ending.LineEnding,
                table.Table));
        } else {
            blocks.Add(new AsciiDocDelimitedBlock(
                syntax,
                kind,
                delimiter,
                opening.FullText,
                content,
                closingText,
                isTerminated,
                ending.LineEnding));
        }

        if (!isTerminated) {
            diagnostics.Add(new AsciiDocDiagnostic(
                "ADOC001",
                AsciiDocDiagnosticSeverity.Error,
                "Delimited block '" + delimiter + "' is not terminated; source was preserved through end of input.",
                source.CreateSpan(opening.Start, opening.ContentEnd)));
        }

        return endIndex + 1;
    }

    private static void AddBlockMacro(
        AsciiDocSourceLine line,
        AsciiDocLineClassifier.BlockMacroParts parts,
        AsciiDocSyntaxFactory factory,
        List<AsciiDocBlock> blocks,
        List<AsciiDocSyntaxNode> syntaxNodes) {
        int separatorStart = line.Start + parts.Separator;
        int attributesStart = line.Start + parts.AttributesOpen;
        var children = new List<AsciiDocSyntaxNode> {
            factory.Node(AsciiDocSyntaxKind.MacroName, line.Start, separatorStart),
            factory.Node(AsciiDocSyntaxKind.MacroSeparator, separatorStart, separatorStart + 2)
        };
        if (attributesStart > separatorStart + 2) children.Add(factory.Node(AsciiDocSyntaxKind.MacroTarget, separatorStart + 2, attributesStart));
        children.Add(factory.Node(AsciiDocSyntaxKind.MacroAttributeList, attributesStart, line.ContentEnd));
        factory.AddLineEnding(children, line);
        AsciiDocSyntaxNode syntax = factory.Node(AsciiDocSyntaxKind.BlockMacro, line.Start, line.End, children);
        syntaxNodes.Add(syntax);
        blocks.Add(new AsciiDocBlockMacro(syntax, parts.Name, parts.Target, parts.AttributeList, line.LineEnding));
    }

    private static int AddParagraph(
        IReadOnlyList<AsciiDocSourceLine> lines,
        int startIndex,
        AsciiDocSyntaxFactory factory,
        AsciiDocInlineParser inlineParser,
        List<AsciiDocBlock> blocks,
        List<AsciiDocSyntaxNode> syntaxNodes) {
        int index = startIndex + 1;
        while (index < lines.Count && !AsciiDocLineClassifier.IsBlockStart(lines[index].Content)) index++;

        AsciiDocSourceLine first = lines[startIndex];
        AsciiDocSourceLine last = lines[index - 1];
        var text = new StringBuilder();
        for (int lineIndex = startIndex; lineIndex < index; lineIndex++) {
            AsciiDocSourceLine line = lines[lineIndex];
            if (lineIndex > startIndex) text.Append('\n');
            text.Append(line.Content);
        }

        var children = new List<AsciiDocSyntaxNode>();
        AsciiDocInlineSequence inlines = inlineParser.Parse(first.Start, last.ContentEnd);
        if (last.ContentEnd > first.Start) children.Add(inlines.Syntax);
        factory.AddLineEnding(children, last);

        AsciiDocSyntaxNode syntax = factory.Node(AsciiDocSyntaxKind.Paragraph, first.Start, last.End, children);
        syntaxNodes.Add(syntax);
        blocks.Add(new AsciiDocParagraph(syntax, text.ToString(), inlines, last.LineEnding));
        return index;
    }

    private static void AddBlockAttributeList(
        AsciiDocSourceLine line,
        string content,
        AsciiDocSyntaxFactory factory,
        List<AsciiDocBlock> blocks,
        List<AsciiDocSyntaxNode> syntaxNodes) {
        var children = new List<AsciiDocSyntaxNode>();
        if (line.ContentLength > 2) {
            children.Add(factory.Node(AsciiDocSyntaxKind.BlockAttributeListContent, line.Start + 1, line.ContentEnd - 1));
        }
        factory.AddLineEnding(children, line);
        AsciiDocSyntaxNode syntax = factory.Node(AsciiDocSyntaxKind.BlockAttributeList, line.Start, line.End, children);
        syntaxNodes.Add(syntax);
        blocks.Add(new AsciiDocBlockAttributeList(syntax, content, line.LineEnding));
    }

    private static void AddBlockTitle(
        AsciiDocSourceLine line,
        string title,
        AsciiDocSyntaxFactory factory,
        AsciiDocInlineParser inlineParser,
        List<AsciiDocBlock> blocks,
        List<AsciiDocSyntaxNode> syntaxNodes) {
        AsciiDocInlineSequence inlines = inlineParser.Parse(line.Start + 1, line.ContentEnd);
        var children = new List<AsciiDocSyntaxNode> { inlines.Syntax };
        factory.AddLineEnding(children, line);
        AsciiDocSyntaxNode syntax = factory.Node(AsciiDocSyntaxKind.BlockTitle, line.Start, line.End, children);
        syntaxNodes.Add(syntax);
        blocks.Add(new AsciiDocBlockTitle(syntax, title, inlines, line.LineEnding));
    }

    private static void AddBlockAnchor(
        AsciiDocSourceLine line,
        string id,
        string? referenceText,
        AsciiDocSyntaxFactory factory,
        List<AsciiDocBlock> blocks,
        List<AsciiDocSyntaxNode> syntaxNodes) {
        var children = new List<AsciiDocSyntaxNode>();
        factory.AddLineEnding(children, line);
        AsciiDocSyntaxNode syntax = factory.Node(AsciiDocSyntaxKind.BlockAnchor, line.Start, line.End, children);
        syntaxNodes.Add(syntax);
        blocks.Add(new AsciiDocBlockAnchor(syntax, id, referenceText, line.LineEnding));
    }

    private static void AddAdmonition(
        AsciiDocSourceLine line,
        AsciiDocLineClassifier.AdmonitionParts parts,
        AsciiDocSyntaxFactory factory,
        AsciiDocInlineParser inlineParser,
        List<AsciiDocBlock> blocks,
        List<AsciiDocSyntaxNode> syntaxNodes) {
        int markerEnd = line.Start + parts.Label.Length + 1;
        var children = new List<AsciiDocSyntaxNode> {
            factory.Node(AsciiDocSyntaxKind.AdmonitionMarker, line.Start, markerEnd)
        };
        AsciiDocInlineSequence inlines = inlineParser.Parse(line.Start + parts.TextStart, line.ContentEnd);
        if (parts.TextStart < line.ContentLength) children.Add(inlines.Syntax);
        factory.AddLineEnding(children, line);
        AsciiDocSyntaxNode syntax = factory.Node(AsciiDocSyntaxKind.Admonition, line.Start, line.End, children);
        syntaxNodes.Add(syntax);
        blocks.Add(new AsciiDocAdmonitionBlock(syntax, parts.Kind, parts.Label, parts.Text, inlines, line.LineEnding));
    }

    private static int AddDescriptionList(
        IReadOnlyList<AsciiDocSourceLine> lines,
        int startIndex,
        AsciiDocSyntaxFactory factory,
        AsciiDocInlineParser inlineParser,
        List<AsciiDocBlock> blocks,
        List<AsciiDocSyntaxNode> syntaxNodes) {
        var items = new List<AsciiDocDescriptionListItem>();
        var itemSyntax = new List<AsciiDocSyntaxNode>();
        int index = startIndex;
        while (index < lines.Count &&
               AsciiDocLineClassifier.TryParseDescriptionListItem(lines[index].Content, out AsciiDocLineClassifier.DescriptionListParts parts)) {
            AsciiDocSourceLine line = lines[index];
            AsciiDocInlineSequence term = inlineParser.Parse(line.Start, line.Start + parts.MarkerStart);
            AsciiDocInlineSequence description = inlineParser.Parse(line.Start + parts.DescriptionStart, line.ContentEnd);
            var children = new List<AsciiDocSyntaxNode> {
                term.Syntax,
                factory.Node(AsciiDocSyntaxKind.DescriptionListMarker,
                    line.Start + parts.MarkerStart,
                    line.Start + parts.MarkerStart + parts.MarkerLength)
            };
            if (parts.DescriptionStart < line.ContentLength) children.Add(description.Syntax);
            factory.AddLineEnding(children, line);
            AsciiDocSyntaxNode syntax = factory.Node(AsciiDocSyntaxKind.DescriptionListItem, line.Start, line.End, children);
            itemSyntax.Add(syntax);
            items.Add(new AsciiDocDescriptionListItem(
                syntax,
                parts.Marker,
                parts.Term,
                parts.Description,
                term,
                description,
                line.LineEnding));
            index++;
        }
        AsciiDocSourceLine first = lines[startIndex];
        AsciiDocSourceLine last = lines[index - 1];
        AsciiDocSyntaxNode blockSyntax = factory.Node(AsciiDocSyntaxKind.DescriptionList, first.Start, last.End, itemSyntax);
        syntaxNodes.Add(blockSyntax);
        blocks.Add(new AsciiDocDescriptionListBlock(blockSyntax, items, last.LineEnding));
        return index;
    }

    private static void AddListContinuation(
        AsciiDocSourceLine line,
        AsciiDocSyntaxFactory factory,
        List<AsciiDocBlock> blocks,
        List<AsciiDocSyntaxNode> syntaxNodes) {
        var children = new List<AsciiDocSyntaxNode>();
        factory.AddLineEnding(children, line);
        AsciiDocSyntaxNode syntax = factory.Node(AsciiDocSyntaxKind.ListContinuation, line.Start, line.End, children);
        syntaxNodes.Add(syntax);
        blocks.Add(new AsciiDocListContinuation(syntax, line.LineEnding));
    }

    private static void BindBlockMetadata(IReadOnlyList<AsciiDocBlock> blocks) {
        var pending = new List<IAsciiDocBlockMetadata>();
        for (int index = 0; index < blocks.Count; index++) {
            AsciiDocBlock block = blocks[index];
            if (block is IAsciiDocBlockMetadata metadata) {
                pending.Add(metadata);
                continue;
            }
            if (block is AsciiDocBlankLine || block is AsciiDocLineComment || block is AsciiDocAttributeEntry) {
                pending.Clear();
                continue;
            }
            for (int pendingIndex = 0; pendingIndex < pending.Count; pendingIndex++) {
                IAsciiDocBlockMetadata source = pending[pendingIndex];
                source.Target = block;
                if (source is AsciiDocBlockAttributeList attributeList) block.AddAttributeList(attributeList);
                else if (source is AsciiDocBlockTitle title) block.SetBlockTitle(title);
                else if (source is AsciiDocBlockAnchor anchor) block.SetBlockAnchor(anchor);
            }
            pending.Clear();
        }
    }

    private static IReadOnlyList<AsciiDocBlockAttributeList> GetPendingAttributeLists(IReadOnlyList<AsciiDocBlock> blocks) {
        var result = new List<AsciiDocBlockAttributeList>();
        for (int index = blocks.Count - 1; index >= 0; index--) {
            AsciiDocBlock block = blocks[index];
            if (block is AsciiDocBlockAttributeList attributeList) {
                result.Insert(0, attributeList);
                continue;
            }
            if (block is AsciiDocBlockTitle || block is AsciiDocBlockAnchor) continue;
            break;
        }
        return result;
    }

    private static void BindListContinuations(IReadOnlyList<AsciiDocBlock> blocks) {
        var attached = new HashSet<AsciiDocBlock>();
        for (int index = 0; index < blocks.Count; index++) {
            if (blocks[index] is not AsciiDocListContinuation continuation) continue;
            AsciiDocListItem? item = FindContinuationItem(blocks, index, attached);
            AsciiDocBlock? target = FindContinuationTarget(blocks, index + 1);
            if (item == null || target == null) continue;
            continuation.TargetItem = item;
            continuation.AttachedBlock = target;
            item.AddAttachedBlock(target);
            attached.Add(target);
        }
    }

    private static AsciiDocListItem? FindContinuationItem(
        IReadOnlyList<AsciiDocBlock> blocks,
        int continuationIndex,
        HashSet<AsciiDocBlock> attached) {
        for (int index = continuationIndex - 1; index >= 0; index--) {
            AsciiDocBlock candidate = blocks[index];
            if (candidate is AsciiDocListBlock list) return list.Items.Count == 0 ? null : list.Items[list.Items.Count - 1];
            if (candidate is IAsciiDocBlockMetadata || candidate is AsciiDocListContinuation || attached.Contains(candidate)) continue;
            return null;
        }
        return null;
    }

    private static AsciiDocBlock? FindContinuationTarget(IReadOnlyList<AsciiDocBlock> blocks, int startIndex) {
        for (int index = startIndex; index < blocks.Count; index++) {
            AsciiDocBlock candidate = blocks[index];
            if (candidate is IAsciiDocBlockMetadata) continue;
            if (candidate is AsciiDocBlankLine || candidate is AsciiDocListContinuation) return null;
            return candidate;
        }
        return null;
    }

    private static void ValidateOptions(string source, AsciiDocParseOptions options) {
        if (options.MaximumInputLength.HasValue && options.MaximumInputLength.Value < 0) {
            throw new ArgumentOutOfRangeException(nameof(options), "MaximumInputLength cannot be negative.");
        }
        if (options.MaximumBlockCount.HasValue && options.MaximumBlockCount.Value < 1) {
            throw new ArgumentOutOfRangeException(nameof(options), "MaximumBlockCount must be positive.");
        }
        if (options.MaximumInlineNestingDepth < 1) {
            throw new ArgumentOutOfRangeException(nameof(options), "MaximumInlineNestingDepth must be positive.");
        }
        if (options.MaximumInlineNodeCount < 1) {
            throw new ArgumentOutOfRangeException(nameof(options), "MaximumInlineNodeCount must be positive.");
        }
        if (options.MaximumInputLength.HasValue && source.Length > options.MaximumInputLength.Value) {
            throw new ArgumentException("AsciiDoc source exceeds MaximumInputLength.", nameof(source));
        }
    }

    private static void EnforceBlockLimit(int blockCount, AsciiDocParseOptions options) {
        if (options.MaximumBlockCount.HasValue && blockCount >= options.MaximumBlockCount.Value) {
            throw new InvalidDataException("AsciiDoc source exceeds MaximumBlockCount.");
        }
    }

}
