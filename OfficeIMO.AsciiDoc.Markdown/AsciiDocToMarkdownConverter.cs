namespace OfficeIMO.AsciiDoc.Markdown;

/// <summary>Loss-aware AsciiDoc to Markdown conversion engine.</summary>
internal static class AsciiDocToMarkdownConverter {
    /// <summary>Converts recognized AsciiDoc semantics and reports every fallback or omission.</summary>
    internal static AsciiDocToMarkdownResult Convert(
        AsciiDocDocument document,
        AsciiDocToMarkdownOptions? options = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        options ??= new AsciiDocToMarkdownOptions();

        var markdown = MarkdownDoc.Create();
        var diagnostics = new List<AsciiDocMarkdownConversionDiagnostic>();
        AsciiDocDocumentAttributes attributes = document.GetAttributes();
        var attachedBlocks = new HashSet<AsciiDocBlock>(
            document.BlocksOfType<AsciiDocListBlock>()
                .SelectMany(static list => list.Items)
                .SelectMany(static item => item.AttachedBlocks));
        AddFrontMatter(document, markdown, options, diagnostics);

        for (int index = 0; index < document.Blocks.Count; index++) {
            AsciiDocBlock block = document.Blocks[index];
            if (attachedBlocks.Contains(block)) continue;
            AddBlock(markdown, block, attributes, options, diagnostics);
        }

        return new AsciiDocToMarkdownResult(markdown, diagnostics);
    }

    internal static AsciiDocToMarkdownResult ConvertBlock(
        AsciiDocBlock block,
        AsciiDocDocumentAttributes attributes,
        AsciiDocToMarkdownOptions? options = null) {
        if (block == null) throw new ArgumentNullException(nameof(block));
        if (attributes == null) throw new ArgumentNullException(nameof(attributes));
        options ??= new AsciiDocToMarkdownOptions();
        var markdown = MarkdownDoc.Create();
        var diagnostics = new List<AsciiDocMarkdownConversionDiagnostic>();
        AddBlock(markdown, block, attributes, options, diagnostics);
        return new AsciiDocToMarkdownResult(markdown, diagnostics);
    }

    private static void AddBlock(
        MarkdownDoc markdown,
        AsciiDocBlock block,
        AsciiDocDocumentAttributes attributes,
        AsciiDocToMarkdownOptions options,
        List<AsciiDocMarkdownConversionDiagnostic> diagnostics) {
        switch (block) {
            case AsciiDocBlankLine:
            case AsciiDocAttributeEntry:
                break;
            case IAsciiDocBlockMetadata metadata when metadata.Target != null:
                break;
            case AsciiDocHeading heading:
                int level = heading.IsDocumentTitle ? 1 : Math.Max(1, Math.Min(6, heading.SectionLevel + 1));
                var markdownHeading = new HeadingBlock(level,
                    AsciiDocInlineToMarkdownConverter.Convert(heading.Inlines, attributes, options, diagnostics, heading));
                ApplyMetadata(markdownHeading, heading);
                markdown.Add(markdownHeading);
                break;
            case AsciiDocParagraph paragraph:
                var markdownParagraph = new ParagraphBlock(
                    AsciiDocInlineToMarkdownConverter.Convert(paragraph.Inlines, attributes, options, diagnostics, paragraph));
                ApplyMetadata(markdownParagraph, paragraph);
                markdown.Add(markdownParagraph);
                break;
            case AsciiDocListBlock list:
                AddList(markdown, list, attributes, options, diagnostics);
                break;
            case AsciiDocDescriptionListBlock descriptionList:
                AddDescriptionList(markdown, descriptionList, attributes, options, diagnostics);
                break;
            case AsciiDocAdmonitionBlock admonition:
                AddAdmonition(markdown, admonition, attributes, options, diagnostics);
                break;
            case AsciiDocTableBlock table:
                TableBlock markdownTable = AsciiDocTableToMarkdownConverter.Convert(table, attributes, options, diagnostics);
                ApplyMetadata(markdownTable, table);
                markdown.Add(markdownTable);
                break;
            case AsciiDocDelimitedBlock delimited:
                AddDelimitedBlock(markdown, delimited, attributes, options, diagnostics);
                break;
            case AsciiDocBlockMacro macro:
                AddMacro(markdown, macro, options, diagnostics);
                break;
            case AsciiDocListContinuation continuation when continuation.TargetItem != null && continuation.AttachedBlock != null:
                break;
            case AsciiDocLineComment comment:
                AddComment(markdown, comment, options, diagnostics);
                break;
            default:
                AddSourceFallback(markdown, block, "ADOCMD099", block.GetType().Name, options, diagnostics);
                break;
        }
    }

    private static void AddFrontMatter(
        AsciiDocDocument source,
        MarkdownDoc target,
        AsciiDocToMarkdownOptions options,
        List<AsciiDocMarkdownConversionDiagnostic> diagnostics) {
        AsciiDocAttributeEntry[] attributes = source.BlocksOfType<AsciiDocAttributeEntry>().ToArray();
        if (attributes.Length == 0) return;

        if (!options.IncludeDocumentAttributesAsFrontMatter) {
            for (int index = 0; index < attributes.Length; index++) {
                Report(diagnostics, "ADOCMD001", AsciiDocMarkdownDiagnosticSeverity.Info, AsciiDocMarkdownConversionOutcome.Omitted,
                    "document-attribute", "Document attribute omitted by conversion options.", attributes[index]);
            }
            return;
        }

        var values = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
        for (int index = 0; index < attributes.Length; index++) {
            AsciiDocAttributeEntry attribute = attributes[index];
            if (attribute.IsUnset) {
                Report(diagnostics, "ADOCMD002", AsciiDocMarkdownDiagnosticSeverity.Info, AsciiDocMarkdownConversionOutcome.Omitted,
                    "unset-document-attribute", "Unset document attribute has no YAML front-matter equivalent.", attribute);
                continue;
            }
            values[attribute.Name] = attribute.Value.Length == 0 ? true : (object)attribute.Value;
        }

        if (values.Count > 0) target.FrontMatter(values);
    }

    private static void AddList(
        MarkdownDoc target,
        AsciiDocListBlock source,
        AsciiDocDocumentAttributes attributes,
        AsciiDocToMarkdownOptions options,
        List<AsciiDocMarkdownConversionDiagnostic> diagnostics) {
        if (source.Kind == AsciiDocListKind.Ordered) {
            var list = new OrderedListBlock();
            for (int index = 0; index < source.Items.Count; index++) {
                AsciiDocListItem sourceItem = source.Items[index];
                var item = new ListItem(AsciiDocInlineToMarkdownConverter.Convert(sourceItem.Inlines, attributes, options, diagnostics, source));
                item.Level = Math.Max(0, sourceItem.Depth - 1);
                AddAttachedBlocks(item, sourceItem, attributes, options, diagnostics);
                list.Items.Add(item);
            }
            ApplyMetadata(list, source);
            target.Add(list);
        } else {
            var list = new UnorderedListBlock();
            for (int index = 0; index < source.Items.Count; index++) {
                AsciiDocListItem sourceItem = source.Items[index];
                var item = new ListItem(AsciiDocInlineToMarkdownConverter.Convert(sourceItem.Inlines, attributes, options, diagnostics, source));
                item.Level = Math.Max(0, sourceItem.Depth - 1);
                AddAttachedBlocks(item, sourceItem, attributes, options, diagnostics);
                list.Items.Add(item);
            }
            ApplyMetadata(list, source);
            target.Add(list);
        }
    }

    private static void AddAttachedBlocks(
        ListItem target,
        AsciiDocListItem source,
        AsciiDocDocumentAttributes attributes,
        AsciiDocToMarkdownOptions options,
        List<AsciiDocMarkdownConversionDiagnostic> diagnostics) {
        for (int index = 0; index < source.AttachedBlocks.Count; index++) {
            var temporary = MarkdownDoc.Create();
            AddBlock(temporary, source.AttachedBlocks[index], attributes, options, diagnostics);
            for (int childIndex = 0; childIndex < temporary.Blocks.Count; childIndex++) {
                target.NestedBlocks.Add(temporary.Blocks[childIndex]);
            }
        }
    }

    private static void AddDescriptionList(
        MarkdownDoc target,
        AsciiDocDescriptionListBlock source,
        AsciiDocDocumentAttributes attributes,
        AsciiDocToMarkdownOptions options,
        List<AsciiDocMarkdownConversionDiagnostic> diagnostics) {
        var list = new DefinitionListBlock();
        for (int index = 0; index < source.Items.Count; index++) {
            AsciiDocDescriptionListItem item = source.Items[index];
            InlineSequence term = AsciiDocInlineToMarkdownConverter.Convert(item.TermInlines, attributes, options, diagnostics, source);
            InlineSequence definition = AsciiDocInlineToMarkdownConverter.Convert(item.DescriptionInlines, attributes, options, diagnostics, source);
            list.AddEntry(new DefinitionListEntry(term, new IMarkdownBlock[] { new ParagraphBlock(definition) }));
            if (item.Depth > 1) {
                Report(diagnostics, "ADOCMD051", AsciiDocMarkdownDiagnosticSeverity.Warning, AsciiDocMarkdownConversionOutcome.Simplified,
                    "nested-description-list", "Nested description-list depth is flattened because Markdown definition lists do not share the same nesting markers.", source);
            }
        }
        ApplyMetadata(list, source);
        target.Add(list);
    }

    private static void AddAdmonition(
        MarkdownDoc target,
        AsciiDocAdmonitionBlock source,
        AsciiDocDocumentAttributes attributes,
        AsciiDocToMarkdownOptions options,
        List<AsciiDocMarkdownConversionDiagnostic> diagnostics) {
        InlineSequence body = AsciiDocInlineToMarkdownConverter.Convert(source.Inlines, attributes, options, diagnostics, source);
        var callout = new CalloutBlock(source.Kind.ToString().ToLowerInvariant(), string.Empty,
            new IMarkdownBlock[] { new ParagraphBlock(body) });
        ApplyMetadata(callout, source);
        target.Add(callout);
    }

    private static void AddDelimitedBlock(
        MarkdownDoc target,
        AsciiDocDelimitedBlock source,
        AsciiDocDocumentAttributes attributes,
        AsciiDocToMarkdownOptions options,
        List<AsciiDocMarkdownConversionDiagnostic> diagnostics) {
        string content = TrimOneTrailingLineEnding(source.Content);
        if (source.AdmonitionKind.HasValue) {
            var callout = new CalloutBlock(source.AdmonitionKind.Value.ToString().ToLowerInvariant(),
                source.BlockTitle?.Title ?? string.Empty,
                content);
            ApplyMetadata(callout, source);
            target.Add(callout);
            return;
        }
        if (source.IsStem) {
            var math = new SemanticFencedBlock(MarkdownSemanticKinds.Math, "latex", content);
            ApplyMetadata(math, source);
            target.Add(math);
            return;
        }
        switch (source.Kind) {
            case AsciiDocDelimitedBlockKind.Listing:
                string language = string.Equals(source.Style, "source", StringComparison.OrdinalIgnoreCase)
                    ? source.AttributeLists.SelectMany(static list => list.Attributes.Entries)
                        .Where(static entry => entry.Kind == AsciiDocElementAttributeKind.Positional)
                        .Skip(1).Select(static entry => entry.Value).FirstOrDefault() ?? string.Empty
                    : string.Empty;
                var code = new CodeBlock(language, content);
                ApplyMetadata(code, source);
                target.Add(code);
                break;
            case AsciiDocDelimitedBlockKind.Literal:
                var literal = new CodeBlock("text", content);
                ApplyMetadata(literal, source);
                target.Add(literal);
                break;
            case AsciiDocDelimitedBlockKind.Quote:
                var quote = new QuoteBlock(content.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n'));
                ApplyMetadata(quote, source);
                target.Add(quote);
                break;
            case AsciiDocDelimitedBlockKind.Comment:
                if (options.PreserveCommentsAsSource) {
                    target.Code("asciidoc", source.OriginalText);
                    Report(diagnostics, "ADOCMD011", AsciiDocMarkdownDiagnosticSeverity.Info, AsciiDocMarkdownConversionOutcome.SourceFallback,
                        "comment-block", "Comment block preserved as visible AsciiDoc source.", source);
                } else {
                    Report(diagnostics, "ADOCMD012", AsciiDocMarkdownDiagnosticSeverity.Info, AsciiDocMarkdownConversionOutcome.Omitted,
                        "comment-block", "Comment block omitted from Markdown output.", source);
                }
                break;
            case AsciiDocDelimitedBlockKind.Example:
            case AsciiDocDelimitedBlockKind.Sidebar:
            case AsciiDocDelimitedBlockKind.Open:
                AsciiDocParagraph? paragraph = AsciiDocDocument.Parse(content).Document.BlocksOfType<AsciiDocParagraph>().FirstOrDefault();
                var simplified = paragraph == null
                    ? new ParagraphBlock(new InlineSequence().Text(content))
                    : new ParagraphBlock(AsciiDocInlineToMarkdownConverter.Convert(paragraph.Inlines, attributes, options, diagnostics, source));
                ApplyMetadata(simplified, source);
                target.Add(simplified);
                Report(diagnostics, "ADOCMD010", AsciiDocMarkdownDiagnosticSeverity.Warning, AsciiDocMarkdownConversionOutcome.Simplified,
                    source.Kind.ToString(), "Delimited container converted to a plain Markdown paragraph.", source);
                break;
            default:
                AddSourceFallback(target, source, "ADOCMD013", source.Kind.ToString(), options, diagnostics);
                break;
        }
    }

    private static void AddMacro(
        MarkdownDoc target,
        AsciiDocBlockMacro source,
        AsciiDocToMarkdownOptions options,
        List<AsciiDocMarkdownConversionDiagnostic> diagnostics) {
        if (string.Equals(source.Name, "image", StringComparison.Ordinal) && source.Target.Length > 0) {
            string? alt = FirstAttribute(source.AttributeList);
            var image = new ImageBlock(source.Target, alt);
            ApplyMetadata(image, source);
            target.Add(image);
            return;
        }

        AddSourceFallback(target, source, "ADOCMD020", "block-macro:" + source.Name, options, diagnostics);
    }

    private static void AddComment(
        MarkdownDoc target,
        AsciiDocLineComment source,
        AsciiDocToMarkdownOptions options,
        List<AsciiDocMarkdownConversionDiagnostic> diagnostics) {
        if (options.PreserveCommentsAsSource) {
            target.Code("asciidoc", source.OriginalText);
            Report(diagnostics, "ADOCMD030", AsciiDocMarkdownDiagnosticSeverity.Info, AsciiDocMarkdownConversionOutcome.SourceFallback,
                "line-comment", "Line comment preserved as visible AsciiDoc source.", source);
        } else {
            Report(diagnostics, "ADOCMD031", AsciiDocMarkdownDiagnosticSeverity.Info, AsciiDocMarkdownConversionOutcome.Omitted,
                "line-comment", "Line comment omitted from Markdown output.", source);
        }
    }

    private static void AddSourceFallback(
        MarkdownDoc target,
        AsciiDocBlock source,
        string code,
        string feature,
        AsciiDocToMarkdownOptions options,
        List<AsciiDocMarkdownConversionDiagnostic> diagnostics) {
        if (options.PreserveUnsupportedAsSource) {
            target.Code("asciidoc", source.OriginalText);
            Report(diagnostics, code, AsciiDocMarkdownDiagnosticSeverity.Warning, AsciiDocMarkdownConversionOutcome.SourceFallback,
                feature, "No equivalent Markdown semantic node is available; original AsciiDoc source was retained in a fenced block.", source);
        } else {
            Report(diagnostics, code, AsciiDocMarkdownDiagnosticSeverity.Warning, AsciiDocMarkdownConversionOutcome.Omitted,
                feature, "No equivalent Markdown semantic node is available; source was omitted by conversion options.", source);
        }
    }

    private static void Report(
        List<AsciiDocMarkdownConversionDiagnostic> diagnostics,
        string code,
        AsciiDocMarkdownDiagnosticSeverity severity,
        AsciiDocMarkdownConversionOutcome outcome,
        string feature,
        string message,
        AsciiDocBlock source) {
        diagnostics.Add(new AsciiDocMarkdownConversionDiagnostic(code, severity, outcome, feature, message, source.Span));
    }

    private static void ApplyMetadata(MarkdownObject target, AsciiDocBlock source) {
        var roles = new List<string>();
        var values = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);
        string? id = source.BlockAnchor?.Id;
        for (int listIndex = 0; listIndex < source.AttributeLists.Count; listIndex++) {
            AsciiDocElementAttributes attributes = source.AttributeLists[listIndex].Attributes;
            id = attributes.Id ?? id;
            roles.AddRange(attributes.Roles);
            for (int entryIndex = 0; entryIndex < attributes.Entries.Count; entryIndex++) {
                AsciiDocElementAttribute entry = attributes.Entries[entryIndex];
                if (entry.Kind == AsciiDocElementAttributeKind.Named && entry.Name != null &&
                    !string.Equals(entry.Name, "id", StringComparison.OrdinalIgnoreCase) &&
                    !string.Equals(entry.Name, "role", StringComparison.OrdinalIgnoreCase)) {
                    values[entry.Name] = entry.Value;
                }
            }
        }
        if (target is TableBlock && source.BlockTitle != null) values["caption"] = source.BlockTitle.Title;
        if (id != null || roles.Count > 0 || values.Count > 0) {
            target.SetAttributes(MarkdownAttributeSet.Create(id, roles, values));
        }
        if (target is ICaptionable captionable && source.BlockTitle != null) captionable.Caption = source.BlockTitle.Title;
    }

    private static string TrimOneTrailingLineEnding(string value) {
        if (value.EndsWith("\r\n", StringComparison.Ordinal)) return value.Substring(0, value.Length - 2);
        if (value.EndsWith("\r", StringComparison.Ordinal) || value.EndsWith("\n", StringComparison.Ordinal)) return value.Substring(0, value.Length - 1);
        return value;
    }

    private static string? FirstAttribute(string value) {
        if (string.IsNullOrWhiteSpace(value)) return null;
        int comma = value.IndexOf(',');
        string first = comma < 0 ? value : value.Substring(0, comma);
        first = first.Trim();
        return first.Length == 0 ? null : first;
    }
}
