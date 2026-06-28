namespace OfficeIMO.Markdown;

public sealed partial class MarkdownNativeDocument {
    /// <summary>Enumerates source-backed block fields in document order.</summary>
    public IEnumerable<MarkdownNativeBlockSourceField> EnumerateBlockSourceFields() {
        foreach (var block in DescendantBlocksAndSelf()) {
            foreach (var field in EnumerateBlockSourceFields(block)) {
                yield return field;
            }
        }
    }

    /// <summary>Enumerates source-backed block fields with the supplied field name in document order.</summary>
    public IEnumerable<MarkdownNativeBlockSourceField> EnumerateBlockSourceFields(string name) {
        if (string.IsNullOrWhiteSpace(name)) {
            yield break;
        }

        foreach (var field in EnumerateBlockSourceFields()) {
            if (string.Equals(field.Name, name, StringComparison.OrdinalIgnoreCase)) {
                yield return field;
            }
        }
    }

    /// <summary>Finds the first source-backed block field whose span contains the supplied 1-based line and column.</summary>
    public MarkdownNativeBlockSourceField? FindBlockSourceFieldAtPosition(int lineNumber, int columnNumber) {
        foreach (var field in EnumerateBlockSourceFields()) {
            if (field.SourceSpan.ContainsPosition(lineNumber, columnNumber)) {
                return field;
            }
        }

        return null;
    }

    /// <summary>Creates a non-mutating source edit that replaces a native block source field.</summary>
    public MarkdownNativeSourceEdit CreateReplaceEdit(MarkdownNativeBlockSourceField field, string replacementMarkdown) {
        if (field == null) {
            throw new ArgumentNullException(nameof(field));
        }

        return CreateReplaceEdit(field.SourceSpan, replacementMarkdown);
    }

    internal static IEnumerable<MarkdownNativeBlockSourceField> EnumerateBlockSourceFields(MarkdownNativeBlock block) {
        switch (block) {
            case MarkdownNativeParagraphBlock paragraph:
                if (paragraph.TextSourceSpan.HasValue) {
                    yield return new MarkdownNativeBlockSourceField("paragraphText", paragraph.Text, paragraph.TextSourceSpan.Value, paragraph);
                }

                break;
            case MarkdownNativeHeadingBlock heading:
                foreach (var field in EnumerateHeadingFields(heading)) {
                    yield return field;
                }

                break;
            case MarkdownNativeCodeBlock code:
                foreach (var field in EnumerateCodeFields(code)) {
                    yield return field;
                }

                break;
            case MarkdownNativeVisualBlock visual:
                foreach (var field in EnumerateVisualFields(visual)) {
                    yield return field;
                }

                break;
            case MarkdownNativeThematicBreakBlock thematicBreak:
                if (thematicBreak.MarkerSourceSpan.HasValue) {
                    yield return new MarkdownNativeBlockSourceField("marker", thematicBreak.MarkerText, thematicBreak.MarkerSourceSpan.Value, thematicBreak);
                }

                break;
            case MarkdownNativeTableBlock table:
                if (table.AlignmentRowSourceSpan.HasValue) {
                    yield return new MarkdownNativeBlockSourceField("alignmentRow", null, table.AlignmentRowSourceSpan.Value, table);
                }

                break;
            case MarkdownNativeListBlock list:
                foreach (var field in EnumerateListFields(list)) {
                    yield return field;
                }

                break;
            case MarkdownNativeImageBlock image:
                foreach (var field in EnumerateImageFields(image)) {
                    yield return field;
                }

                break;
            case MarkdownNativeDefinitionListBlock definitionList:
                foreach (var field in EnumerateDefinitionListFields(definitionList)) {
                    yield return field;
                }

                break;
            case MarkdownNativeQuoteBlock quote:
                for (var i = 0; i < quote.MarkerSourceSpans.Count; i++) {
                    yield return new MarkdownNativeBlockSourceField("quoteMarker", ">", quote.MarkerSourceSpans[i], quote, i);
                }

                if (quote.BodySourceSpan.HasValue) {
                    yield return new MarkdownNativeBlockSourceField("quoteBody", null, quote.BodySourceSpan.Value, quote);
                }

                break;
            case MarkdownNativeCalloutBlock callout:
                if (callout.KindSourceSpan.HasValue) {
                    yield return new MarkdownNativeBlockSourceField("calloutKind", callout.CalloutKind, callout.KindSourceSpan.Value, callout);
                }

                if (callout.TitleSourceSpan.HasValue) {
                    yield return new MarkdownNativeBlockSourceField("title", callout.Title, callout.TitleSourceSpan.Value, callout);
                }

                if (callout.BodySourceSpan.HasValue) {
                    yield return new MarkdownNativeBlockSourceField("calloutBody", callout.Body, callout.BodySourceSpan.Value, callout);
                }

                break;
            case MarkdownNativeDetailsBlock details:
                if (details.SummarySourceSpan.HasValue) {
                    yield return new MarkdownNativeBlockSourceField("summary", details.Summary, details.SummarySourceSpan.Value, details);
                }

                if (details.BodySourceSpan.HasValue) {
                    yield return new MarkdownNativeBlockSourceField("detailsBody", null, details.BodySourceSpan.Value, details);
                }

                break;
            case MarkdownNativeFootnoteDefinitionBlock footnote:
                if (footnote.OpeningMarkerSourceSpan.HasValue) {
                    yield return new MarkdownNativeBlockSourceField("footnoteOpeningMarker", "[^", footnote.OpeningMarkerSourceSpan.Value, footnote);
                }

                if (footnote.LabelSourceSpan.HasValue) {
                    yield return new MarkdownNativeBlockSourceField("label", footnote.Label, footnote.LabelSourceSpan.Value, footnote);
                }

                if (footnote.SeparatorMarkerSourceSpan.HasValue) {
                    yield return new MarkdownNativeBlockSourceField("footnoteSeparatorMarker", "]:", footnote.SeparatorMarkerSourceSpan.Value, footnote);
                }

                if (footnote.BodySourceSpan.HasValue) {
                    yield return new MarkdownNativeBlockSourceField("footnoteBody", footnote.Text, footnote.BodySourceSpan.Value, footnote);
                }

                break;
            case MarkdownNativeFrontMatterBlock frontMatter:
                foreach (var field in EnumerateFrontMatterFields(frontMatter)) {
                    yield return field;
                }

                break;
            case MarkdownNativeHtmlBlock html:
                foreach (var field in EnumerateHtmlFields(html)) {
                    yield return field;
                }

                break;
        }
    }

    private static IEnumerable<MarkdownNativeBlockSourceField> EnumerateHtmlFields(MarkdownNativeHtmlBlock html) {
        if (!html.IsComment) {
            if (html.OpeningTagSourceSpan.HasValue) {
                yield return new MarkdownNativeBlockSourceField("htmlOpeningTag", html.OpeningTag, html.OpeningTagSourceSpan.Value, html);
            }

            if (html.RawBodySourceSpan.HasValue) {
                yield return new MarkdownNativeBlockSourceField("htmlBody", html.Body, html.RawBodySourceSpan.Value, html);
            }

            if (html.ClosingTagSourceSpan.HasValue) {
                yield return new MarkdownNativeBlockSourceField("htmlClosingTag", html.ClosingTag, html.ClosingTagSourceSpan.Value, html);
            }

            if (html.SourceSpan.HasValue) {
                yield return new MarkdownNativeBlockSourceField("html", html.Html, html.SourceSpan.Value, html);
            }

            yield break;
        }

        if (html.OpeningMarkerSourceSpan.HasValue) {
            yield return new MarkdownNativeBlockSourceField("htmlCommentOpeningMarker", "<!--", html.OpeningMarkerSourceSpan.Value, html);
        }

        if (html.BodySourceSpan.HasValue) {
            yield return new MarkdownNativeBlockSourceField("htmlCommentBody", html.CommentBody, html.BodySourceSpan.Value, html);
        }

        if (html.ClosingMarkerSourceSpan.HasValue) {
            yield return new MarkdownNativeBlockSourceField("htmlCommentClosingMarker", "-->", html.ClosingMarkerSourceSpan.Value, html);
        }

        if (html.SourceSpan.HasValue) {
            yield return new MarkdownNativeBlockSourceField("html", html.Html, html.SourceSpan.Value, html);
        }
    }

    private static IEnumerable<MarkdownNativeBlockSourceField> EnumerateListFields(MarkdownNativeListBlock list) {
        for (var i = 0; i < list.Items.Count; i++) {
            var item = list.Items[i];
            if (item.MarkerSourceSpan.HasValue) {
                yield return new MarkdownNativeBlockSourceField("listMarker", item.MarkerText, item.MarkerSourceSpan.Value, list, i);
            }

            if (item.TaskMarkerSourceSpan.HasValue) {
                yield return new MarkdownNativeBlockSourceField("taskMarker", item.TaskMarkerText, item.TaskMarkerSourceSpan.Value, list, i);
            }
        }
    }

    private static IEnumerable<MarkdownNativeBlockSourceField> EnumerateHeadingFields(MarkdownNativeHeadingBlock heading) {
        if (heading.LevelSourceSpan.HasValue) {
            yield return new MarkdownNativeBlockSourceField(
                "level",
                heading.Level.ToString(System.Globalization.CultureInfo.InvariantCulture),
                heading.LevelSourceSpan.Value,
                heading);
        }

        if (heading.TextSourceSpan.HasValue) {
            yield return new MarkdownNativeBlockSourceField("text", heading.Text, heading.TextSourceSpan.Value, heading);
        }

        if (heading.OpeningMarkerSourceSpan.HasValue) {
            yield return new MarkdownNativeBlockSourceField("openingMarker", heading.OpeningMarkerText, heading.OpeningMarkerSourceSpan.Value, heading);
        }

        if (heading.SetextUnderlineMarkerSourceSpan.HasValue) {
            yield return new MarkdownNativeBlockSourceField("setextUnderlineMarker", heading.SetextUnderlineMarkerText, heading.SetextUnderlineMarkerSourceSpan.Value, heading);
        }

        if (heading.ClosingMarkerSourceSpan.HasValue) {
            yield return new MarkdownNativeBlockSourceField("closingMarker", heading.ClosingMarkerText, heading.ClosingMarkerSourceSpan.Value, heading);
        }
    }

    private static IEnumerable<MarkdownNativeBlockSourceField> EnumerateCodeFields(MarkdownNativeCodeBlock code) {
        if (code.OpeningFenceSourceSpan.HasValue) {
            yield return new MarkdownNativeBlockSourceField("openingFence", null, code.OpeningFenceSourceSpan.Value, code);
        }

        if (code.InfoStringSourceSpan.HasValue) {
            yield return new MarkdownNativeBlockSourceField("infoString", code.InfoString, code.InfoStringSourceSpan.Value, code);
        }

        if (code.ContentSourceSpan.HasValue) {
            yield return new MarkdownNativeBlockSourceField("content", code.Content, code.ContentSourceSpan.Value, code);
        }

        if (code.ClosingFenceSourceSpan.HasValue) {
            yield return new MarkdownNativeBlockSourceField("closingFence", null, code.ClosingFenceSourceSpan.Value, code);
        }
    }

    private static IEnumerable<MarkdownNativeBlockSourceField> EnumerateVisualFields(MarkdownNativeVisualBlock visual) {
        if (visual.OpeningFenceSourceSpan.HasValue) {
            yield return new MarkdownNativeBlockSourceField("openingFence", null, visual.OpeningFenceSourceSpan.Value, visual);
        }

        if (visual.InfoStringSourceSpan.HasValue) {
            yield return new MarkdownNativeBlockSourceField("infoString", visual.InfoString, visual.InfoStringSourceSpan.Value, visual);
        }

        if (visual.ContentSourceSpan.HasValue) {
            yield return new MarkdownNativeBlockSourceField("content", visual.Content, visual.ContentSourceSpan.Value, visual);
        }

        if (visual.ClosingFenceSourceSpan.HasValue) {
            yield return new MarkdownNativeBlockSourceField("closingFence", null, visual.ClosingFenceSourceSpan.Value, visual);
        }
    }

    private static IEnumerable<MarkdownNativeBlockSourceField> EnumerateImageFields(MarkdownNativeImageBlock image) {
        if (image.AltSourceSpan.HasValue) {
            yield return new MarkdownNativeBlockSourceField("alt", image.Alt, image.AltSourceSpan.Value, image);
        }

        if (image.SourceSourceSpan.HasValue) {
            yield return new MarkdownNativeBlockSourceField("source", image.Source, image.SourceSourceSpan.Value, image);
        }

        if (image.TitleSourceSpan.HasValue) {
            yield return new MarkdownNativeBlockSourceField("title", image.Title, image.TitleSourceSpan.Value, image);
        }

        if (image.LinkUrlSourceSpan.HasValue) {
            yield return new MarkdownNativeBlockSourceField("linkUrl", image.LinkUrl, image.LinkUrlSourceSpan.Value, image);
        }

        if (image.LinkTitleSourceSpan.HasValue) {
            yield return new MarkdownNativeBlockSourceField("linkTitle", image.LinkTitle, image.LinkTitleSourceSpan.Value, image);
        }
    }

    private static IEnumerable<MarkdownNativeBlockSourceField> EnumerateFrontMatterFields(MarkdownNativeFrontMatterBlock frontMatter) {
        if (frontMatter.OpeningFenceSourceSpan.HasValue) {
            yield return new MarkdownNativeBlockSourceField("openingFence", null, frontMatter.OpeningFenceSourceSpan.Value, frontMatter);
        }

        for (var i = 0; i < frontMatter.Entries.Count; i++) {
            var entry = frontMatter.Entries[i];
            if (entry.KeySourceSpan.HasValue) {
                yield return new MarkdownNativeBlockSourceField("frontMatterKey", entry.Key, entry.KeySourceSpan.Value, frontMatter, i);
            }

            if (entry.ValueSourceSpan.HasValue) {
                yield return new MarkdownNativeBlockSourceField("frontMatterValue", FrontMatterBlock.FormatSyntaxValue(entry.Value), entry.ValueSourceSpan.Value, frontMatter, i);
            }
        }

        if (frontMatter.ClosingFenceSourceSpan.HasValue) {
            yield return new MarkdownNativeBlockSourceField("closingFence", null, frontMatter.ClosingFenceSourceSpan.Value, frontMatter);
        }
    }

    private static IEnumerable<MarkdownNativeBlockSourceField> EnumerateDefinitionListFields(MarkdownNativeDefinitionListBlock definitionList) {
        var termIndex = 0;
        var definitionIndex = 0;
        for (var groupIndex = 0; groupIndex < definitionList.Groups.Count; groupIndex++) {
            var group = definitionList.Groups[groupIndex];
            for (var termOffset = 0; termOffset < group.Terms.Count; termOffset++) {
                var term = group.Terms[termOffset];
                if (term.SourceSpan.HasValue) {
                    yield return new MarkdownNativeBlockSourceField("definitionTerm", term.Text, term.SourceSpan.Value, definitionList, termIndex);
                }

                termIndex++;
            }

            for (var definitionOffset = 0; definitionOffset < group.Definitions.Count; definitionOffset++) {
                var definition = group.Definitions[definitionOffset];
                if (definition.SourceSpan.HasValue) {
                    yield return new MarkdownNativeBlockSourceField("definitionBody", definition.Markdown, definition.SourceSpan.Value, definitionList, definitionIndex);
                }

                definitionIndex++;
            }
        }
    }
}
