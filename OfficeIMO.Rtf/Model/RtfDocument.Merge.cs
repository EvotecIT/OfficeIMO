namespace OfficeIMO.Rtf;

/// <content>Provides native semantic document append with resource remapping and explicit degradation.</content>
public sealed partial class RtfDocument {
    /// <summary>
    /// Appends an independent semantic clone of another document, remapping fonts, colors, and revision authors.
    /// Styles and modern list bindings are flattened when their resource identities cannot be imported safely.
    /// </summary>
    public RtfDocumentMergeResult AppendDocument(RtfDocument source) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        if (ReferenceEquals(source, this)) source = source.Clone();
        RtfDocument imported = source.Clone();
        var report = new RtfConversionReport();
        Dictionary<int, int> fontMap = ImportFonts(imported);
        Dictionary<int, int> colorMap = ImportColors(imported);
        Dictionary<int, int> revisionAuthorMap = ImportRevisionAuthors(imported);
        var remappedNotes = new HashSet<RtfNote>();
        int flattenedStyles = 0;
        int flattenedLists = 0;

        foreach (IRtfBlock block in imported.Blocks) {
            RemapMergedBlock(block, fontMap, colorMap, revisionAuthorMap, remappedNotes, ref flattenedStyles, ref flattenedLists);
        }

        foreach (RtfNote note in imported.Notes) {
            RemapMergedNote(note, fontMap, colorMap, revisionAuthorMap, remappedNotes, ref flattenedStyles, ref flattenedLists);
            if (!_notes.Contains(note)) _notes.Add(note);
        }

        int appended = imported.Blocks.Count;
        foreach (IRtfBlock block in imported.Blocks.ToArray()) InsertBlock(_blocks.Count, block);

        if (flattenedStyles > 0) {
            report.Add(RtfConversionSeverity.Warning, "RtfMergeStylesFlattened",
                "Style bindings were flattened to the direct formatting already present on appended content.",
                RtfConversionAction.Flattened, feature: "Styles", count: flattenedStyles);
        }
        if (flattenedLists > 0) {
            report.Add(RtfConversionSeverity.Warning, "RtfMergeListsFlattened",
                "Modern list resource bindings were removed while visible list fallback content was retained where available.",
                RtfConversionAction.Flattened, feature: "Lists", count: flattenedLists);
        }
        if (imported.HeaderFooters.Count > 0) {
            report.Add(RtfConversionSeverity.Warning, "RtfMergeHeaderFootersOmitted",
                "Source headers and footers were not appended because they belong to section-level document layout.",
                RtfConversionAction.Omitted, feature: "HeaderFooter", count: imported.HeaderFooters.Count);
        }
        int omittedSectionLayouts = imported.Sections.Count(section => section.HasAnyLayoutValue);
        if (omittedSectionLayouts > 0) {
            report.Add(RtfConversionSeverity.Warning, "RtfMergeSectionLayoutOmitted",
                "Source section layout, page setup, and column settings were not appended because document append flattens content into the destination layout.",
                RtfConversionAction.Omitted, feature: "SectionLayout", count: omittedSectionLayouts);
        }

        return new RtfDocumentMergeResult(this, appended, report);
    }

    private Dictionary<int, int> ImportFonts(RtfDocument source) {
        var map = new Dictionary<int, int>();
        foreach (RtfFont font in source.Fonts) {
            RtfFont? existing = _fonts.FirstOrDefault(item =>
                string.Equals(item.Name, font.Name, StringComparison.OrdinalIgnoreCase) &&
                item.Charset == font.Charset && item.CodePage == font.CodePage);
            if (existing == null) {
                int id = _fonts.Count == 0 ? 0 : _fonts.Max(item => item.Id) + 1;
                existing = new RtfFont(id, font.Name) {
                    Family = font.Family,
                    Charset = font.Charset,
                    Pitch = font.Pitch,
                    CodePage = font.CodePage,
                    Bias = font.Bias,
                    AlternateName = font.AlternateName,
                    Panose = font.Panose,
                    NonTaggedName = font.NonTaggedName,
                    Embedding = font.Embedding
                };
                _fonts.Add(existing);
            }
            map[font.Id] = existing.Id;
        }
        return map;
    }

    private Dictionary<int, int> ImportColors(RtfDocument source) {
        var map = new Dictionary<int, int>();
        for (int index = 0; index < source.Colors.Count; index++) {
            RtfColor color = source.Colors[index];
            int existingIndex = _colors.FindIndex(item => item.Red == color.Red && item.Green == color.Green && item.Blue == color.Blue &&
                item.ThemeColor == color.ThemeColor && item.Tint == color.Tint && item.Shade == color.Shade);
            if (existingIndex < 0) {
                _colors.Add(new RtfColor(color.Red, color.Green, color.Blue) {
                    ThemeColor = color.ThemeColor,
                    Tint = color.Tint,
                    Shade = color.Shade
                });
                existingIndex = _colors.Count - 1;
            }
            map[index + 1] = existingIndex + 1;
        }
        return map;
    }

    private Dictionary<int, int> ImportRevisionAuthors(RtfDocument source) {
        var map = new Dictionary<int, int>();
        for (int index = 0; index < source.RevisionAuthors.Count; index++) {
            RtfRevisionAuthor author = source.RevisionAuthors[index];
            int destinationIndex = _revisionAuthors.FindIndex(item => string.Equals(item.Name, author.Name, StringComparison.Ordinal));
            if (destinationIndex < 0) {
                _revisionAuthors.Add(new RtfRevisionAuthor(author.Name));
                destinationIndex = _revisionAuthors.Count - 1;
            }
            map[index] = destinationIndex;
        }
        return map;
    }

    private static void RemapMergedBlock(IRtfBlock block, Dictionary<int, int> fonts, Dictionary<int, int> colors,
        Dictionary<int, int> revisionAuthors, ISet<RtfNote> remappedNotes, ref int flattenedStyles, ref int flattenedLists) {
        if (block is RtfParagraph paragraph) {
            RemapMergedParagraph(paragraph, fonts, colors, revisionAuthors, remappedNotes, ref flattenedStyles, ref flattenedLists);
        } else if (block is RtfTable table) {
            foreach (RtfTableRow row in table.Rows) {
                row.BackgroundColorIndex = MapIndex(row.BackgroundColorIndex, colors);
                row.ShadingForegroundColorIndex = MapIndex(row.ShadingForegroundColorIndex, colors);
                RemapBorder(row.TopBorder, colors);
                RemapBorder(row.LeftBorder, colors);
                RemapBorder(row.BottomBorder, colors);
                RemapBorder(row.RightBorder, colors);
                RemapBorder(row.HorizontalBorder, colors);
                RemapBorder(row.VerticalBorder, colors);
                foreach (RtfTableCell cell in row.Cells) {
                    cell.BackgroundColorIndex = MapIndex(cell.BackgroundColorIndex, colors);
                    cell.ShadingForegroundColorIndex = MapIndex(cell.ShadingForegroundColorIndex, colors);
                    RemapBorder(cell.TopBorder, colors);
                    RemapBorder(cell.LeftBorder, colors);
                    RemapBorder(cell.BottomBorder, colors);
                    RemapBorder(cell.RightBorder, colors);
                    RemapBorder(cell.TopLeftToBottomRightBorder, colors);
                    RemapBorder(cell.TopRightToBottomLeftBorder, colors);
                    foreach (IRtfBlock child in cell.Blocks) RemapMergedBlock(child, fonts, colors, revisionAuthors, remappedNotes, ref flattenedStyles, ref flattenedLists);
                }
            }
        } else if (block is RtfObject rtfObject) {
            RemapMergedParagraph(rtfObject.Result, fonts, colors, revisionAuthors, remappedNotes, ref flattenedStyles, ref flattenedLists);
        } else if (block is RtfShape shape) {
            foreach (RtfParagraph shapeParagraph in shape.TextBoxParagraphs) RemapMergedParagraph(shapeParagraph, fonts, colors, revisionAuthors, remappedNotes, ref flattenedStyles, ref flattenedLists);
        }
    }

    private static void RemapMergedParagraph(RtfParagraph paragraph, Dictionary<int, int> fonts, Dictionary<int, int> colors,
        Dictionary<int, int> revisionAuthors, ISet<RtfNote> remappedNotes, ref int flattenedStyles, ref int flattenedLists) {
        if (paragraph.StyleId.HasValue) {
            paragraph.StyleId = null;
            flattenedStyles++;
        }
        if (paragraph.ListId.HasValue || paragraph.ListDefinitionId.HasValue) {
            paragraph.ListId = null;
            paragraph.ListDefinitionId = null;
            flattenedLists++;
        }
        paragraph.BackgroundColorIndex = MapIndex(paragraph.BackgroundColorIndex, colors);
        paragraph.ShadingForegroundColorIndex = MapIndex(paragraph.ShadingForegroundColorIndex, colors);
        paragraph.LegacyNumbering.FontId = MapIndex(paragraph.LegacyNumbering.FontId, fonts);
        RemapBorder(paragraph.TopBorder, colors);
        RemapBorder(paragraph.LeftBorder, colors);
        RemapBorder(paragraph.BottomBorder, colors);
        RemapBorder(paragraph.RightBorder, colors);
        if (paragraph.ListText != null) RemapMergedParagraph(paragraph.ListText, fonts, colors, revisionAuthors, remappedNotes, ref flattenedStyles, ref flattenedLists);

        foreach (IRtfInline inline in paragraph.Inlines) {
            if (inline is RtfRun run) {
                run.FontId = MapIndex(run.FontId, fonts);
                run.ForegroundColorIndex = MapIndex(run.ForegroundColorIndex, colors);
                run.HighlightColorIndex = MapIndex(run.HighlightColorIndex, colors);
                run.CharacterBackgroundColorIndex = MapIndex(run.CharacterBackgroundColorIndex, colors);
                run.CharacterShadingForegroundColorIndex = MapIndex(run.CharacterShadingForegroundColorIndex, colors);
                run.UnderlineColorIndex = MapIndex(run.UnderlineColorIndex, colors);
                run.CharacterBorder.ColorIndex = MapIndex(run.CharacterBorder.ColorIndex, colors);
                run.RevisionAuthorIndex = MapIndex(run.RevisionAuthorIndex, revisionAuthors);
                if (run.StyleId.HasValue) {
                    run.StyleId = null;
                    flattenedStyles++;
                }
                if (run.Note != null) RemapMergedNote(run.Note, fonts, colors, revisionAuthors, remappedNotes, ref flattenedStyles, ref flattenedLists);
            } else if (inline is RtfGeneratedText generated && generated.Note != null) {
                RemapMergedNote(generated.Note, fonts, colors, revisionAuthors, remappedNotes, ref flattenedStyles, ref flattenedLists);
            } else if (inline is RtfField field) {
                RemapMergedParagraph(field.Result, fonts, colors, revisionAuthors, remappedNotes, ref flattenedStyles, ref flattenedLists);
            } else if (inline is RtfObject rtfObject) {
                RemapMergedParagraph(rtfObject.Result, fonts, colors, revisionAuthors, remappedNotes, ref flattenedStyles, ref flattenedLists);
            } else if (inline is RtfShape shape) {
                foreach (RtfParagraph shapeParagraph in shape.TextBoxParagraphs) RemapMergedParagraph(shapeParagraph, fonts, colors, revisionAuthors, remappedNotes, ref flattenedStyles, ref flattenedLists);
            }
        }
    }

    private static void RemapMergedNote(RtfNote note, Dictionary<int, int> fonts, Dictionary<int, int> colors,
        Dictionary<int, int> revisionAuthors, ISet<RtfNote> remappedNotes, ref int flattenedStyles, ref int flattenedLists) {
        if (!remappedNotes.Add(note)) return;
        foreach (RtfParagraph paragraph in note.Paragraphs) {
            RemapMergedParagraph(paragraph, fonts, colors, revisionAuthors, remappedNotes, ref flattenedStyles, ref flattenedLists);
        }
    }

    private static int? MapIndex(int? value, IReadOnlyDictionary<int, int> map) =>
        value.HasValue && map.TryGetValue(value.Value, out int mapped) ? mapped : value;

    private static void RemapBorder(RtfParagraphBorder border, IReadOnlyDictionary<int, int> colors) => border.ColorIndex = MapIndex(border.ColorIndex, colors);
    private static void RemapBorder(RtfTableCellBorder border, IReadOnlyDictionary<int, int> colors) => border.ColorIndex = MapIndex(border.ColorIndex, colors);
    private static void RemapBorder(RtfTableRowBorder border, IReadOnlyDictionary<int, int> colors) => border.ColorIndex = MapIndex(border.ColorIndex, colors);
}
