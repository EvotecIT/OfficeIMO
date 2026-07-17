namespace OfficeIMO.OneNote.Markdown;

/// <summary>Shared text and Markdown projection over the typed offline OneNote model.</summary>
public static class OneNoteMarkdownProjection {
    /// <summary>Projects one semantic element and its descendants to plain text.</summary>
    public static string ToText(OneNoteElement element) {
        if (element == null) throw new ArgumentNullException(nameof(element));
        OneNoteMarkdownModelValidator.ValidateElement(element, new OneNoteMarkdownOptions().CloneValidated());
        var builder = new StringBuilder();
        AppendText(builder, element);
        return builder.ToString().Trim();
    }

    /// <summary>Projects one table cell to plain text.</summary>
    public static string ToText(OneNoteTableCell cell) {
        if (cell == null) throw new ArgumentNullException(nameof(cell));
        OneNoteMarkdownModelValidator.ValidateCell(cell, new OneNoteMarkdownOptions().CloneValidated());
        var builder = new StringBuilder();
        foreach (OneNoteElement element in cell.Content) AppendText(builder, element);
        return builder.ToString().Trim();
    }

    /// <summary>Projects one semantic element and its descendants to Markdown.</summary>
    public static string ToMarkdown(
        OneNoteElement element,
        Func<OneNoteBinaryElement, string?>? assetUriResolver = null) {
        if (element == null) throw new ArgumentNullException(nameof(element));
        OneNoteMarkdownOptions operation = new OneNoteMarkdownOptions { AssetUriResolver = assetUriResolver }.CloneValidated();
        OneNoteMarkdownModelValidator.ValidateElement(element, operation);
        var builder = new StringBuilder();
        AppendMarkdown(builder, element, operation.AssetUriResolver);
        return builder.ToString().TrimEnd();
    }

    /// <summary>Projects one page to plain text.</summary>
    public static string ToText(OneNotePage page) {
        if (page == null) throw new ArgumentNullException(nameof(page));
        OneNoteMarkdownModelValidator.ValidatePageContent(page, new OneNoteMarkdownOptions().CloneValidated());
        var builder = new StringBuilder();
        builder.AppendLine(PageTitle(page));
        foreach (OneNoteElement element in PageRoots(page)) AppendText(builder, element);
        return builder.ToString().Trim();
    }

    /// <summary>Projects one page to Markdown.</summary>
    public static string ToMarkdown(
        OneNotePage page,
        int headingLevel = 1,
        Func<OneNoteBinaryElement, string?>? assetUriResolver = null) {
        if (page == null) throw new ArgumentNullException(nameof(page));
        OneNoteMarkdownOptions operation = new OneNoteMarkdownOptions {
            HeadingLevel = headingLevel,
            AssetUriResolver = assetUriResolver
        }.CloneValidated();
        OneNoteMarkdownModelValidator.ValidatePageContent(page, operation);
        var builder = new StringBuilder();
        AppendPage(builder, page, operation.HeadingLevel, null, operation.AssetUriResolver);
        return builder.ToString().TrimEnd();
    }

    /// <summary>Projects a section, optionally including conflict and historical pages.</summary>
    public static string ToMarkdown(OneNoteSection section, OneNoteMarkdownOptions? options = null) {
        if (section == null) throw new ArgumentNullException(nameof(section));
        OneNoteMarkdownOptions operation = (options ?? new OneNoteMarkdownOptions()).CloneValidated();
        OneNoteMarkdownModelValidator.ValidateSection(section, operation);
        var builder = new StringBuilder();
        AppendHeading(builder, operation.HeadingLevel, string.IsNullOrWhiteSpace(section.Name) ? "OneNote section" : section.Name);
        foreach (OneNotePage page in section.Pages) {
            AppendPageWithRelated(builder, page, Math.Min(6, operation.HeadingLevel + 1), operation);
        }
        return builder.ToString().TrimEnd();
    }

    /// <summary>Projects a notebook and its section-group hierarchy.</summary>
    public static string ToMarkdown(OneNoteNotebook notebook, OneNoteMarkdownOptions? options = null) {
        if (notebook == null) throw new ArgumentNullException(nameof(notebook));
        OneNoteMarkdownOptions operation = (options ?? new OneNoteMarkdownOptions()).CloneValidated();
        OneNoteMarkdownModelValidator.ValidateNotebook(notebook, operation);
        var builder = new StringBuilder();
        AppendHeading(builder, operation.HeadingLevel, string.IsNullOrWhiteSpace(notebook.Name) ? "OneNote notebook" : notebook.Name);
        AppendHierarchy(builder, notebook.Sections, notebook.SectionGroups, Math.Min(6, operation.HeadingLevel + 1), operation);
        return builder.ToString().TrimEnd();
    }

    private static void AppendGroup(StringBuilder builder, OneNoteSectionGroup group, int headingLevel, OneNoteMarkdownOptions options) {
        AppendHeading(builder, headingLevel, string.IsNullOrWhiteSpace(group.Name) ? "Section group" : group.Name);
        AppendHierarchy(builder, group.Sections, group.SectionGroups, Math.Min(6, headingLevel + 1), options);
    }

    private static void AppendHierarchy(
        StringBuilder builder,
        IList<OneNoteSection> sections,
        IList<OneNoteSectionGroup> groups,
        int headingLevel,
        OneNoteMarkdownOptions options) {
        foreach (OneNoteNotebookHierarchyItem item in OneNoteNotebookHierarchy.Order(sections, groups)) {
            if (item.Section != null) {
                AppendSection(builder, item.Section, headingLevel, options);
            } else {
                AppendGroup(builder, item.Group!, headingLevel, options);
            }
        }
    }

    private static void AppendSection(StringBuilder builder, OneNoteSection section, int headingLevel, OneNoteMarkdownOptions options) {
        AppendHeading(builder, headingLevel, string.IsNullOrWhiteSpace(section.Name) ? "OneNote section" : section.Name);
        foreach (OneNotePage page in section.Pages) AppendPageWithRelated(builder, page, Math.Min(6, headingLevel + 1), options);
    }

    private static void AppendPageWithRelated(StringBuilder builder, OneNotePage page, int headingLevel, OneNoteMarkdownOptions options) {
        AppendPageWithRelated(builder, page, headingLevel, null, options);
    }

    private static void AppendPageWithRelated(
        StringBuilder builder,
        OneNotePage page,
        int headingLevel,
        string? prefix,
        OneNoteMarkdownOptions options) {
        AppendPage(builder, page, headingLevel, prefix, options.AssetUriResolver);
        int relatedHeadingLevel = Math.Min(6, headingLevel + 1);
        if (options.IncludeConflictPages) {
            foreach (OneNotePage conflict in page.ConflictPages) {
                AppendPageWithRelated(builder, conflict, relatedHeadingLevel, "Conflict", options);
            }
        }
        if (options.IncludeVersionHistory) {
            foreach (OneNotePage version in page.VersionHistory) {
                AppendPageWithRelated(builder, version, relatedHeadingLevel, "Version", options);
            }
        }
    }

    private static void AppendPage(
        StringBuilder builder,
        OneNotePage page,
        int headingLevel,
        string? prefix,
        Func<OneNoteBinaryElement, string?>? assetUriResolver) {
        string title = PageTitle(page);
        AppendHeading(builder, headingLevel, string.IsNullOrEmpty(prefix) ? title : prefix + ": " + title);
        foreach (OneNoteElement element in PageRoots(page)) AppendMarkdown(builder, element, assetUriResolver);
    }

    private static void AppendHeading(StringBuilder builder, int level, string title) {
        builder.Append(new string('#', Math.Max(1, Math.Min(6, level))));
        builder.Append(' ');
        builder.AppendLine(Escape(title));
        builder.AppendLine();
    }

    private static void AppendText(StringBuilder builder, OneNoteElement element) {
        if (element is OneNoteParagraph paragraph) {
            string text = OneNoteTextProjection.Normalize(string.Concat(paragraph.Runs.Select(run => run.Text)));
            if (!string.IsNullOrWhiteSpace(text)) {
                if (paragraph.List != null) builder.Append(paragraph.List.Ordered ? "1. " : "- ");
                builder.AppendLine(text);
            }
            foreach (OneNoteElement child in paragraph.Children) AppendText(builder, child);
        } else if (element is OneNoteOutline outline) {
            foreach (OneNoteElement child in outline.Children) AppendText(builder, child);
        } else if (element is OneNoteTable table) {
            foreach (OneNoteTableRow row in table.Rows) builder.AppendLine(string.Join(" | ", row.Cells.Select(CellText)));
        } else if (element is OneNoteImage image) {
            builder.AppendLine("[Image: " + OneNoteTextProjection.Normalize(image.AltText ?? image.FileName ?? "image") + "]");
        } else if (element is OneNoteEmbeddedFile file) {
            builder.AppendLine("[Embedded file: " + OneNoteTextProjection.Normalize(file.FileName ?? "attachment") + "]");
        } else if (element is OneNoteMath math && !string.IsNullOrWhiteSpace(math.Text)) {
            builder.AppendLine(OneNoteTextProjection.Normalize(math.Text));
        } else if (element is OneNoteMedia media) {
            builder.AppendLine("[Media: " + OneNoteTextProjection.Normalize(media.FileName ?? "recording") + "]");
        } else if (element is OneNoteInk) {
            builder.AppendLine("[Ink]");
        }
    }

    private static void AppendMarkdown(
        StringBuilder builder,
        OneNoteElement element,
        Func<OneNoteBinaryElement, string?>? assetUriResolver) {
        if (element is OneNoteParagraph paragraph) {
            string content = string.Concat(paragraph.Runs.Select(FormatRun));
            if (!string.IsNullOrWhiteSpace(content)) {
                if (paragraph.List != null) {
                    int level = Math.Min(OneNoteListInfo.MaxLevel, Math.Max(0, paragraph.List.Level));
                    builder.Append(new string(' ', level * 2));
                    builder.Append(paragraph.List.Ordered ? "1. " : "- ");
                }
                builder.AppendLine(content);
                builder.AppendLine();
            }
            foreach (OneNoteElement child in paragraph.Children) AppendMarkdown(builder, child, assetUriResolver);
        } else if (element is OneNoteOutline outline) {
            foreach (OneNoteElement child in outline.Children) AppendMarkdown(builder, child, assetUriResolver);
        } else if (element is OneNoteTable table) {
            AppendTable(builder, table, assetUriResolver);
        } else if (element is OneNoteImage image) {
            string label = Escape(image.AltText ?? image.FileName ?? "image");
            string? uri = ResolveAssetUri(image, assetUriResolver);
            string value = uri == null ? "[Image: " + label + "]" : "![" + label + "](" + EscapeDestination(uri) + ")";
            if (!string.IsNullOrWhiteSpace(image.Hyperlink)) value = "[" + value + "](" + EscapeDestination(image.Hyperlink!) + ")";
            builder.AppendLine(value);
            builder.AppendLine();
        } else if (element is OneNoteEmbeddedFile file) {
            AppendBinaryLink(builder, file, file.FileName ?? "attachment", assetUriResolver);
        } else if (element is OneNoteMath math) {
            string value = OneNoteTextProjection.Normalize(math.Latex ?? math.Text ?? math.MathMl);
            string fence = CreateCodeFence(value);
            builder.Append(fence).AppendLine("math");
            builder.AppendLine(value);
            builder.AppendLine(fence);
            builder.AppendLine();
        } else if (element is OneNoteMedia media) {
            AppendBinaryLink(builder, media, media.FileName ?? "recording", assetUriResolver);
        } else if (element is OneNoteInk ink) {
            AppendBinaryLink(builder, ink, "Ink", assetUriResolver);
        }
    }

    private static void AppendBinaryLink(
        StringBuilder builder,
        OneNoteBinaryElement element,
        string label,
        Func<OneNoteBinaryElement, string?>? assetUriResolver) {
        string? uri = ResolveAssetUri(element, assetUriResolver);
        builder.Append('[');
        builder.Append(Escape(label));
        if (uri == null) builder.AppendLine("]");
        else builder.AppendLine("](" + EscapeDestination(uri) + ")");
        builder.AppendLine();
    }

    private static void AppendTable(
        StringBuilder builder,
        OneNoteTable table,
        Func<OneNoteBinaryElement, string?>? assetUriResolver) {
        int columns = table.Rows.Count == 0 ? table.ColumnWidths.Count : table.Rows.Max(row => row.Cells.Count);
        if (columns == 0) return;
        builder.Append('|');
        for (int column = 0; column < columns; column++) builder.Append(" Column " + (column + 1) + " |");
        builder.AppendLine();
        builder.Append('|');
        for (int column = 0; column < columns; column++) builder.Append(" --- |");
        builder.AppendLine();
        foreach (OneNoteTableRow row in table.Rows) {
            builder.Append('|');
            for (int column = 0; column < columns; column++) {
                string value = column < row.Cells.Count ? CellMarkdown(row.Cells[column], assetUriResolver) : string.Empty;
                builder.Append(' ');
                builder.Append(value.Replace("\r", " ").Replace("\n", "<br>"));
                builder.Append(" |");
            }
            builder.AppendLine();
        }
        builder.AppendLine();
    }

    private static string CellMarkdown(OneNoteTableCell cell, Func<OneNoteBinaryElement, string?>? resolver) {
        var parts = new List<string>();
        foreach (OneNoteElement element in cell.Content) {
            string value = InlineMarkdown(element, resolver);
            if (!string.IsNullOrWhiteSpace(value)) parts.Add(value);
        }
        return string.Join("<br>", parts);
    }

    private static string InlineMarkdown(OneNoteElement element, Func<OneNoteBinaryElement, string?>? resolver) {
        if (element is OneNoteParagraph paragraph) {
            var parts = new List<string>();
            string text = string.Concat(paragraph.Runs.Select(FormatRun));
            if (!string.IsNullOrWhiteSpace(text)) parts.Add(text);
            foreach (OneNoteElement child in paragraph.Children) {
                string nested = InlineMarkdown(child, resolver);
                if (!string.IsNullOrWhiteSpace(nested)) parts.Add(nested);
            }
            return string.Join("<br>", parts);
        }
        if (element is OneNoteOutline outline) return string.Join("<br>", outline.Children.Select(child => InlineMarkdown(child, resolver)).Where(value => !string.IsNullOrWhiteSpace(value)));
        if (element is OneNoteTable table) return string.Join("; ", table.Rows.SelectMany(row => row.Cells).Select(cell => CellMarkdown(cell, resolver)));
        if (element is OneNoteImage image) {
            string label = Escape(image.AltText ?? image.FileName ?? "image");
            string? uri = ResolveAssetUri(image, resolver);
            string value = uri == null ? "[Image: " + label + "]" : "![" + label + "](" + EscapeDestination(uri) + ")";
            return string.IsNullOrWhiteSpace(image.Hyperlink) ? value : "[" + value + "](" + EscapeDestination(image.Hyperlink!) + ")";
        }
        if (element is OneNoteBinaryElement binary) {
            string label = Escape(binary is OneNoteInk ? "Ink" : binary.FileName ?? "attachment");
            string? uri = ResolveAssetUri(binary, resolver);
            return uri == null ? "[" + label + "]" : "[" + label + "](" + EscapeDestination(uri) + ")";
        }
        if (element is OneNoteMath math) return Escape(math.Latex ?? math.Text ?? math.MathMl ?? string.Empty);
        return string.Empty;
    }

    private static string CellText(OneNoteTableCell cell) {
        var builder = new StringBuilder();
        foreach (OneNoteElement element in cell.Content) AppendText(builder, element);
        return builder.ToString().Trim();
    }

    private static string FormatRun(OneNoteTextRun run) {
        string value = Escape(run.Text);
        if (run.Style.Bold == true) value = "**" + value + "**";
        if (run.Style.Italic == true) value = "*" + value + "*";
        if (run.Style.Strikethrough == true) value = "~~" + value + "~~";
        if (run.Style.IsMath == true) value = "$" + value + "$";
        if (!string.IsNullOrWhiteSpace(run.Hyperlink)) value = "[" + value + "](" + EscapeDestination(run.Hyperlink!) + ")";
        return value;
    }

    private static string? ResolveAssetUri(OneNoteBinaryElement element, Func<OneNoteBinaryElement, string?>? resolver) =>
        element.Payload == null || resolver == null ? null : resolver(element);

    private static IEnumerable<OneNoteElement> PageRoots(OneNotePage page) => page.Outlines.Cast<OneNoteElement>().Concat(page.DirectContent);
    private static string PageTitle(OneNotePage page) => string.IsNullOrWhiteSpace(page.Title) ? "Untitled page" : OneNoteTextProjection.Normalize(page.Title);
    private static string Escape(string? value) => OneNoteTextProjection.Normalize(value)
        .Replace("\\", "\\\\")
        .Replace("`", "\\`")
        .Replace("~", "\\~")
        .Replace("*", "\\*")
        .Replace("_", "\\_")
        .Replace("[", "\\[")
        .Replace("]", "\\]")
        .Replace("|", "\\|")
        .Replace("<", "&lt;")
        .Replace(">", "&gt;")
        .Replace("\r\n", "<br>")
        .Replace("\r", "<br>")
        .Replace("\n", "<br>");
    private static string EscapeDestination(string value) {
        value = OneNoteTextProjection.Normalize(value);
        var builder = new StringBuilder(value.Length);
        foreach (char character in value) {
            if (MustEncodeDestinationCharacter(character)) {
                foreach (byte item in Encoding.UTF8.GetBytes(character.ToString())) {
                    builder.Append('%').Append(item.ToString("X2"));
                }
            } else {
                builder.Append(character);
            }
        }
        return builder.ToString();
    }

    private static bool MustEncodeDestinationCharacter(char character) =>
        char.IsControl(character) ||
        char.IsWhiteSpace(character) ||
        character == '(' ||
        character == ')' ||
        character == '<' ||
        character == '>' ||
        character == '"' ||
        character == '\'' ||
        character == '\\' ||
        character == '|';

    private static string CreateCodeFence(string value) {
        int backtickRun = LongestRun(value, '`');
        int tildeRun = LongestRun(value, '~');
        char delimiter = backtickRun <= tildeRun ? '`' : '~';
        int longestRun = delimiter == '`' ? backtickRun : tildeRun;
        return new string(delimiter, Math.Max(3, longestRun + 1));
    }

    private static int LongestRun(string value, char delimiter) {
        int longest = 0;
        int current = 0;
        foreach (char character in value) {
            if (character == delimiter) {
                current++;
                if (current > longest) longest = current;
            } else {
                current = 0;
            }
        }
        return longest;
    }
}
