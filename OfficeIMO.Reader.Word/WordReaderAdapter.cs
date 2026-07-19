using System.Globalization;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using OfficeIMO.Reader.FormatInternals;
using OfficeIMO.Word;

namespace OfficeIMO.Reader.Word;

internal static class WordReaderAdapter {
    internal static ReaderWordOptions Clone(ReaderWordOptions? source) => new ReaderWordOptions {
        IncludeFootnotes = source?.IncludeFootnotes ?? true,
        IncludePageLocations = source?.IncludePageLocations ?? false,
        PageLocationOptions = source?.PageLocationOptions?.Clone()
    };

    internal static OfficeDocumentReadResult ReadDocument(string path, ReaderOptions readerOptions, ReaderWordOptions options, CancellationToken cancellationToken) {
        using WordDocument document = Load(path, readerOptions);
        return Project(document, path, readerOptions, options, cancellationToken);
    }

    internal static OfficeDocumentReadResult ReadDocument(Stream stream, string? sourceName, ReaderOptions readerOptions, ReaderWordOptions options, CancellationToken cancellationToken) {
        using WordDocument document = Load(stream, readerOptions);
        return Project(document, string.IsNullOrWhiteSpace(sourceName) ? "document.docx" : sourceName!, readerOptions, options, cancellationToken);
    }

    internal static bool ProbeEncryptedOpenXml(Stream stream, ReaderOptions options, CancellationToken cancellationToken) {
        if (string.IsNullOrEmpty(options.OpenPassword) || !stream.CanSeek) return false;
        long position = stream.Position;
        try {
            cancellationToken.ThrowIfCancellationRequested();
            using WordDocument document = Load(stream, options);
            cancellationToken.ThrowIfCancellationRequested();
            return document.OpenXmlDocument.GetAllParts().Any(static part =>
                string.Equals(part.Uri.OriginalString, "/word/document.xml",
                    StringComparison.OrdinalIgnoreCase));
        } catch (OperationCanceledException) {
            throw;
        } catch {
            return false;
        } finally {
            stream.Position = position;
        }
    }

    private static WordDocument Load(string path, ReaderOptions options) {
        var loadOptions = new WordLoadOptions { AccessMode = DocumentAccessMode.ReadOnly };
        try { return WordDocument.Load(path, loadOptions); }
        catch (Exception exception) when (!string.IsNullOrEmpty(options.OpenPassword) && exception is InvalidDataException or IOException) {
            return WordDocument.LoadEncrypted(path, options.OpenPassword!, loadOptions);
        }
    }

    private static WordDocument Load(Stream stream, ReaderOptions options) {
        var loadOptions = new WordLoadOptions { AccessMode = DocumentAccessMode.ReadOnly };
        try { return WordDocument.Load(stream, loadOptions); }
        catch (Exception exception) when (!string.IsNullOrEmpty(options.OpenPassword) && exception is InvalidDataException or IOException) {
            if (stream.CanSeek) stream.Position = 0;
            return WordDocument.LoadEncrypted(stream, options.OpenPassword!, loadOptions);
        }
    }

    private static OfficeDocumentReadResult Project(WordDocument document, string sourceName, ReaderOptions readerOptions, ReaderWordOptions options, CancellationToken cancellationToken) {
        WordDocumentSnapshot snapshot = document.CreateInspectionSnapshot();
        IReadOnlyList<WordDocumentVisualSnapshot> pageSnapshots = Array.Empty<WordDocumentVisualSnapshot>();
        if (options.IncludePageLocations) {
            cancellationToken.ThrowIfCancellationRequested();
            WordImageExportOptions layoutOptions = options.PageLocationOptions?.Clone() ?? new WordImageExportOptions();
            layoutOptions.IncludeDocumentContent = true;
            layoutOptions.PageIndex = 0;
            layoutOptions.PageCount = null;
            pageSnapshots = document.CreateVisualSnapshots(layoutOptions, cancellationToken);
        }
        IReadOnlyList<string>? legacyWarnings = BuildLegacyWarnings(document);
        var chunks = new List<ReaderChunk>();
        IReadOnlyList<OfficeDocumentAsset> assets = OpenXmlImageAssetCollector.CollectWord(
            document.OpenXmlDocument, sourceName, readerOptions, options.IncludeFootnotes, cancellationToken);
        int blockIndex = 0;
        int tableIndex = 0;
        int imageIndex = 0;
        foreach (WordSectionSnapshot section in snapshot.Sections) {
            foreach (WordBlockSnapshot block in section.Elements) {
                cancellationToken.ThrowIfCancellationRequested();
                if (block is WordParagraphSnapshot paragraph) {
                    IReadOnlyList<ParagraphRepresentation> representations = BuildParagraphRepresentations(
                        paragraph,
                        options.IncludeFootnotes,
                        ref imageIndex,
                        readerOptions.MaxChars,
                        out IReadOnlyList<string>? projectionWarnings);
                    int representationIndex = 0;
                    foreach (ParagraphRepresentation representation in representations) {
                        chunks.Add(new ReaderChunk {
                            Id = $"word:{Path.GetFileName(sourceName)}:b{blockIndex.ToString("D4", CultureInfo.InvariantCulture)}",
                            Kind = ReaderInputKind.Word,
                            Location = new ReaderLocation {
                                Path = sourceName,
                                BlockIndex = blockIndex,
                                SourceBlockIndex = block.Order,
                                SourceBlockKind = "paragraph",
                                HeadingPath = IsHeading(paragraph) ? paragraph.Text : null,
                                BlockAnchor = paragraph.BookmarkName
                            },
                            Text = representation.Text,
                            Markdown = representation.Markdown,
                            Warnings = representationIndex == 0
                                ? Combine(blockIndex == 0 ? legacyWarnings : null, projectionWarnings)
                                : null
                        });
                        blockIndex++;
                        representationIndex++;
                    }
                } else if (block is WordTableSnapshot table) {
                    ReaderTable readerTable = MapTable(
                        table,
                        sourceName,
                        blockIndex,
                        tableIndex++,
                        readerOptions.MaxTableRows);
                    string markdown = RenderTable(readerTable);
                    chunks.Add(new ReaderChunk {
                        Id = $"word:{Path.GetFileName(sourceName)}:table:{tableIndex.ToString("D4", CultureInfo.InvariantCulture)}",
                        Kind = ReaderInputKind.Word,
                        Location = readerTable.Location!,
                        Text = string.Join(Environment.NewLine, readerTable.Rows.Select(static row => string.Join("\t", row))),
                        Markdown = markdown.Length <= readerOptions.MaxChars ? markdown : markdown.Substring(0, readerOptions.MaxChars),
                        Tables = new[] { readerTable },
                        Warnings = readerTable.Truncated ? new[] { "Word table rows were truncated due to MaxTableRows." } : null
                    });
                    blockIndex++;
                }
            }
        }
        OfficeDocumentReadResult result = DocumentReaderEngine.CreateDocumentResult(
            chunks,
            ReaderInputKind.Word,
            source: null,
            capabilities: new[] { OfficeDocumentReaderBuilderWordExtensions.HandlerId },
            assets: assets);
        return WordRichMapping.Apply(snapshot, pageSnapshots, readerOptions, options, result);
    }

    private static IReadOnlyList<ParagraphRepresentation> BuildParagraphRepresentations(
        WordParagraphSnapshot paragraph,
        bool includeFootnotes,
        ref int imageIndex,
        int maxChars,
        out IReadOnlyList<string>? warnings) {
        int limit = Math.Max(256, maxChars);
        var fragments = new List<ParagraphRepresentation>();
        var warningList = new List<string>();
        string prefix = string.Empty;
        int headingLevel = GetHeadingLevel(paragraph);
        if (headingLevel > 0) prefix = new string('#', headingLevel) + " ";
        else if (paragraph.IsListItem) prefix = paragraph.IsOrderedList == true ? "1. " : "- ";

        foreach (WordRunSnapshot run in paragraph.Runs) {
            int firstRunFragmentIndex = fragments.Count;
            string runText = run.Text ?? string.Empty;
            if (runText.Length > 0) {
                int offset = 0;
                while (offset < runText.Length) {
                    int sourceLength = FindRunSourceLength(run, runText, offset, prefix.Length, limit);
                    if (sourceLength == 0) {
                        sourceLength = GetUnicodeSafeLength(runText, offset, Math.Min(limit, runText.Length - offset));
                        warningList.Add("Word paragraph contains an atomic formatted run whose Markdown exceeds MaxChars; the complete Markdown construct was preserved.");
                    }

                    string sourcePart = runText.Substring(offset, sourceLength);
                    string markdownPart = prefix + RenderRunText(run, sourcePart);
                    fragments.Add(new ParagraphRepresentation(sourcePart, markdownPart));
                    prefix = string.Empty;
                    offset += sourceLength;
                }
            } else {
                string renderedEmptyRun = RenderRunText(run, string.Empty);
                if (renderedEmptyRun.Length > 0) {
                    AddAtomicMarkdownFragment(fragments, prefix + renderedEmptyRun, limit, warningList);
                    prefix = string.Empty;
                }
            }

            if (includeFootnotes && run.Footnote != null) {
                string note = string.Join(" ", run.Footnote.Paragraphs.Select(static item => item.Text).Where(static item => !string.IsNullOrWhiteSpace(item)));
                if (note.Length > 0) {
                    AppendRunAnnotation(fragments, firstRunFragmentIndex,
                        prefix + " [^note: " + note + "]", limit, warningList);
                    prefix = string.Empty;
                }
            }
            if (run.InlineImage?.Bytes is { Length: > 0 }) {
                WordInlineImageSnapshot image = run.InlineImage;
                string id = "word-image-" + imageIndex.ToString("D4", CultureInfo.InvariantCulture);
                string imageMarkdown = " ![" + (image.Description ?? image.Title ?? image.FileName ?? "image") + "](" + id + ")";
                AppendRunAnnotation(fragments, firstRunFragmentIndex,
                    prefix + imageMarkdown, limit, warningList);
                prefix = string.Empty;
                imageIndex++;
            }
        }

        if (fragments.Count == 0) {
            string fallback = paragraph.Text ?? string.Empty;
            int offset = 0;
            while (offset < fallback.Length) {
                int length = GetUnicodeSafeLength(fallback, offset, Math.Min(limit - prefix.Length, fallback.Length - offset));
                if (length == 0) {
                    AddAtomicMarkdownFragment(fragments, prefix, limit, warningList);
                    prefix = string.Empty;
                    continue;
                }

                string part = fallback.Substring(offset, length);
                fragments.Add(new ParagraphRepresentation(part, prefix + part));
                prefix = string.Empty;
                offset += length;
            }
        }

        if (prefix.Length > 0) AddAtomicMarkdownFragment(fragments, prefix, limit, warningList);
        warnings = warningList.Count == 0 ? null : warningList.Distinct(StringComparer.Ordinal).ToArray();
        return CombineParagraphFragments(fragments, limit);
    }

    private static string RenderRunText(WordRunSnapshot run, string sourceText) {
        string markdown = sourceText;
        if (run.IsHyperlink && !string.IsNullOrWhiteSpace(run.HyperlinkUri)) markdown = $"[{markdown}]({run.HyperlinkUri})";
        if (run.Bold && markdown.Length > 0) markdown = "**" + markdown + "**";
        if (run.Italic && markdown.Length > 0) markdown = "*" + markdown + "*";
        return markdown;
    }

    private static int FindRunSourceLength(WordRunSnapshot run, string value, int offset, int markdownPrefixLength, int limit) {
        int maximum = Math.Min(limit, value.Length - offset);
        int low = 1;
        int high = maximum;
        int best = 0;
        while (low <= high) {
            int midpoint = low + ((high - low) / 2);
            int candidate = GetUnicodeSafeLength(value, offset, midpoint);
            if (candidate == 0) {
                low = midpoint + 1;
                continue;
            }
            int renderedLength = markdownPrefixLength + RenderRunText(run, value.Substring(offset, candidate)).Length;
            if (renderedLength <= limit) {
                best = candidate;
                low = midpoint + 1;
            } else {
                high = midpoint - 1;
            }
        }

        return best;
    }

    private static int GetUnicodeSafeLength(string value, int offset, int requestedLength) {
        if (requestedLength <= 0) return 0;
        int length = Math.Min(requestedLength, value.Length - offset);
        int end = offset + length;
        if (end < value.Length && char.IsHighSurrogate(value[end - 1]) && char.IsLowSurrogate(value[end])) length--;
        return length;
    }

    private static void AddAtomicMarkdownFragment(
        List<ParagraphRepresentation> fragments,
        string markdown,
        int limit,
        List<string> warnings) {
        if (markdown.Length == 0) return;
        if (markdown.Length > limit) {
            warnings.Add("Word paragraph contains an atomic Markdown annotation that exceeds MaxChars; the complete Markdown construct was preserved.");
        }
        fragments.Add(new ParagraphRepresentation(string.Empty, markdown));
    }

    private static void AppendRunAnnotation(
        List<ParagraphRepresentation> fragments,
        int firstRunFragmentIndex,
        string markdown,
        int limit,
        List<string> warnings) {
        if (markdown.Length == 0) return;
        if (fragments.Count <= firstRunFragmentIndex) {
            AddAtomicMarkdownFragment(fragments, markdown, limit, warnings);
            return;
        }

        int lastIndex = fragments.Count - 1;
        ParagraphRepresentation source = fragments[lastIndex];
        string combinedMarkdown = source.Markdown + markdown;
        if (combinedMarkdown.Length > limit) {
            warnings.Add("Word paragraph contains an atomic Markdown annotation that exceeds MaxChars when kept with its source run; the complete Markdown construct was preserved.");
        }
        fragments[lastIndex] = new ParagraphRepresentation(source.Text, combinedMarkdown);
    }

    private static IReadOnlyList<ParagraphRepresentation> CombineParagraphFragments(
        IReadOnlyList<ParagraphRepresentation> fragments,
        int limit) {
        if (fragments.Count == 0) return Array.Empty<ParagraphRepresentation>();
        var combined = new List<ParagraphRepresentation>();
        var text = new StringBuilder();
        var markdown = new StringBuilder();
        foreach (ParagraphRepresentation fragment in fragments) {
            bool exceedsLimit = text.Length > 0 || markdown.Length > 0
                ? text.Length + fragment.Text.Length > limit || markdown.Length + fragment.Markdown.Length > limit
                : false;
            if (exceedsLimit) {
                combined.Add(new ParagraphRepresentation(text.ToString(), markdown.ToString()));
                text.Clear();
                markdown.Clear();
            }

            text.Append(fragment.Text);
            markdown.Append(fragment.Markdown);
        }

        if (text.Length > 0 || markdown.Length > 0) {
            combined.Add(new ParagraphRepresentation(text.ToString(), markdown.ToString()));
        }
        return combined;
    }

    private static ReaderTable MapTable(
        WordTableSnapshot table,
        string sourceName,
        int blockIndex,
        int tableIndex,
        int maxRows) => WordTableProjection.Map(
            table,
            new ReaderLocation {
                Path = sourceName,
                BlockIndex = blockIndex,
                SourceBlockIndex = table.Order,
                SourceBlockKind = "table",
                TableIndex = tableIndex
            },
            tableIndex,
            maxRows);

    private static string RenderTable(ReaderTable table) {
        var result = new StringBuilder();
        result.Append('|').Append(string.Join(" | ", table.Columns.Select(Escape))).AppendLine(" |");
        result.Append('|').Append(string.Join(" | ", table.Columns.Select(static _ => "---"))).AppendLine(" |");
        foreach (IReadOnlyList<string> row in table.Rows) result.Append('|').Append(string.Join(" | ", row.Select(Escape))).AppendLine(" |");
        return result.ToString().TrimEnd();
    }

    private static string Escape(string value) => (value ?? string.Empty).Replace("|", "\\|").Replace("\r", " ").Replace("\n", " ");
    private static bool IsHeading(WordParagraphSnapshot paragraph) => GetHeadingLevel(paragraph) > 0;
    private static int GetHeadingLevel(WordParagraphSnapshot paragraph) {
        string style = paragraph.StyleName ?? paragraph.StyleId ?? string.Empty;
        for (int level = 1; level <= 6; level++) if (style.IndexOf("Heading " + level, StringComparison.OrdinalIgnoreCase) >= 0 || style.Equals("Heading" + level, StringComparison.OrdinalIgnoreCase)) return level;
        return 0;
    }

    private static IReadOnlyList<string>? Combine(IReadOnlyList<string>? first, IReadOnlyList<string>? second) {
        if (first == null || first.Count == 0) return second;
        if (second == null || second.Count == 0) return first;
        return first.Concat(second).ToArray();
    }

    private static IReadOnlyList<string>? BuildLegacyWarnings(WordDocument document) {
        if (document.SourceFormat != WordFileFormat.Doc) return null;
        string[] warnings = document.LegacyDocImportDiagnostics.Select(static warning => "Legacy DOC import diagnostic: " + warning)
            .Concat(document.LegacyDocUnsupportedFeatures.Select(static feature => $"Legacy DOC unsupported feature: {feature.Code} ({feature.Kind}) - {feature.Description}"))
            .Take(16).ToArray();
        return warnings.Length == 0 ? null : warnings;
    }

    private sealed class ParagraphRepresentation {
        internal ParagraphRepresentation(string text, string markdown) {
            Text = text;
            Markdown = markdown;
        }

        internal string Text { get; }
        internal string Markdown { get; }
    }
}
