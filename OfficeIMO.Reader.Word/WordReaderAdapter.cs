using System.Globalization;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using OfficeIMO.Reader.FormatInternals;
using OfficeIMO.Word;

namespace OfficeIMO.Reader.Word;

internal static class WordReaderAdapter {
    internal static ReaderWordOptions Clone(ReaderWordOptions? source) => new ReaderWordOptions {
        IncludeFootnotes = source?.IncludeFootnotes ?? true
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
                    string markdown = RenderParagraph(paragraph, options.IncludeFootnotes, ref imageIndex);
                    foreach (string part in Split(markdown, readerOptions.MaxChars)) {
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
                            Text = paragraph.Text,
                            Markdown = part,
                            Warnings = blockIndex == 0 ? legacyWarnings : null
                        });
                        blockIndex++;
                    }
                } else if (block is WordTableSnapshot table) {
                    ReaderTable readerTable = MapTable(table, sourceName, blockIndex, tableIndex++);
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
        return WordRichMapping.Apply(snapshot, readerOptions, options, result);
    }

    private static string RenderParagraph(WordParagraphSnapshot paragraph, bool includeFootnotes, ref int imageIndex) {
        var markdown = new StringBuilder();
        int headingLevel = GetHeadingLevel(paragraph);
        if (headingLevel > 0) markdown.Append(new string('#', headingLevel)).Append(' ');
        else if (paragraph.IsListItem) markdown.Append(paragraph.IsOrderedList == true ? "1. " : "- ");
        foreach (WordRunSnapshot run in paragraph.Runs) {
            string text = run.Text ?? string.Empty;
            if (run.IsHyperlink && !string.IsNullOrWhiteSpace(run.HyperlinkUri)) text = $"[{text}]({run.HyperlinkUri})";
            if (run.Bold && text.Length > 0) text = "**" + text + "**";
            if (run.Italic && text.Length > 0) text = "*" + text + "*";
            markdown.Append(text);
            if (includeFootnotes && run.Footnote != null) {
                string note = string.Join(" ", run.Footnote.Paragraphs.Select(static item => item.Text).Where(static item => !string.IsNullOrWhiteSpace(item)));
                if (note.Length > 0) markdown.Append(" [^note: ").Append(note).Append(']');
            }
            if (run.InlineImage?.Bytes is { Length: > 0 }) {
                WordInlineImageSnapshot image = run.InlineImage;
                string id = "word-image-" + imageIndex.ToString("D4", CultureInfo.InvariantCulture);
                markdown.Append(" ![").Append(image.Description ?? image.Title ?? image.FileName ?? "image").Append("](").Append(id).Append(')');
                imageIndex++;
            }
        }
        if (markdown.Length == 0) markdown.Append(paragraph.Text ?? string.Empty);
        return markdown.ToString();
    }

    private static ReaderTable MapTable(WordTableSnapshot table, string sourceName, int blockIndex, int tableIndex) {
        int columnCount = Math.Max(1, table.ColumnCount);
        string[] columns = Enumerable.Range(1, columnCount).Select(index => "Column " + index.ToString(CultureInfo.InvariantCulture)).ToArray();
        var rows = new List<IReadOnlyList<string>>();
        foreach (WordTableRowSnapshot row in table.Rows) {
            string[] cells = new string[columnCount];
            foreach (WordTableCellSnapshot cell in row.Cells) {
                if (cell.ColumnIndex >= 0 && cell.ColumnIndex < cells.Length) cells[cell.ColumnIndex] = string.Join(" ", cell.Paragraphs.Select(static paragraph => paragraph.Text));
            }
            for (int index = 0; index < cells.Length; index++) cells[index] ??= string.Empty;
            rows.Add(cells);
        }
        bool headers = rows.Count > 0;
        if (headers) {
            columns = rows[0].Select((value, index) => string.IsNullOrWhiteSpace(value) ? "Column " + (index + 1).ToString(CultureInfo.InvariantCulture) : value).ToArray();
            rows.RemoveAt(0);
        }
        return new ReaderTable {
            Title = table.Title ?? table.Description,
            Kind = "word-table",
            Columns = columns,
            Rows = rows,
            TotalRowCount = rows.Count,
            ColumnProfiles = ReaderTableProfiler.CreateProfiles(columns, rows),
            Location = new ReaderLocation { Path = sourceName, BlockIndex = blockIndex, SourceBlockIndex = table.Order, SourceBlockKind = "table", TableIndex = tableIndex }
        };
    }

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

    private static IEnumerable<string> Split(string value, int maxChars) {
        int limit = Math.Max(256, maxChars);
        if (value.Length == 0) yield break;
        for (int offset = 0; offset < value.Length; offset += limit) yield return value.Substring(offset, Math.Min(limit, value.Length - offset));
    }

    private static IReadOnlyList<string>? BuildLegacyWarnings(WordDocument document) {
        if (document.SourceFormat != WordFileFormat.Doc) return null;
        string[] warnings = document.LegacyDocImportDiagnostics.Select(static warning => "Legacy DOC import diagnostic: " + warning)
            .Concat(document.LegacyDocUnsupportedFeatures.Select(static feature => $"Legacy DOC unsupported feature: {feature.Code} ({feature.Kind}) - {feature.Description}"))
            .Take(16).ToArray();
        return warnings.Length == 0 ? null : warnings;
    }
}
