using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Xml.Linq;
using OfficeIMO;
using OfficeIMO.DocBook;
using OfficeIMO.Reader;

namespace OfficeIMO.Reader.DocBook;

internal static partial class DocBookReaderAdapter {
    internal static IEnumerable<ReaderChunk> Read(string path, ReaderOptions? readerOptions = null, ReaderDocBookOptions? docBookOptions = null, CancellationToken cancellationToken = default) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        ReaderOptions reader = readerOptions ?? new ReaderOptions(); ReaderInputLimits.EnforceFileSize(path, reader.MaxInputBytes);
        ReaderDocBookOptions adapter = ReaderDocBookOptionsCloner.Clone(docBookOptions); ApplyReaderLimit(adapter.ReadOptions, reader.MaxInputBytes);
        cancellationToken.ThrowIfCancellationRequested();
        return Build(DocBookDocument.Load(path, adapter.ReadOptions, cancellationToken), path, reader, adapter, cancellationToken).ToArray();
    }

    internal static IEnumerable<ReaderChunk> Read(Stream stream, string? sourceName = null, ReaderOptions? readerOptions = null, ReaderDocBookOptions? docBookOptions = null, CancellationToken cancellationToken = default) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        ReaderOptions reader = readerOptions ?? new ReaderOptions(); ReaderDocBookOptions adapter = ReaderDocBookOptionsCloner.Clone(docBookOptions);
        ApplyReaderLimit(adapter.ReadOptions, reader.MaxInputBytes);
        return Build(DocBookDocument.Load(stream, adapter.ReadOptions, cancellationToken), string.IsNullOrWhiteSpace(sourceName) ? "document.xml" : sourceName!, reader, adapter, cancellationToken).ToArray();
    }

    internal static IEnumerable<ReaderChunk> Read(DocBookDocument document, string sourceName = "document.xml", ReaderOptions? readerOptions = null, ReaderDocBookOptions? options = null, CancellationToken cancellationToken = default) =>
        Build(document ?? throw new ArgumentNullException(nameof(document)), sourceName, readerOptions ?? new ReaderOptions(), ReaderDocBookOptionsCloner.Clone(options), cancellationToken);

    private static DocBookProjection CreateProjection(DocBookDocument document, string sourceName, ReaderOptions reader, ReaderDocBookOptions options, bool includeChunkWarnings, CancellationToken cancellationToken) {
        DocBookConversionOptions conversionOptions = options.ConversionOptions;
        conversionOptions.MaxTableRows = Math.Min(conversionOptions.MaxTableRows, Math.Max(1, reader.MaxTableRows));
        DocBookConversionResult<OfficeDocumentModel> conversion = document.ToOfficeDocumentModel(sourceName, conversionOptions, cancellationToken);
        DocBookDiagnostic[] diagnostics = document.Validate(cancellationToken: cancellationToken).Diagnostics.Concat(conversion.Diagnostics).ToArray();
        IReadOnlyList<string>? warnings = includeChunkWarnings && options.IncludeDiagnostics
            ? diagnostics.Where(d => d.Severity != DocBookDiagnosticSeverity.Info)
                .Select(d => d.Code + ": " + d.Message).ToArray()
            : null;
        OfficeDocumentModel model = conversion.Value;
        return new DocBookProjection(model, diagnostics,
            BuildChunks(model, sourceName, reader, warnings, cancellationToken).ToArray());
    }

    private static IEnumerable<ReaderChunk> Build(DocBookDocument document, string sourceName, ReaderOptions reader, ReaderDocBookOptions options, CancellationToken cancellationToken) =>
        CreateProjection(document, sourceName, reader, options, includeChunkWarnings: true, cancellationToken).Chunks;

    private static IEnumerable<ReaderChunk> BuildChunks(OfficeDocumentModel model, string sourceName, ReaderOptions reader, IReadOnlyList<string>? warnings, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        ReaderTable[] tables = model.Tables.Select(table => MapTable(table, reader.MaxTableRows, sourceName)).ToArray();
        bool tablesAttached = false;
        bool warningsAttached = false;
        int sourceIndex = 0, emittedIndex = 0;
        foreach (OfficeDocumentModelNode root in model.Structure) {
            foreach (ReaderChunk chunk in BuildNode(root)) {
                if (!tablesAttached && tables.Length > 0) {
                    chunk.Tables = tables;
                    tablesAttached = true;
                }
                yield return chunk;
            }
        }
        if (!tablesAttached && tables.Length > 0) {
            yield return new ReaderChunk {
                Id = "docbook-tables",
                Kind = ReaderInputKind.DocBook,
                Text = string.Empty,
                Markdown = string.Empty,
                Tables = tables,
                Location = new ReaderLocation { Path = sourceName, BlockIndex = emittedIndex, SourceBlockKind = "table" },
                Diagnostics = new ReaderChunkDiagnostics { SourceKind = "docbook", TableCount = tables.Length },
                Warnings = TakeWarnings()
            };
        }
        if (!warningsAttached && warnings != null && warnings.Count > 0) {
            yield return new ReaderChunk {
                Id = "docbook-diagnostics",
                Kind = ReaderInputKind.DocBook,
                Text = string.Empty,
                Markdown = string.Empty,
                Location = new ReaderLocation { Path = sourceName, BlockIndex = emittedIndex, SourceBlockKind = "diagnostic" },
                Diagnostics = new ReaderChunkDiagnostics { SourceKind = "docbook" },
                Warnings = TakeWarnings()
            };
        }

        IReadOnlyList<string>? TakeWarnings() {
            if (warningsAttached || warnings == null) return null;
            warningsAttached = true;
            return warnings;
        }

        IEnumerable<ReaderChunk> BuildNode(OfficeDocumentModelNode node, ListMarker? listMarker = null,
            string? admonitionContext = null, OfficeDocumentModelNode? suppressedTitle = null) {
            cancellationToken.ThrowIfCancellationRequested();
            if (ReferenceEquals(node, suppressedTitle)) yield break;
            if (IsIndexTerm(node.Kind)) yield break;
            string? nestedAdmonitionContext = IsAdmonition(node.Kind) ? node.Kind : admonitionContext;
            if (node.Kind == "itemized-list" || node.Kind == "ordered-list") {
                bool ordered = node.Kind == "ordered-list";
                long ordinal = 1;
                if (ordered && node.Attributes.TryGetValue("startingnumber", out string? startingNumber) &&
                    long.TryParse(startingNumber, System.Globalization.NumberStyles.Integer,
                        System.Globalization.CultureInfo.InvariantCulture, out long parsedOrdinal) && parsedOrdinal > 0) {
                    ordinal = parsedOrdinal;
                }
                foreach (OfficeDocumentModelNode child in node.Children) {
                    if (child.Kind == "list-item") {
                        string indentation = listMarker?.ContinuationPrefix ?? string.Empty;
                        var childMarker = new ListMarker(indentation + (ordered ? ordinal + ". " : "- "));
                        foreach (ReaderChunk chunk in BuildNode(child, childMarker, nestedAdmonitionContext, suppressedTitle)) yield return chunk;
                        if (ordinal < long.MaxValue) ordinal++;
                    } else {
                        foreach (ReaderChunk chunk in BuildNode(child, listMarker, nestedAdmonitionContext, suppressedTitle)) yield return chunk;
                    }
                }
                yield break;
            }
            int currentSource = sourceIndex++;
            bool preformatted = IsPreformatted(node.Kind);
            bool ownsInlineText = OwnsInlineText(node);
            OfficeDocumentModelNode? structuralTitle = ownsInlineText ? null : GetStructuralTitle(node);
            OfficeDocumentModelNode? inlineProjectionNode = ownsInlineText ? node : structuralTitle;
            bool emittedInlineProjection = false;
            if (!preformatted && inlineProjectionNode != null &&
                TryBuildInlineFragments(inlineProjectionNode, out IReadOnlyList<InlineFragment> inlineFragments)) {
                int inlinePart = 0;
                foreach (InlineFragment fragment in inlineFragments) {
                    IReadOnlyList<string> fragmentParts = fragment.Text.Length == 0
                        ? new[] { string.Empty }
                        : DocumentReaderEngine.SplitAdapterProjection(fragment.Text, reader.MaxChars);
                    foreach (string fragmentPart in fragmentParts) {
                        string markdown = fragment.ToMarkdown(fragmentPart);
                        if (inlinePart == 0 && IsHeadingNode(node)) {
                            markdown = new string('#', Math.Min(node.Level ?? 1, 6)) + " " + markdown;
                        }
                        if (inlinePart == 0 && listMarker != null) markdown = listMarker.TakePrefix() + markdown;
                        yield return new ReaderChunk {
                            Id = inlinePart == 0 ? "docbook-" + currentSource : "docbook-" + currentSource + "-part-" + (inlinePart + 1),
                            Kind = ReaderInputKind.DocBook,
                            Text = fragmentPart,
                            Markdown = markdown,
                            ContinuesPreviousChunk = inlinePart > 0,
                            Location = new ReaderLocation { Path = sourceName, BlockIndex = emittedIndex++, SourceBlockIndex = currentSource,
                                HeadingPath = node.Location.HeadingPath,
                                SourceBlockKind = admonitionContext ?? (listMarker == null ? node.Kind : "list-item"),
                                BlockAnchor = "docbook-node-" + currentSource },
                            Diagnostics = new ReaderChunkDiagnostics { SourceKind = "docbook" },
                            Warnings = TakeWarnings()
                        };
                        inlinePart++;
                    }
                }
                emittedInlineProjection = true;
                if (ownsInlineText) yield break;
            }
            bool compoundExtension = node.Kind.StartsWith("extension:", StringComparison.Ordinal) &&
                node.Children.Any(child => child.Kind != "text");
            string projectedText = structuralTitle?.Text ?? node.Text;
            if (!emittedInlineProjection && (!compoundExtension || structuralTitle != null) && !string.IsNullOrWhiteSpace(projectedText) &&
                node.Kind != "metadata" && node.Kind != "author") {
                IReadOnlyList<string> parts = DocumentReaderEngine.SplitAdapterProjection(projectedText, reader.MaxChars);
                string codeFence = preformatted ? CreateCodeFence(projectedText) : string.Empty;
                for (int part = 0; part < parts.Count; part++) {
                    string markdown;
                    if (preformatted) {
                        markdown = parts.Count == 1 ? codeFence + "\n" + parts[part] + "\n" + codeFence
                            : (part == 0 ? codeFence + "\n" : string.Empty) + parts[part] +
                              (part == parts.Count - 1 ? "\n" + codeFence : string.Empty);
                    } else {
                        markdown = part == 0 && IsHeadingNode(node)
                            ? new string('#', Math.Min(node.Level ?? 1, 6)) + " " + parts[part]
                            : parts[part];
                    }
                    if (part == 0 && listMarker != null) markdown = listMarker.TakePrefix() + markdown;
                    yield return new ReaderChunk {
                        Id = parts.Count == 1 ? "docbook-" + currentSource : "docbook-" + currentSource + "-part-" + (part + 1),
                        Kind = ReaderInputKind.DocBook, Text = parts[part], Markdown = markdown,
                        ContinuesPreviousChunk = part > 0,
                        Location = new ReaderLocation { Path = sourceName, BlockIndex = emittedIndex++, SourceBlockIndex = currentSource,
                            HeadingPath = node.Location.HeadingPath,
                            SourceBlockKind = admonitionContext ?? (listMarker == null ? node.Kind : "list-item"),
                            BlockAnchor = "docbook-node-" + currentSource },
                        Diagnostics = new ReaderChunkDiagnostics { SourceKind = "docbook" }, Warnings = TakeWarnings()
                    };
                }
            }
            if (node.Kind == "list-item" && listMarker != null && !listMarker.Applied &&
                BeginsWithNestedList(node)) {
                yield return new ReaderChunk {
                    Id = "docbook-" + currentSource,
                    Kind = ReaderInputKind.DocBook,
                    Text = string.Empty,
                    Markdown = listMarker.TakePrefix().TrimEnd(),
                    Location = new ReaderLocation { Path = sourceName, BlockIndex = emittedIndex++, SourceBlockIndex = currentSource,
                        HeadingPath = node.Location.HeadingPath,
                        SourceBlockKind = admonitionContext ?? "list-item",
                        BlockAnchor = "docbook-node-" + currentSource },
                    Diagnostics = new ReaderChunkDiagnostics { SourceKind = "docbook" },
                    Warnings = TakeWarnings()
                };
            }
            if (!ownsInlineText) {
                OfficeDocumentModelNode? childSuppressedTitle = structuralTitle ?? suppressedTitle;
                foreach (OfficeDocumentModelNode child in node.Children) {
                    foreach (ReaderChunk chunk in BuildNode(child, listMarker, nestedAdmonitionContext, childSuppressedTitle)) yield return chunk;
                }
            }
        }
    }

    private static bool IsAdmonition(string kind) =>
        kind == "note" || kind == "tip" || kind == "important" || kind == "caution" || kind == "warning";

    private static bool IsPreformatted(string kind) => kind == "code" || kind == "screen";

    private static bool IsIndexTerm(string kind) =>
        string.Equals(kind, "index-term", StringComparison.OrdinalIgnoreCase);

    private static bool BeginsWithNestedList(OfficeDocumentModelNode node) {
        OfficeDocumentModelNode? firstContent = node.Children.FirstOrDefault(child =>
            !IsIndexTerm(child.Kind) && (!string.Equals(child.Kind, "text", StringComparison.OrdinalIgnoreCase) ||
            !string.IsNullOrWhiteSpace(child.Text)));
        return firstContent != null &&
            (firstContent.Kind == "itemized-list" || firstContent.Kind == "ordered-list");
    }

    private static bool OwnsInlineText(OfficeDocumentModelNode node) =>
        node.Kind == "paragraph" || node.Kind == "code" || node.Kind == "screen" ||
        node.Kind == "title" || node.Kind == "subtitle" || node.Kind == "author" ||
        node.Kind == "link" || node.Kind == "cross-reference" || node.Kind == "table-cell" || node.Kind == "caption" ||
        (node.Kind.StartsWith("extension:", StringComparison.Ordinal) && node.Children.Count == 0 && !string.IsNullOrWhiteSpace(node.Text)) ||
        (!string.IsNullOrWhiteSpace(node.Text) && node.Children.Count > 0 &&
         node.Children.All(child => child.Kind == "text") &&
         string.Equals(string.Concat(node.Children.Select(child => child.Text)), node.Text, StringComparison.Ordinal));

    private static OfficeDocumentModelNode? GetStructuralTitle(OfficeDocumentModelNode node) {
        if (!IsStructuralTitleOwner(node)) return null;
        OfficeDocumentModelNode? title = node.Children.FirstOrDefault(child => child.Kind == "title");
        return title ?? node.Children
            .Where(child => child.Kind == "metadata")
            .SelectMany(child => child.Children)
            .FirstOrDefault(child => child.Kind == "title");
    }

    private static bool IsHeadingNode(OfficeDocumentModelNode node) =>
        node.Kind == "title" || IsStructuralTitleOwner(node);

    private static bool IsStructuralTitleOwner(OfficeDocumentModelNode node) =>
        node.Kind == "section" || node.Kind == "table" || node.Kind == "figure" || IsBookComponentKind(node.Kind);

    private static bool IsBookComponentKind(string kind) {
        const string prefix = "extension:";
        if (!kind.StartsWith(prefix, StringComparison.Ordinal)) return false;
        XName name;
        try {
            name = XName.Get(kind.Substring(prefix.Length));
        } catch (ArgumentException) {
            return false;
        }
        if (name.NamespaceName.Length > 0 &&
            !string.Equals(name.NamespaceName, DocBookSchemaProfiles.DocBook52.NamespaceUri, StringComparison.Ordinal)) return false;
        switch (name.LocalName) {
            case "chapter":
            case "appendix":
            case "article":
            case "bibliography":
            case "glossary":
            case "index":
            case "part":
            case "preface":
            case "reference":
            case "setindex":
                return true;
            default:
                return false;
        }
    }

    private static string CreateCodeFence(string text) {
        int backticks = LongestRun(text, '`');
        int tildes = LongestRun(text, '~');
        char marker = backticks <= tildes ? '`' : '~';
        int longest = marker == '`' ? backticks : tildes;
        return new string(marker, Math.Max(3, longest + 1));

        static int LongestRun(string value, char markerCharacter) {
            int longestRun = 0;
            int currentRun = 0;
            foreach (char character in value) {
                if (character == markerCharacter) {
                    currentRun++;
                    if (currentRun > longestRun) longestRun = currentRun;
                } else {
                    currentRun = 0;
                }
            }
            return longestRun;
        }
    }

    private sealed class ListMarker {
        internal ListMarker(string prefix) {
            Prefix = prefix;
            ContinuationPrefix = new string(' ', prefix.Length);
        }
        internal string Prefix { get; }
        internal string ContinuationPrefix { get; }
        internal bool Applied { get; set; }

        internal string TakePrefix() {
            if (Applied) return ContinuationPrefix;
            Applied = true;
            return Prefix;
        }
    }

    private sealed class DocBookProjection {
        internal DocBookProjection(OfficeDocumentModel model, IReadOnlyList<DocBookDiagnostic> diagnostics, ReaderChunk[] chunks) {
            Model = model;
            Diagnostics = diagnostics;
            Chunks = chunks;
        }
        internal OfficeDocumentModel Model { get; }
        internal IReadOnlyList<DocBookDiagnostic> Diagnostics { get; }
        internal ReaderChunk[] Chunks { get; }
    }

    private static void ApplyReaderLimit(DocBookReadOptions options, long? maxBytes) {
        if (maxBytes.HasValue) options.MaxInputBytes = Math.Min(options.MaxInputBytes, maxBytes.Value);
    }
    private static ReaderTable MapTable(OfficeDocumentModelTable table, int maxRows, string sourceName) {
        int rowLimit = Math.Max(1, maxRows);
        IReadOnlyList<IReadOnlyList<string>> rows = table.Rows.Take(rowLimit).ToArray();
        return new ReaderTable {
            Title = table.Title,
            Kind = table.Kind,
            Summary = table.Summary,
            PayloadHash = table.PayloadHash,
            Location = MapLocation(table.Location, sourceName, "table"),
            Columns = table.Columns,
            ColumnProfiles = ReaderTableProfiler.CreateProfiles(table.Columns, rows),
            Rows = rows,
            TotalRowCount = Math.Max(table.TotalRowCount, table.Rows.Count),
            Truncated = table.Truncated || rows.Count < table.Rows.Count
        };
    }
    private static ReaderLocation MapLocation(OfficeDocumentModelLocation? location, string sourceName, string? defaultSourceBlockKind = null) => new ReaderLocation {
        Path = location?.Path ?? sourceName,
        BlockIndex = location?.BlockIndex,
        SourceBlockIndex = location?.SourceBlockIndex,
        StartLine = location?.StartLine,
        EndLine = location?.EndLine,
        NormalizedStartLine = location?.NormalizedStartLine,
        NormalizedEndLine = location?.NormalizedEndLine,
        HeadingPath = location?.HeadingPath,
        HeadingSlug = location?.HeadingSlug,
        SourceBlockKind = location?.SourceBlockKind ?? defaultSourceBlockKind,
        BlockAnchor = location?.BlockAnchor,
        Sheet = location?.Sheet,
        A1Range = location?.A1Range,
        Slide = location?.Slide,
        Page = location?.Page,
        TableIndex = location?.TableIndex
    };
}
